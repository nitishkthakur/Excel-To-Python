"""DependencyGraph: Build topological order for formula calculation."""

from typing import List, Dict, Set, Tuple, Union, Optional
from collections import deque
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


class DependencyGraph:
    """Build and analyze formula dependencies for calculation order."""

    def __init__(self):
        """Initialize DependencyGraph."""
        # Nodes: (sheet, cell) or group_id
        self.nodes: Set[Union[Tuple[str, str], int]] = set()

        # Edges: node -> set of nodes it depends on (precedents)
        self.precedents: Dict[Union[Tuple[str, str], int], Set[Union[Tuple[str, str], int]]] = {}

        # Edges: node -> set of nodes that depend on it (dependents)
        self.dependents: Dict[Union[Tuple[str, str], int], Set[Union[Tuple[str, str], int]]] = {}

        # Cell to group mapping
        self.cell_to_group: Dict[Tuple[str, str], int] = {}

        # Group metadata
        self.group_info: Dict[int, Dict] = {}

    def add_cell(self, sheet: str, cell: str, group_id: Optional[int] = None):
        """
        Add a cell node to the graph.

        Args:
            sheet: Sheet name
            cell: Cell coordinate (e.g., "A1")
            group_id: Optional group ID if cell belongs to a group
        """
        if group_id:
            # Cell belongs to a group - use group as node
            self.nodes.add(group_id)
            self.cell_to_group[(sheet, cell)] = group_id
        else:
            # Individual cell
            node = (sheet, cell)
            self.nodes.add(node)

    def add_dependency(
        self,
        dependent_sheet: str,
        dependent_cell: str,
        precedent_sheet: str,
        precedent_cell: str
    ):
        """
        Add a dependency: dependent_cell depends on precedent_cell.

        Args:
            dependent_sheet: Sheet of the dependent cell
            dependent_cell: Cell that depends on another
            precedent_sheet: Sheet of the precedent cell
            precedent_cell: Cell that is depended upon
        """
        # Get nodes (may be groups or cells)
        dependent_node = self.cell_to_group.get((dependent_sheet, dependent_cell), (dependent_sheet, dependent_cell))
        precedent_node = self.cell_to_group.get((precedent_sheet, precedent_cell), (precedent_sheet, precedent_cell))

        # Skip self-dependencies
        if dependent_node == precedent_node:
            return

        # Add to precedents
        if dependent_node not in self.precedents:
            self.precedents[dependent_node] = set()
        self.precedents[dependent_node].add(precedent_node)

        # Add to dependents
        if precedent_node not in self.dependents:
            self.dependents[precedent_node] = set()
        self.dependents[precedent_node].add(dependent_node)

    def topological_sort(self) -> List[Union[Tuple[str, str], int]]:
        """
        Perform topological sort using Kahn's algorithm.

        Returns:
            Ordered list of nodes (cells or group IDs) in calculation order
        """
        logger.info("Computing calculation order (topological sort)...")

        # Calculate in-degrees
        in_degree: Dict[Union[Tuple[str, str], int], int] = {node: 0 for node in self.nodes}

        for node in self.nodes:
            if node in self.precedents:
                in_degree[node] = len(self.precedents[node])

        # Queue of nodes with no dependencies
        queue = deque([node for node, degree in in_degree.items() if degree == 0])

        result = []

        while queue:
            node = queue.popleft()
            result.append(node)

            # Process dependents
            if node in self.dependents:
                for dependent in self.dependents[node]:
                    in_degree[dependent] -= 1
                    if in_degree[dependent] == 0:
                        queue.append(dependent)

        # Check for cycles
        if len(result) < len(self.nodes):
            remaining = [node for node in self.nodes if node not in result]
            logger.warning(f"Circular dependencies detected for {len(remaining)} nodes")
            logger.warning(f"  Remaining nodes: {remaining[:10]}...")  # Show first 10

            # Add remaining nodes in arbitrary order
            result.extend(remaining)

        logger.info(f"Calculation order determined: {len(result)} nodes")

        return result

    def detect_cycles(self) -> List[List[Union[Tuple[str, str], int]]]:
        """
        Detect circular dependencies in the graph.

        Returns:
            List of cycles (each cycle is a list of nodes)
        """
        cycles = []
        visited = set()
        rec_stack = set()

        def dfs(node, path):
            visited.add(node)
            rec_stack.add(node)
            path.append(node)

            if node in self.dependents:
                for neighbor in self.dependents[node]:
                    if neighbor not in visited:
                        if dfs(neighbor, path.copy()):
                            return True
                    elif neighbor in rec_stack:
                        # Found cycle
                        cycle_start = path.index(neighbor)
                        cycle = path[cycle_start:]
                        cycles.append(cycle)
                        return True

            rec_stack.remove(node)
            return False

        for node in self.nodes:
            if node not in visited:
                dfs(node, [])

        if cycles:
            logger.warning(f"Detected {len(cycles)} circular dependency cycles")

        return cycles

    def get_calculation_level(self, node: Union[Tuple[str, str], int]) -> int:
        """
        Get the calculation level (depth) of a node.

        Level 0 = no dependencies
        Level N = max(level of precedents) + 1

        Args:
            node: Cell tuple or group ID

        Returns:
            Calculation level
        """
        if node not in self.precedents or not self.precedents[node]:
            return 0

        max_precedent_level = 0
        for precedent in self.precedents[node]:
            level = self.get_calculation_level(precedent)
            max_precedent_level = max(max_precedent_level, level)

        return max_precedent_level + 1

    def group_by_level(self) -> Dict[int, List[Union[Tuple[str, str], int]]]:
        """
        Group nodes by calculation level for potential parallel execution.

        Returns:
            Dictionary mapping level to list of nodes at that level
        """
        levels: Dict[int, List[Union[Tuple[str, str], int]]] = {}

        for node in self.nodes:
            level = self.get_calculation_level(node)
            if level not in levels:
                levels[level] = []
            levels[level].append(node)

        logger.info(f"Grouped into {len(levels)} calculation levels")
        for level, nodes in sorted(levels.items()):
            logger.debug(f"  Level {level}: {len(nodes)} nodes")

        return levels

    def get_statistics(self) -> Dict:
        """
        Get graph statistics.

        Returns:
            Dictionary with statistics
        """
        total_nodes = len(self.nodes)
        total_edges = sum(len(deps) for deps in self.precedents.values())

        group_nodes = sum(1 for node in self.nodes if isinstance(node, int))
        cell_nodes = total_nodes - group_nodes

        cycles = self.detect_cycles()

        return {
            'total_nodes': total_nodes,
            'group_nodes': group_nodes,
            'cell_nodes': cell_nodes,
            'total_dependencies': total_edges,
            'circular_cycles': len(cycles),
            'max_calculation_level': max(self.get_calculation_level(node) for node in self.nodes) if self.nodes else 0,
        }
