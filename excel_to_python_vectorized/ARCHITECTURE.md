# Architecture & Algorithm Diagrams

## End-to-End Pipeline

```mermaid
flowchart TD
    A[Excel Workbook .xlsx] --> B[Parse Workbook]
    B --> C[Classify Cells]
    C --> D{Formula or Hardcoded?}
    D -->|Formula| E[Formula Cells]
    D -->|Hardcoded| F[Hardcoded Cells]
    E --> G[Normalise Patterns]
    G --> H[Group by Pattern]
    H --> I[Vertical Groups]
    H --> J[Horizontal Groups]
    H --> K[Singletons]
    I --> L[Topological Sort]
    J --> L
    K --> L
    F --> M[Filter Unreferenced]
    M --> N[Generate Input Template]
    L --> O[Generate Vectorised Script]
    E --> P[Discover External Files]
    P --> Q{External Refs?}
    Q -->|Yes| R[Generate input_files_config.json]
    Q -->|No| S[Skip]
    E --> T[Analyse References]
    T --> U[Generate Analysis Report]
    O --> V[calculate.py]
    N --> W[input_template.xlsx]
    U --> X[analysis_report.xlsx]

    style V fill:#4CAF50,color:#fff
    style W fill:#2196F3,color:#fff
    style X fill:#FF9800,color:#fff
    style R fill:#9C27B0,color:#fff
```

## Pattern Normalisation

```mermaid
flowchart LR
    A["=B2*C2 at D2"] --> B["Extract refs:\nB2 (rel), C2 (rel)"]
    B --> C["Compute offsets:\nB2 → (dcol=-2, drow=0)\nC2 → (dcol=-1, drow=0)"]
    C --> D["Skeleton:\n@0 * @1"]
    D --> E["Pattern Key:\n('@0*@1',\n(('cell',None,None,('rel',-2),('rel',0)),\n ('cell',None,None,('rel',-1),('rel',0))))"]

    A2["=B3*C3 at D3"] --> B2["Extract refs:\nB3 (rel), C3 (rel)"]
    B2 --> C2["Compute offsets:\nB3 → (dcol=-2, drow=0)\nC3 → (dcol=-1, drow=0)"]
    C2 --> D2["Skeleton:\n@0 * @1"]
    D2 --> E2["Same Pattern Key ✓"]

    E --> F["Group Together"]
    E2 --> F

    style F fill:#4CAF50,color:#fff
```

## Grouping & Code Generation

```mermaid
flowchart TD
    A[Formulas with Same Pattern] --> B{Check Contiguity}
    B -->|Same column,\nconsecutive rows| C[Vertical Group]
    B -->|Same row,\nconsecutive cols| D[Horizontal Group]
    B -->|Neither| E[Individual Formulas]

    C --> F["for _r in range(start, end+1):\n    c[(sheet, col, _r)] = expr(_r)"]
    D --> G["for _ci in range(start, end+1):\n    c[(sheet, _cl(_ci), row)] = expr(_ci)"]
    E --> H["c[(sheet, col, row)] = expr"]

    style F fill:#4CAF50,color:#fff
    style G fill:#2196F3,color:#fff
    style H fill:#FF9800,color:#fff
```

## Dependency Resolution

```mermaid
flowchart TD
    A[Groups + Singles] --> B[For each item:\ncompute dependencies]
    B --> C[Build DAG:\nitem → items it depends on]
    C --> D["Topological Sort\n(Kahn's Algorithm)"]
    D --> E[Ordered list of\ngroups and singles]
    E --> F[Generate code\nin this order]

    subgraph Example
        G1["Group: D2:D4\n= B*C"] --> G3["Single: D6\n= SUM(D2:D4)"]
        G1 --> G4["Single: D7\n= AVG(D2:D4)"]
        G3 --> G5["Single: D9\n= D6*B8"]
        G5 --> G6["Single: D10\n= D6+D9"]
        G3 --> G6
    end
```

## External File Reference Flow

```mermaid
flowchart TD
    A[Formula with\nexternal ref] --> B["Detect [filename]\npattern"]
    B --> C[Generate\ninput_files_config.json]
    C --> D[User fills in\nreal file paths]
    D --> E[Generated script\nloads config at runtime]
    E --> F[Opens external\nworkbooks]
    F --> G["Reads cells into\nc[('file|sheet', col, row)]"]
    G --> H[Formula evaluation\nuses external values]

    style C fill:#9C27B0,color:#fff
    style D fill:#FF9800,color:#fff
```

## Module Structure

```mermaid
graph LR
    subgraph excel_to_python_vectorized
        MAIN[main.py\nCLI entry point]
        CONV[converter.py\nOrchestration]
        VEC[vectorizer.py\nPattern detection\n& grouping]
        GEN[code_generator.py\nPython code generation]
    end

    subgraph "Parent modules (reused)"
        EP[excel_to_python.py\nparse_workbook\nclassify_cells\nload_config]
        FC[formula_converter.py\nFormulaConverter\nhelper functions]
    end

    MAIN --> CONV
    CONV --> VEC
    CONV --> GEN
    CONV --> EP
    GEN --> FC
    VEC --> FC

    style MAIN fill:#4CAF50,color:#fff
    style CONV fill:#2196F3,color:#fff
    style VEC fill:#FF9800,color:#fff
    style GEN fill:#9C27B0,color:#fff
```

## Generated Script Data Flow

```mermaid
flowchart LR
    A[input_template.xlsx] --> B["Load workbook"]
    C[input_files_config.json] --> D["Load external\nworkbooks"]
    B --> E["Cell dict c = {}"]
    D --> E
    E --> F["Compute formulas\n(vectorised loops\n+ individual)"]
    F --> G["Write all cells\nto output workbook"]
    G --> H[result.xlsx]

    style A fill:#2196F3,color:#fff
    style H fill:#4CAF50,color:#fff
```
