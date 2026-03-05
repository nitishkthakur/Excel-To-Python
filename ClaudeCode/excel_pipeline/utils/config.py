"""Configuration management for the Excel pipeline."""

import yaml
from pathlib import Path
from typing import Any, Dict


class Config:
    """Pipeline configuration manager."""

    _instance = None
    _config: Dict[str, Any] = {}

    def __new__(cls):
        """Singleton pattern to ensure single config instance."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def load(self, config_path: str = "config.yaml") -> None:
        """
        Load configuration from YAML file.

        Args:
            config_path: Path to configuration file (default: config.yaml)

        Raises:
            FileNotFoundError: If config file doesn't exist
            yaml.YAMLError: If config file is invalid YAML
        """
        path = Path(config_path)
        if not path.exists():
            raise FileNotFoundError(f"Configuration file not found: {config_path}")

        with open(path, 'r', encoding='utf-8') as f:
            self._config = yaml.safe_load(f)

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get configuration value by dot-notation key.

        Args:
            key: Configuration key (e.g., "pipeline.input_folder")
            default: Default value if key not found

        Returns:
            Configuration value or default

        Examples:
            >>> config = Config()
            >>> config.load()
            >>> config.get("pipeline.input_folder")
            "ExcelFiles/"
        """
        keys = key.split('.')
        value = self._config

        for k in keys:
            if isinstance(value, dict):
                value = value.get(k)
                if value is None:
                    return default
            else:
                return default

        return value

    def get_all(self) -> Dict[str, Any]:
        """Get entire configuration dictionary."""
        return self._config.copy()

    @property
    def input_folder(self) -> str:
        """Get input folder path."""
        return self.get("pipeline.input_folder", "ExcelFiles/")

    @property
    def output_folder(self) -> str:
        """Get output folder path."""
        return self.get("pipeline.output_folder", "output/")

    @property
    def temp_folder(self) -> str:
        """Get temporary folder path."""
        return self.get("pipeline.temp_folder", "temp/")

    @property
    def log_level(self) -> str:
        """Get logging level."""
        return self.get("logging.level", "INFO")

    @property
    def log_file(self) -> str:
        """Get log file path."""
        return self.get("logging.file", "pipeline.log")

    @property
    def vectorization_threshold(self) -> int:
        """Get minimum cells for vectorization."""
        return self.get("performance.vectorization_threshold", 10)

    @property
    def chunk_size(self) -> int:
        """Get chunk size for large file processing."""
        return self.get("performance.chunk_size", 10000)

    @property
    def validation_tolerance(self) -> float:
        """Get numerical comparison tolerance."""
        return self.get("validation.tolerance", 1e-9)

    @property
    def check_formatting(self) -> bool:
        """Check if formatting validation is enabled."""
        return self.get("validation.check_formatting", True)

    @property
    def check_formulas(self) -> bool:
        """Check if formula validation is enabled."""
        return self.get("validation.check_formulas", True)

    @property
    def version(self) -> str:
        """Get pipeline version."""
        return self.get("version", "1.0.0")


# Global config instance
config = Config()
