#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path

from excel_pipeline.runner import run_pipeline_for_directory


if __name__ == "__main__":
    summary = run_pipeline_for_directory(
        excel_dir=Path("ExcelFiles"),
        output_root=Path("artifacts"),
        cache_dir=Path(".cache/normalized"),
    )
    print(json.dumps(summary, indent=2, ensure_ascii=True))
    raise SystemExit(0 if summary.get("failure_count", 0) == 0 else 1)
