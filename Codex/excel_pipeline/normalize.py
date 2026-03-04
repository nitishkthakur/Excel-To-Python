from __future__ import annotations

import os
import platform
import shutil
import subprocess
from pathlib import Path


def _convert_xls_with_excel_com(source_path: Path, target_path: Path) -> bool:
    if platform.system().lower() != "windows":
        return False

    try:
        import win32com.client  # type: ignore
    except Exception:
        return False

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        workbook = excel.Workbooks.Open(str(source_path.resolve()))
        workbook.SaveAs(str(target_path.resolve()), FileFormat=51)
        return True
    except Exception:
        return False
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def _convert_xls_with_soffice(source_path: Path, target_path: Path) -> bool:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False

    out_dir = target_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    profile_dir = out_dir / ".libreoffice_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        f"-env:UserInstallation=file://{profile_dir.resolve()}",
        "--convert-to",
        "xlsx",
        "--outdir",
        str(out_dir.resolve()),
        str(source_path.resolve()),
    ]

    env = os.environ.copy()
    env["HOME"] = str(out_dir.resolve())

    result = subprocess.run(cmd, capture_output=True, text=True, env=env)
    if result.returncode != 0:
        return False

    converted_name = f"{source_path.stem}.xlsx"
    converted_path = out_dir / converted_name
    if not converted_path.exists():
        return False

    if converted_path.resolve() != target_path.resolve():
        shutil.move(str(converted_path), str(target_path))

    return True


def convert_xls_to_xlsx(source_path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    target_path = output_dir / f"{source_path.stem}.xlsx"

    if target_path.exists() and target_path.stat().st_mtime >= source_path.stat().st_mtime:
        return target_path

    if _convert_xls_with_excel_com(source_path, target_path):
        return target_path

    if _convert_xls_with_soffice(source_path, target_path):
        return target_path

    raise RuntimeError(
        "Unable to convert .xls workbook. On Windows install pywin32 and Microsoft Excel, "
        "or install LibreOffice and make sure `soffice` is on PATH."
    )


def normalize_workbook(source_path: Path, cache_dir: Path) -> Path:
    source_path = source_path.resolve()
    cache_dir = cache_dir.resolve()
    cache_dir.mkdir(parents=True, exist_ok=True)

    if source_path.suffix.lower() == ".xlsx":
        return source_path

    if source_path.suffix.lower() == ".xls":
        return convert_xls_to_xlsx(source_path, cache_dir)

    raise ValueError(f"Unsupported workbook extension: {source_path.suffix}")
