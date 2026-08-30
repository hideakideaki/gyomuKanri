from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path


def create_backup(workbook: Path, backup_dir: Path) -> Path:
    if not workbook.is_file():
        raise FileNotFoundError(workbook)
    target = make_backup_path(workbook, backup_dir)
    shutil.copy2(workbook, target)
    return target


def make_backup_path(workbook: Path, backup_dir: Path) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return backup_dir / f"{workbook.stem}_backup_{timestamp}{workbook.suffix}"
