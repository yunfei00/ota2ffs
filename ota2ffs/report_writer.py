from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .utils import get_default_log_dir


def write_conversion_log(lines: list[str], log_dir: str | Path | None = None) -> Path:
    target_dir = Path(log_dir) if log_dir is not None else get_default_log_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = target_dir / f"ota2ffs_conversion_{timestamp}.log"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
