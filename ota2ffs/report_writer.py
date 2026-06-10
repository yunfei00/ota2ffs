from __future__ import annotations

from datetime import datetime
from pathlib import Path


def write_conversion_log(output_dir: str | Path, lines: list[str]) -> Path:
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = Path(output_dir) / f"ota2ffs_conversion_{timestamp}.log"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
