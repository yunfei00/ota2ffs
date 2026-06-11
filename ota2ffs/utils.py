from __future__ import annotations

import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable


FieldMap = dict[tuple[float, float], float]
DEFAULT_FREQUENCY_HZ = 1_000_000_000.0
FREQUENCY_UNITS = ("MHz", "GHz", "KHz", "Hz")
FREQUENCY_UNIT_MULTIPLIERS = {
    "Hz": 1.0,
    "KHz": 1_000.0,
    "MHz": 1_000_000.0,
    "GHz": 1_000_000_000.0,
}


@dataclass(slots=True)
class FarFieldSource:
    sheet_name: str
    theta_angles: list[float]
    phi_angles: list[float]
    e_theta_db: FieldMap = field(default_factory=dict)
    e_phi_db: FieldMap = field(default_factory=dict)
    version: str = ""
    suffix: str = ""
    frequency_mhz: float | None = None

    def __post_init__(self) -> None:
        self.theta_angles = sorted_unique(self.theta_angles)
        self.phi_angles = sorted_unique(self.phi_angles)


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_NUMBER_RE = re.compile(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?")


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_text(value: Any, expected: str) -> bool:
    return cell_text(value).casefold() == expected.casefold()


def contains_text(value: Any, expected: str) -> bool:
    actual = re.sub(r"[^0-9a-z]+", "", cell_text(value).casefold())
    normalized_expected = re.sub(r"[^0-9a-z]+", "", expected.casefold())
    return bool(normalized_expected and normalized_expected in actual)


def extract_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None

    text = cell_text(value).replace(",", "")
    if not text:
        return None
    match = _NUMBER_RE.search(text)
    if not match:
        return None
    try:
        number = float(match.group(0))
    except ValueError:
        return None
    return number if math.isfinite(number) else None


def to_float(value: Any) -> float | None:
    return extract_number(value)


def normalize_angle(value: float, precision: int = 10) -> float:
    rounded = round(float(value), precision)
    if math.isclose(rounded, round(rounded), abs_tol=10 ** -precision):
        return float(round(rounded))
    return rounded


def sorted_unique(values: Iterable[float]) -> list[float]:
    unique = {normalize_angle(value) for value in values}
    return sorted(unique)


def infer_step(values: Iterable[float]) -> float | None:
    ordered = sorted_unique(values)
    diffs = [
        normalize_angle(b - a)
        for a, b in zip(ordered, ordered[1:])
        if not math.isclose(b, a, abs_tol=1e-9)
    ]
    positive_diffs = [diff for diff in diffs if diff > 0]
    if not positive_diffs:
        return None
    return min(positive_diffs)


def complete_angle_range(start: float, end: float, step: float | None) -> list[float]:
    if step is None or step <= 0:
        return []
    values: list[float] = []
    current = float(start)
    limit = float(end)
    guard = 0
    while current <= limit + 1e-7 and guard < 10000:
        values.append(normalize_angle(current))
        current += step
        guard += 1
    if not values or not math.isclose(values[-1], limit, abs_tol=1e-7):
        values.append(normalize_angle(limit))
    return sorted_unique(values)


def sanitize_filename(name: str) -> str:
    cleaned = _INVALID_FILENAME_CHARS.sub("_", name).strip().strip(".")
    return cleaned or "Sheet"


def output_path_for(source: FarFieldSource, output_dir: str | Path, mode: str) -> Path:
    mode_suffix = "Tx" if mode.upper() == "TX" else "Rx"
    filename = f"{sanitize_filename(source.sheet_name)}{source.suffix}_{mode_suffix}.ffs"
    return Path(output_dir) / filename


def db_to_linear(db_value: float | None, mode: str) -> float:
    if db_value is None:
        return 0.0
    db = -db_value if mode.upper() == "TX" else db_value
    return 10 ** (db / 20)


def format_number(value: float) -> str:
    value = normalize_angle(value)
    if math.isclose(value, round(value), abs_tol=1e-9):
        return str(int(round(value)))
    return f"{value:.10f}".rstrip("0").rstrip(".")


def format_linear(value: float) -> str:
    if math.isclose(value, 0.0, abs_tol=1e-15):
        return "0"
    return f"{value:.12g}"


def get_tool_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def get_default_output_dir() -> Path:
    return get_tool_base_dir() / "output"


def get_default_log_dir() -> Path:
    return get_tool_base_dir() / "log"


def frequency_to_hz(value: str | float | int | None, unit: str) -> float | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None
    if unit not in FREQUENCY_UNIT_MULTIPLIERS:
        raise ValueError(f"不支持的频率单位: {unit}")

    number = extract_number(text)
    if number is None:
        raise ValueError("频率必须是数字")
    return number * FREQUENCY_UNIT_MULTIPLIERS[unit]


def format_frequency_hz(value: float) -> str:
    if math.isclose(value, round(value), rel_tol=0, abs_tol=1e-6):
        return str(int(round(value)))
    return f"{value:.12g}"


def build_cst_header(frequency_hz: float) -> list[str]:
    return [
        "// CST Farfield Source File",
        "",
        "// Version:",
        "3.0",
        "",
        "// Data Type",
        "Farfield",
        "",
        "// #Frequencies",
        "1",
        "",
        "// Position",
        "0 0 0          // 单位 m",
        "",
        "// z-Axis",
        "0 0 1",
        "",
        "// x-Axis",
        "1 0 0",
        "",
        "// 对每个频点，各 4 行",
        "-1",
        "-1",
        "-1",
        format_frequency_hz(frequency_hz),
        "",
    ]
