from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class PatternMatrix:
    sheet_name: str
    block_name: str
    row_label: str
    col_label: str
    row_angles: list[float]
    col_angles: list[float]
    values: list[list[float]]

    def normalized_values(self) -> list[list[float]]:
        return [[-value if value else 0.0 for value in row] for row in self.values]

    def row_values(self, row_index: int) -> list[float]:
        return self.normalized_values()[row_index]

    def col_values(self, col_index: int) -> list[float]:
        return [row[col_index] if col_index < len(row) else 0.0 for row in self.normalized_values()]
