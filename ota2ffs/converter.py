from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook

from . import parser_v1, parser_v2
from .ffs_writer import write_ffs
from .report_writer import write_conversion_log


@dataclass(slots=True)
class ConversionResult:
    generated_files: list[Path] = field(default_factory=list)
    failures: dict[str, str] = field(default_factory=dict)
    log_lines: list[str] = field(default_factory=list)
    log_path: Path | None = None

    @property
    def generated_count(self) -> int:
        return len(self.generated_files)


def convert_excel(
    excel_path: str | Path,
    output_dir: str | Path,
    sheet_names: Iterable[str] | None = None,
) -> ConversionResult:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    result = ConversionResult()
    result.log_lines.append(f"Excel 文件: {excel_path}")
    result.log_lines.append(f"输出目录: {output_dir}")

    workbook = load_workbook(excel_path, data_only=True)
    selected_sheets = list(sheet_names) if sheet_names is not None else list(workbook.sheetnames)

    for sheet_name in selected_sheets:
        if sheet_name not in workbook.sheetnames:
            message = "sheet 不存在"
            result.failures[sheet_name] = message
            result.log_lines.append(f"[失败] {sheet_name}: {message}")
            continue

        ws = workbook[sheet_name]
        try:
            if parser_v2.is_v2_sheet(ws):
                sources = parser_v2.parse_sheet(ws)
                version = "V2"
            elif parser_v1.is_v1_sheet(ws):
                sources = [parser_v1.parse_sheet(ws)]
                version = "V1"
            else:
                raise ValueError("无法识别为 V1 或 V2 格式")

            before_count = len(result.generated_files)
            for source in sources:
                result.generated_files.append(write_ffs(source, output_dir, "RX"))
                result.generated_files.append(write_ffs(source, output_dir, "TX"))

            produced = len(result.generated_files) - before_count
            result.log_lines.append(f"[成功] {sheet_name}: {version}, 生成 {produced} 个文件")
        except Exception as exc:
            message = str(exc)
            result.failures[sheet_name] = message
            result.log_lines.append(f"[失败] {sheet_name}: {message}")

    workbook.close()

    result.log_lines.append(f"生成文件总数: {result.generated_count}")
    result.log_lines.append(f"失败 sheet 数量: {len(result.failures)}")
    result.log_path = write_conversion_log(output_dir, result.log_lines)
    return result
