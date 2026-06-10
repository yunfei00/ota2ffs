import pytest
from openpyxl import Workbook

from ota2ffs.converter import convert_excel


def _add_v2_table(ws, start_row, polarization, value):
    ws.cell(row=start_row, column=1, value="Polarization")
    ws.cell(row=start_row, column=2, value=polarization)
    ws.cell(row=start_row, column=3, value=1800)
    ws.cell(row=start_row + 1, column=1, value="Phi\\Theta")
    ws.cell(row=start_row + 1, column=2, value=0)
    ws.cell(row=start_row + 1, column=3, value=180)
    ws.cell(row=start_row + 2, column=1, value=0)
    ws.cell(row=start_row + 2, column=2, value=value)
    ws.cell(row=start_row + 2, column=3, value=value)


def test_converter_keeps_going_when_one_sheet_fails(tmp_path):
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Valid"
    _add_v2_table(ws, 1, "Theta", 20)
    _add_v2_table(ws, 6, "Phi", 6)
    _add_v2_table(ws, 11, "Total", 3)

    invalid = workbook.create_sheet("Invalid")
    invalid["A1"] = "not a known format"

    excel_path = tmp_path / "input.xlsx"
    workbook.save(excel_path)

    output_dir = tmp_path / "out"
    log_dir = tmp_path / "log"
    result = convert_excel(
        excel_path,
        output_dir,
        ["Valid", "Invalid"],
        frequency_value="9999",
        frequency_unit="GHz",
        log_dir=log_dir,
    )

    assert result.generated_count == 4
    assert "Invalid" in result.failures
    excel_output_dir = output_dir / "input"
    assert (excel_output_dir / "Valid_Rx.ffs").exists()
    assert (excel_output_dir / "Valid_Tx.ffs").exists()
    assert (excel_output_dir / "Valid_total_Rx.ffs").exists()
    assert (excel_output_dir / "Valid_total_Tx.ffs").exists()
    assert result.log_path is not None
    assert result.log_path.exists()
    assert result.log_path.parent == log_dir

    rx_lines = (excel_output_dir / "Valid_Rx.ffs").read_text(encoding="utf-8").splitlines()
    tx_lines = (excel_output_dir / "Valid_Tx.ffs").read_text(encoding="utf-8").splitlines()
    assert rx_lines[0] == "// CST Farfield Source File"
    assert tx_lines[0] == "// CST Farfield Source File"
    assert rx_lines[9] == "1"
    assert tx_lines[9] == "1"
    assert rx_lines[24] == "1800000000"
    assert tx_lines[24] == "1800000000"

    rx_line = rx_lines[27]
    tx_line = tx_lines[27]

    assert float(rx_line.split(",")[2]) == pytest.approx(10)
    assert float(tx_line.split(",")[2]) == pytest.approx(0.1)
