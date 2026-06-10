import pytest
from openpyxl import Workbook

from ota2ffs.converter import convert_excel


def _add_v2_table(ws, start_row, polarization, value):
    ws.cell(row=start_row, column=1, value="Polarization")
    ws.cell(row=start_row, column=2, value=polarization)
    ws.cell(row=start_row, column=5, value="1800 MHz")
    ws.cell(row=start_row + 1, column=1, value="Phi\\Theta")
    ws.cell(row=start_row + 1, column=2, value=0)
    ws.cell(row=start_row + 2, column=1, value=0)
    ws.cell(row=start_row + 2, column=2, value=value)


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
    result = convert_excel(excel_path, output_dir, ["Valid", "Invalid"])

    assert result.generated_count == 4
    assert "Invalid" in result.failures
    assert (output_dir / "Valid_RX.ffs").exists()
    assert (output_dir / "Valid_TX.ffs").exists()
    assert result.log_path is not None
    assert result.log_path.exists()

    rx_line = (output_dir / "Valid_RX.ffs").read_text(encoding="utf-8").splitlines()[1]
    tx_line = (output_dir / "Valid_TX.ffs").read_text(encoding="utf-8").splitlines()[1]

    assert float(rx_line.split(",")[2]) == pytest.approx(10)
    assert float(tx_line.split(",")[2]) == pytest.approx(0.1)
