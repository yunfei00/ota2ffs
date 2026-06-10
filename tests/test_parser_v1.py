from openpyxl import Workbook

from ota2ffs import parser_v1


def _build_v1_sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = "V1Sheet"

    theta_angles = [0, 90, 180]
    phi_angles = [-180, 0]

    ws["A1"] = "Theta"
    ws["B1"] = "Phi Angle"
    ws["B2"] = "Theta Angle"
    for index, theta in enumerate(theta_angles, start=3):
        ws.cell(row=1, column=index, value=theta)
    for row_index, phi in enumerate(phi_angles, start=3):
        ws.cell(row=row_index, column=2, value=phi)
        for col_index, theta in enumerate(theta_angles, start=3):
            ws.cell(row=row_index, column=col_index, value=-(row_index + col_index))

    ws["A30"] = "Phi"
    ws["B30"] = "Phi Angle"
    ws["B31"] = "Theta Angle"
    for index, theta in enumerate(theta_angles, start=3):
        ws.cell(row=30, column=index, value=theta)
    for row_index, phi in enumerate(phi_angles, start=32):
        ws.cell(row=row_index, column=2, value=phi)
        for col_index, theta in enumerate(theta_angles, start=3):
            ws.cell(row=row_index, column=col_index, value=-(row_index + col_index + 10))

    return ws


def test_v1_parser_detects_and_completes_angles():
    ws = _build_v1_sheet()

    assert parser_v1.is_v1_sheet(ws)
    source = parser_v1.parse_sheet(ws)

    assert source.version == "V1"
    assert source.theta_angles == [0, 90, 180]
    assert source.phi_angles == [0, 180, 360]
    assert source.e_theta_db[(180, 90)] == -8
    assert (360, 90) not in source.e_theta_db
