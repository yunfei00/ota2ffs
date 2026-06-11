from __future__ import annotations

import os
import threading
import time
from pathlib import Path
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QListWidgetItem

from ota2ffs import app as app_module


def _qt_app() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _process_until(condition, timeout: float = 3.0) -> bool:
    app = _qt_app()
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        app.processEvents()
        if condition():
            return True
        time.sleep(0.01)
    return False


def _add_checked_sheet(window: app_module.MainWindow, sheet_name: str) -> None:
    item = QListWidgetItem(sheet_name)
    item.setData(Qt.UserRole, sheet_name)
    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
    item.setCheckState(Qt.Checked)
    window.sheet_list.addItem(item)
    window._sheet_selection_order = [sheet_name]


def test_radar_report_generation_runs_on_background_thread(monkeypatch, tmp_path):
    _qt_app()
    monkeypatch.setattr(app_module.MainWindow, "_load_last_excel", lambda self: None)
    monkeypatch.setattr(app_module.QMessageBox, "information", lambda *args, **kwargs: None)
    monkeypatch.setattr(app_module.QMessageBox, "warning", lambda *args, **kwargs: None)
    monkeypatch.setattr(app_module.QMessageBox, "critical", lambda *args, **kwargs: None)

    started = threading.Event()
    release_worker = threading.Event()
    worker_thread_ids: list[int] = []
    main_thread_id = threading.get_ident()

    def fake_generate_radar_report(
        excel_path: Path,
        output_dir: Path,
        selected_sheets: list[str],
        include_delta: bool = False,
    ) -> object:
        worker_thread_ids.append(threading.get_ident())
        started.set()
        assert selected_sheets == ["S1"]
        release_worker.wait(timeout=3)
        return SimpleNamespace(
            output_path=tmp_path / "report.xlsx",
            matrix_count=1,
            single_chart_count=0,
            compare_chart_count=1,
            delta_chart_count=0,
        )

    monkeypatch.setattr(app_module, "generate_radar_report", fake_generate_radar_report)

    window = app_module.MainWindow()
    window.excel_path = tmp_path / "input.xlsx"
    window.output_dir = tmp_path
    _add_checked_sheet(window, "S1")

    window.start_radar_report()

    assert _process_until(started.is_set)
    assert worker_thread_ids
    assert worker_thread_ids[0] != main_thread_id
    assert not window.convert_button.isEnabled()
    assert not window.radar_button.isEnabled()

    release_worker.set()

    assert _process_until(lambda: window._task_thread is None)
    assert window.convert_button.isEnabled()
    assert window.radar_button.isEnabled()
    assert window.last_output_dir == tmp_path
    window.close()
