from __future__ import annotations

import os
from types import SimpleNamespace

os.environ["QT_QPA_PLATFORM"] = "offscreen"

import pytest
from PyQt5.QtGui import QCloseEvent
from PyQt5.QtWidgets import QApplication, QMessageBox

import customs_extract_tab
from customs_extract_tab import CustomsExtractTab, CustomsExtractionWorker
from main import MainWindow


@pytest.fixture(scope="module")
def application():
    app = QApplication.instance() or QApplication([])
    yield app
    app.processEvents()


def test_set_running_disables_and_restores_list_and_drop(application):
    tab = CustomsExtractTab()
    controls = (
        tab.file_list,
        tab.add_button,
        tab.delete_button,
        tab.clear_button,
        tab.extract_button,
        tab.ocr_checkbox,
    )

    tab._set_running(True)

    assert not tab.file_list.acceptDrops()
    assert all(not control.isEnabled() for control in controls)

    tab._set_running(False)

    assert tab.file_list.acceptDrops()
    assert all(control.isEnabled() for control in controls)
    tab.deleteLater()
    application.processEvents()


def test_main_window_refuses_close_while_customs_thread_runs(application, monkeypatch):
    class RunningThread:
        @staticmethod
        def isRunning() -> bool:
            return True

    warnings = []
    monkeypatch.setattr(
        QMessageBox,
        "warning",
        lambda *args: warnings.append(args) or QMessageBox.Ok,
    )
    window = MainWindow()
    window.customs_extract_tab.thread = RunningThread()
    event = QCloseEvent()

    window.closeEvent(event)

    assert not event.isAccepted()
    assert len(warnings) == 1
    assert warnings[0][1] == "报关单提取进行中"
    assert "暂时不能关闭程序" in warnings[0][2]
    window.customs_extract_tab.thread = None
    window.deleteLater()
    application.processEvents()


def test_worker_permission_error_tells_user_to_close_excel(application, monkeypatch, tmp_path):
    class FakeExtractor:
        def __init__(self, enable_ocr: bool):
            self.enable_ocr = enable_ocr

        @staticmethod
        def extract_files(paths, progress_callback=None):
            return SimpleNamespace(declarations=[object()])

    output = tmp_path / "locked.xlsx"
    monkeypatch.setattr(customs_extract_tab, "CustomsDeclarationExtractor", FakeExtractor)
    monkeypatch.setattr(
        customs_extract_tab,
        "export_excel",
        lambda batch, output_path: (_ for _ in ()).throw(PermissionError("file is locked")),
    )
    worker = CustomsExtractionWorker(["source.pdf"], str(output), enable_ocr=False)
    failures = []
    completions = []
    worker.failed.connect(failures.append)
    worker.finished.connect(lambda *args: completions.append(args))

    worker.run()

    assert completions == []
    assert len(failures) == 1
    assert str(output) in failures[0]
    assert "请关闭Excel中已打开的同名文件" in failures[0]
    assert "file is locked" in failures[0]
    worker.deleteLater()
    application.processEvents()
