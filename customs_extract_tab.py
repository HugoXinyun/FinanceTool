from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

from PyQt5.QtCore import QObject, QThread, Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QDragEnterEvent, QDropEvent
from PyQt5.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from customs_declaration_extractor import (
    CustomsDeclarationExtractor,
    ExtractionBatch,
    ExtractionResult,
    export_excel,
)


class CustomsPdfListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if any(Path(url.toLocalFile()).suffix.lower() == ".pdf" for url in event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        self.add_paths(url.toLocalFile() for url in event.mimeData().urls())
        event.acceptProposedAction()

    def add_paths(self, paths):
        existing = {
            os.path.normcase(os.path.abspath(self.item(index).data(Qt.UserRole)))
            for index in range(self.count())
        }
        for raw_path in paths:
            path = Path(raw_path)
            if path.suffix.lower() != ".pdf" or not path.is_file():
                continue
            absolute = str(path.resolve())
            normalized = os.path.normcase(absolute)
            if normalized in existing:
                continue
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            item = QListWidgetItem(f"{path.name} (修改日期: {modified})")
            item.setData(Qt.UserRole, absolute)
            self.addItem(item)
            existing.add(normalized)

    def paths(self) -> list[str]:
        return [self.item(index).data(Qt.UserRole) for index in range(self.count())]


class CustomsExtractionWorker(QObject):
    progress = pyqtSignal(int, int, str, str, str)
    finished = pyqtSignal(object, str)
    failed = pyqtSignal(str)

    def __init__(self, pdf_paths: list[str], output_path: str, enable_ocr: bool):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.output_path = output_path
        self.enable_ocr = enable_ocr

    @pyqtSlot()
    def run(self):
        try:
            extractor = CustomsDeclarationExtractor(enable_ocr=self.enable_ocr)

            def report(index: int, total: int, result: ExtractionResult):
                messages = "；".join(message for message in result.messages if message)
                self.progress.emit(
                    index,
                    total,
                    result.source_pdf,
                    result.status.value,
                    messages,
                )

            batch = extractor.extract_files(self.pdf_paths, progress_callback=report)
            if not batch.declarations:
                details = "；".join(
                    message
                    for result in batch.results
                    for message in result.messages
                )
                raise ValueError(details or "没有成功提取任何报关单")
            try:
                output = export_excel(batch, self.output_path)
            except PermissionError as exc:
                message = (
                    f"无法写入Excel文件：\n{self.output_path}\n\n"
                    "请关闭Excel中已打开的同名文件，并确认目标文件夹可写后重试。"
                )
                detail = str(exc).strip()
                if detail:
                    message += f"\n\n系统信息：{detail}"
                raise RuntimeError(message) from exc
            self.finished.emit(batch, str(output))
        except Exception as exc:
            self.failed.emit(str(exc).strip() or type(exc).__name__)


class CustomsExtractTab(QWidget):
    def __init__(self):
        super().__init__()
        self.thread: QThread | None = None
        self.worker: CustomsExtractionWorker | None = None
        self._init_ui()

    def _init_ui(self):
        root_layout = QHBoxLayout(self)
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        self.file_list = CustomsPdfListWidget()
        left_layout.addWidget(self.file_list, 3)

        self.status_log = QPlainTextEdit()
        self.status_log.setReadOnly(True)
        self.status_log.setPlaceholderText("提取状态将在这里显示")
        left_layout.addWidget(self.status_log, 1)

        self.add_button = QPushButton("添加报关单PDF")
        self.add_button.clicked.connect(self.add_pdfs)
        right_layout.addWidget(self.add_button)

        self.delete_button = QPushButton("删除选择")
        self.delete_button.clicked.connect(self.delete_selected)
        right_layout.addWidget(self.delete_button)

        self.clear_button = QPushButton("清除列表")
        self.clear_button.clicked.connect(self.clear_list)
        right_layout.addWidget(self.clear_button)

        self.ocr_checkbox = QCheckBox("扫描件启用OCR")
        self.ocr_checkbox.setChecked(True)
        right_layout.addWidget(self.ocr_checkbox)

        self.extract_button = QPushButton("提取到Excel")
        self.extract_button.clicked.connect(self.extract_to_excel)
        right_layout.addWidget(self.extract_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)
        right_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("等待开始")
        self.progress_label.setWordWrap(True)
        right_layout.addWidget(self.progress_label)
        right_layout.addStretch(1)

        root_layout.addLayout(left_layout, 72)
        root_layout.addLayout(right_layout, 28)

    def add_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择报关单PDF", "", "PDF Files (*.pdf)")
        self.file_list.add_paths(files)

    def delete_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def clear_list(self):
        self.file_list.clear()
        self.status_log.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("等待开始")

    def _set_running(self, running: bool):
        self.file_list.setAcceptDrops(not running)
        for widget in (
            self.file_list,
            self.add_button,
            self.delete_button,
            self.clear_button,
            self.extract_button,
            self.ocr_checkbox,
        ):
            widget.setEnabled(not running)

    def is_running(self) -> bool:
        return self.thread is not None

    def extract_to_excel(self):
        if self.thread is not None:
            QMessageBox.warning(self, "提取进行中", "报关单正在提取，请等待当前任务完成")
            return

        pdf_paths = self.file_list.paths()
        if not pdf_paths:
            QMessageBox.warning(self, "提示", "请先添加报关单PDF")
            return
        missing = [path for path in pdf_paths if not Path(path).is_file()]
        if missing:
            QMessageBox.warning(self, "文件不存在", "\n".join(missing))
            return

        default_path = str(Path(pdf_paths[0]).parent / "报关单提取.xlsx")
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存提取结果",
            default_path,
            "Excel Files (*.xlsx)",
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        self.status_log.clear()
        self.progress_bar.setMaximum(len(pdf_paths))
        self.progress_bar.setValue(0)
        self.progress_label.setText("准备提取")
        self._set_running(True)

        self.thread = QThread(self)
        self.worker = CustomsExtractionWorker(pdf_paths, output_path, self.ocr_checkbox.isChecked())
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.failed.connect(self._on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self._thread_finished)
        self.thread.start()

    @pyqtSlot(int, int, str, str, str)
    def _on_progress(
        self,
        index: int,
        total: int,
        source_pdf: str,
        status: str,
        messages: str,
    ):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(index)
        self.progress_label.setText(f"{index}/{total}  {source_pdf}")
        log_entry = f"{source_pdf}: {status}"
        if messages:
            log_entry += f" - {messages}"
        self.status_log.appendPlainText(log_entry)

    @pyqtSlot(object, str)
    def _on_finished(self, batch: ExtractionBatch, output_path: str):
        self.status_log.appendPlainText(f"完成: {output_path}")
        QMessageBox.information(
            self,
            "提取完成",
            f"报关单: {len(batch.declarations)} 份\n"
            f"商品明细: {len(batch.items)} 行\n"
            f"需复核: {batch.review_count} 份\n"
            f"失败/需OCR: {batch.failure_count} 份\n\n"
            f"已保存到:\n{output_path}",
        )

    @pyqtSlot(str)
    def _on_failed(self, message: str):
        self.status_log.appendPlainText(f"失败: {message}")
        QMessageBox.critical(self, "提取失败", message)

    def _thread_finished(self):
        self._set_running(False)
        self.worker = None
        self.thread = None
