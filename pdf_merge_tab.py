# pdf_merge_tab.py

from PyQt5.QtWidgets import QWidget, QListWidget, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QListWidgetItem
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent
from PyPDF2 import PdfReader, PdfWriter
import os
from datetime import datetime
import webbrowser
import time

# 尝试导入pyautogui，如果失败则标记为不可用
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
    pyautogui = None


class PdfManager:
    def __init__(self):
        self.pdf_files = []

    def add_pdf(self, file_path):
        if file_path.endswith('.pdf'):
            self.pdf_files.append(file_path)

    def merge_pdfs(self, pdf_files, output_path):
        pdf_writer = PdfWriter()
        for file_path in pdf_files:
            pdf_reader = PdfReader(file_path)
            for page in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page])
        with open(output_path, 'wb') as out:
            pdf_writer.write(out)


class FileListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.endswith('.pdf'):
                    modification_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                    file_name = os.path.basename(file_path)
                    item_text = f"{file_name} (修改日期: {modification_time})"
                    item = QListWidgetItem(item_text)
                    item.setData(Qt.UserRole, file_path)
                    self.addItem(item)
            event.acceptProposedAction()
        else:
            event.ignore()


class PdfMergeTab(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_manager = PdfManager()
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout()

        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        self.file_list = FileListWidget()
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        left_layout.addWidget(self.file_list)

        self.add_button = QPushButton("添加PDF")
        self.add_button.clicked.connect(self.add_pdf)
        right_layout.addWidget(self.add_button)

        self.delete_button = QPushButton("删除选择")
        self.delete_button.clicked.connect(self.delete_pdf)
        right_layout.addWidget(self.delete_button)

        self.clear_button = QPushButton("清除列表")
        self.clear_button.clicked.connect(self.clear_list)
        right_layout.addWidget(self.clear_button)

        self.remove_duplicates_button = QPushButton("清除重复项")
        self.remove_duplicates_button.clicked.connect(self.remove_duplicates)
        right_layout.addWidget(self.remove_duplicates_button)

        self.top_button = QPushButton("置顶")
        self.top_button.clicked.connect(self.move_to_top)
        right_layout.addWidget(self.top_button)

        self.up_button = QPushButton("上移")
        self.up_button.clicked.connect(self.move_up)
        right_layout.addWidget(self.up_button)

        self.down_button = QPushButton("下移")
        self.down_button.clicked.connect(self.move_down)
        right_layout.addWidget(self.down_button)

        self.bottom_button = QPushButton("置底")
        self.bottom_button.clicked.connect(self.move_to_bottom)
        right_layout.addWidget(self.bottom_button)

        self.merge_button = QPushButton("合并列表文件")
        self.merge_button.clicked.connect(self.merge_files)
        right_layout.addWidget(self.merge_button)

        self.print_button = QPushButton("打印列表文件")
        self.print_button.clicked.connect(self.print_files)
        right_layout.addWidget(self.print_button)

        layout.addLayout(left_layout)
        layout.addLayout(right_layout)
        self.setLayout(layout)

    def add_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择PDF文件", "", "PDF Files (*.pdf)")
        if files:
            for file in files:
                modification_time = datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S')
                file_name = os.path.basename(file)
                item_text = f"{file_name} (修改日期: {modification_time})"
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, file)
                self.file_list.addItem(item)

    def delete_pdf(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def clear_list(self):
        self.file_list.clear()

    def remove_duplicates(self):
        pdf_files = [self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())]
        # Remove duplicates while preserving order
        unique_files = []
        seen_files = set()
        for file in reversed(pdf_files):
            if file not in seen_files:
                unique_files.append(file)
                seen_files.add(file)
        unique_files.reverse()

        self.file_list.clear()
        for file in unique_files:
            item_text = f"{os.path.basename(file)} (修改日期: {datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S')})"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, file)
            self.file_list.addItem(item)

    def move_to_top(self):
        current_row = self.file_list.currentRow()
        if current_row > 0:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(0, item)
            self.file_list.setCurrentRow(0)

    def move_up(self):
        current_row = self.file_list.currentRow()
        if current_row > 0:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row - 1, item)
            self.file_list.setCurrentRow(current_row - 1)

    def move_down(self):
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            item = self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row + 1, item)
            self.file_list.setCurrentRow(current_row + 1)

    def move_to_bottom(self):
        current_row = self.file_list.currentRow()
        if current_row < self.file_list.count() - 1:
            item = self.file_list.takeItem(current_row)
            self.file_list.addItem(item)
            self.file_list.setCurrentRow(self.file_list.count() - 1)

    def merge_files(self):
        output_path, _ = QFileDialog.getSaveFileName(self, "保存合并后的PDF", "D:/PDF", "PDF Files (*.pdf)")
        if output_path:
            try:
                pdf_files = [self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())]
                if not pdf_files:
                    QMessageBox.warning(self, "警告", "请先添加PDF文件")
                    return

                unique_files = []
                seen_files = set()
                for file in reversed(pdf_files):
                    if file not in seen_files:
                        unique_files.append(file)
                        seen_files.add(file)
                unique_files.reverse()

                if self.pdf_manager.merge_pdfs(unique_files, output_path):
                    QMessageBox.information(self, "成功", f"PDF文件已成功合并到:\n{output_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"合并PDF失败:\n{str(e)}")

    def print_files(self):
        try:
            current_time = datetime.now().strftime("%y%m%d-%H%M%S")
            output_path = os.path.abspath(f"{current_time}.pdf")
            pdf_files = [self.file_list.item(i).data(Qt.UserRole) for i in range(self.file_list.count())]
            if not pdf_files:
                QMessageBox.warning(self, "警告", "请先添加PDF文件")
                return

            unique_files = []
            seen_files = set()
            for file in reversed(pdf_files):
                if file not in seen_files:
                    unique_files.append(file)
                    seen_files.add(file)
            unique_files.reverse()

            if not self.pdf_manager.merge_pdfs(unique_files, output_path):
                return

            if not os.path.exists(output_path):
                QMessageBox.critical(self, "错误", "生成的PDF文件不存在")
                return

            try:
                webbrowser.open_new_tab(f'file:///{output_path}')
                time.sleep(1)
                
                # 检查pyautogui是否可用
                if PYAUTOGUI_AVAILABLE and pyautogui:
                    pyautogui.hotkey('ctrl', 'p')
                    QMessageBox.information(self, "成功", "打印任务已启动")
                else:
                    QMessageBox.information(self, "提示", f"PDF文件已生成并打开:\n{output_path}\n请手动按 Ctrl+P 打印")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"打印失败:\n{str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打印准备失败:\n{str(e)}")
