# a4_split_tab.py

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget,
    QFileDialog, QMessageBox, QListWidgetItem
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent
from PyPDF2 import PdfReader, PdfWriter
import os
from datetime import datetime


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

    def dropEvent(self, event):
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


class A4SplitTab(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout()

        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        # 左侧文件列表
        self.file_list = FileListWidget()
        self.file_list.setSelectionMode(QListWidget.ExtendedSelection)  # 允许选择多个文件
        left_layout.addWidget(self.file_list)

        # 右侧按钮
        self.add_button = QPushButton("添加A4 PDF文件")
        self.add_button.clicked.connect(self.add_pdf)
        right_layout.addWidget(self.add_button)

        self.delete_button = QPushButton("删除选择")
        self.delete_button.clicked.connect(self.delete_selected)
        right_layout.addWidget(self.delete_button)

        self.clear_button = QPushButton("清除列表")
        self.clear_button.clicked.connect(self.clear_list)
        right_layout.addWidget(self.clear_button)

        # 添加去重按钮
        self.remove_duplicates_button = QPushButton("清除重复项")
        self.remove_duplicates_button.clicked.connect(self.remove_duplicates)
        right_layout.addWidget(self.remove_duplicates_button)

        # 拆分按钮
        self.split_2_button = QPushButton("拆分2份 (A4→2×A5)")
        self.split_2_button.clicked.connect(lambda: self.split_pdf(2))
        right_layout.addWidget(self.split_2_button)

        self.split_3_button = QPushButton("拆分3份 (A4→3×A5)")
        self.split_3_button.clicked.connect(lambda: self.split_pdf(3))
        right_layout.addWidget(self.split_3_button)

        self.split_4_button = QPushButton("拆分4份 (A4→4×A5)")
        self.split_4_button.clicked.connect(lambda: self.split_pdf(4))
        right_layout.addWidget(self.split_4_button)

        self.split_5_button = QPushButton("拆分5份 (A4→5×A5)")
        self.split_5_button.clicked.connect(lambda: self.split_pdf(5))
        right_layout.addWidget(self.split_5_button)

        layout.addLayout(left_layout, 70)
        layout.addLayout(right_layout, 30)
        self.setLayout(layout)

    def add_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择A4 PDF文件", "", "PDF Files (*.pdf)")
        if files:
            for file in files:
                modification_time = datetime.fromtimestamp(os.path.getmtime(file)).strftime('%Y-%m-%d %H:%M:%S')
                file_name = os.path.basename(file)
                item_text = f"{file_name} (修改日期: {modification_time})"
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, file)
                self.file_list.addItem(item)
                
            # 默认选择所有文件
            self.file_list.selectAll()

    def delete_selected(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))

    def clear_list(self):
        self.file_list.clear()

    def remove_duplicates(self):
        """去除重复文件"""
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
            
        # 重新选择所有文件
        self.file_list.selectAll()

    def split_pdf(self, num_parts):
        """将选中的PDF文件的每页拆分为指定数量的A5页面"""
        # 先去除重复项
        self.remove_duplicates()
        
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请至少选择一个PDF文件")
            return

        # 处理所有选中的文件
        for item in selected_items:
            file_path = item.data(Qt.UserRole)
            if not file_path or not os.path.exists(file_path):
                QMessageBox.critical(self, "错误", f"文件不存在: {file_path}")
                continue
                
            try:
                # 读取PDF文件
                reader = PdfReader(file_path)
                total_pages = len(reader.pages)

                # 获取原文件路径和文件名信息
                file_dir = os.path.dirname(file_path)
                file_name = os.path.splitext(os.path.basename(file_path))[0]
                
                # 创建输出文件
                output_filename = f"{file_name}-{num_parts}a5.pdf"
                output_path = os.path.join(file_dir, output_filename)
                
                writer = PdfWriter()

                # 处理每一页
                for page_num in range(total_pages):
                    # 获取原始页面
                    original_page = reader.pages[page_num]
                    
                    # 获取页面尺寸
                    width = float(original_page.mediabox.width)
                    height = float(original_page.mediabox.height)
                    
                    # 根据拆分数量计算每个子页面的尺寸和位置
                    for i in range(num_parts):
                        # 为每个裁剪创建一个独立的页面副本
                        # 我们需要创建一个新的PDF来获取独立的页面对象
                        temp_writer = PdfWriter()
                        temp_writer.add_page(original_page)
                        
                        # 将页面写入临时缓冲区
                        from io import BytesIO
                        buffer = BytesIO()
                        temp_writer.write(buffer)
                        
                        # 从缓冲区创建新的reader以获取独立的页面对象
                        buffer.seek(0)
                        new_reader = PdfReader(buffer)
                        new_page = new_reader.pages[0]
                        
                        # 计算裁剪区域 (垂直分割，从上到下)
                        # 注意：PDF坐标系统是从左下角(0,0)开始的
                        left = 0
                        bottom = height * (num_parts - i - 1) / num_parts
                        right = width
                        top = height * (num_parts - i) / num_parts
                        
                        # 设置裁剪框
                        new_page.cropbox.lower_left = (left, bottom)
                        new_page.cropbox.upper_right = (right, top)
                        
                        # 添加到输出文件
                        writer.add_page(new_page)

                # 保存文件
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)

                QMessageBox.information(self, "成功", f"A4 PDF已成功拆分为{num_parts}份A5页面\n文件保存为:\n{output_path}")

            except Exception as e:
                QMessageBox.critical(self, "错误", f"拆分PDF失败:\n{str(e)}")
