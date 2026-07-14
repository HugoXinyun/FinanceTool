# main.py

from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from pdf_merge_tab import PdfMergeTab
from excel_merge_tab import ExcelMergeTab
from a4_split_tab import A4SplitTab
from customs_extract_tab import CustomsExtractTab
from PyQt5.QtWidgets import QTabWidget
from PyQt5.QtCore import Qt
import sys
from datetime import datetime


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"多瑞财务工具 - {datetime.now().strftime('%Y-%m-%d')}")
        self.setGeometry(100, 100, 800, 600)

        tab_widget = QTabWidget()
        tab_widget.addTab(PdfMergeTab(), "PDF合并")
        tab_widget.addTab(ExcelMergeTab(), "Excel合并")
        tab_widget.addTab(A4SplitTab(), "A4 PDF拆分")
        self.customs_extract_tab = CustomsExtractTab()
        tab_widget.addTab(self.customs_extract_tab, "报关单提取")

        self.setCentralWidget(tab_widget)

    def closeEvent(self, event):
        if self.customs_extract_tab.is_running():
            event.ignore()
            QMessageBox.warning(
                self,
                "报关单提取进行中",
                "报关单正在提取，暂时不能关闭程序。请等待当前任务完成后再关闭。",
            )
            return
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
