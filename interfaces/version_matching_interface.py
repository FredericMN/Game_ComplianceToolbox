from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFileDialog
)
from PySide6.QtCore import Qt, QObject, QThread, Signal
from qfluentwidgets import PrimaryPushButton
from .base_interface import BaseInterface
from utils.crawler import match_version_numbers
import traceback

class VersionMatchingWorker(QObject):
    """版号匹配工作线程"""
    finished = Signal()
    progress = Signal(str)

    def __init__(self, excel_file):
        super().__init__()
        # self.init_ui()  # 确保已移除或注释掉此行
        self.excel_file = excel_file

    def run(self):
        try:
            def progress_callback(message):
                self.progress.emit(message)
            # 传递 append_suffix=True 以添加后缀
            match_version_numbers(self.excel_file, progress_callback, append_suffix=True)
        except Exception as e:
            error_message = f"发生错误: {str(e)}\n疑似网络存在问题，请重试。"
            self.progress.emit(error_message)
        finally:
            self.finished.emit()

class VersionMatchingInterface(BaseInterface):
    """版号匹配界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()  # 调用初始化界面的方法

    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)

        header_layout = QHBoxLayout()
        self.upload_button = PrimaryPushButton("上传文件")
        self.upload_button.clicked.connect(self.handle_upload)
        header_layout.addWidget(self.upload_button)
        header_layout.addStretch()

        explanation_label = QLabel("说明：请提供excel格式并包含游戏名称列的文件进行版号信息匹配。")
        explanation_label.setWordWrap(True)

        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("信息输出区域")

        self.layout.addLayout(header_layout)
        self.layout.addWidget(explanation_label)
        self.layout.addWidget(self.output_text_edit)

    def handle_upload(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择Excel文件")
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                file_path = selected_files[0]
                self.thread = QThread()
                self.worker = VersionMatchingWorker(file_path)
                self.worker.moveToThread(self.thread)

                self.thread.started.connect(self.worker.run)
                self.worker.finished.connect(self.thread.quit)
                self.worker.finished.connect(self.worker.deleteLater)
                self.thread.finished.connect(self.thread.deleteLater)
                self.worker.progress.connect(self.report_progress)

                self.thread.start()

                self.upload_button.setEnabled(False)
                self.thread.finished.connect(lambda: self.upload_button.setEnabled(True))

    def report_progress(self, message):
        self.output_text_edit.append(message)
