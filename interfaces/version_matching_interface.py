# interfaces/version_matching_interface.py

from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFileDialog,
    QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from qfluentwidgets import PrimaryPushButton
from .base_interface import BaseInterface
from utils.crawler import match_version_numbers

class VersionMatchWorker(QObject):
    finished = Signal()
    progress = Signal(str)
    progress_percent = Signal(int)

    def __init__(self, excel_file):
        super().__init__()
        self.excel_file = excel_file

    def run(self):
        def pcallback(msg):
            self.progress.emit(msg)
        def local_percent(val, _unused):
            """match_version_numbers(val, stage) => 2 params
               这里只用 val，忽略 stage"""
            self.progress_percent.emit(val)

        try:
            match_version_numbers(
                excel_filename=self.excel_file,
                progress_callback=pcallback,
                progress_percent_callback=local_percent,
                create_new_file=True  # 在单独界面 => 另存
            )
        except Exception as e:
            self.progress.emit(f"版号匹配过程中发生错误: {e}")
        finally:
            self.finished.emit()

class VersionMatchingInterface(BaseInterface):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)

        header_layout = QHBoxLayout()
        self.upload_button = PrimaryPushButton("选择Excel文件并匹配")
        self.upload_button.clicked.connect(self.handle_upload)
        header_layout.addWidget(self.upload_button)
        header_layout.addStretch()

        explanation_label = QLabel(
            "说明：请选择包含“游戏名称”列的Excel，执行自动版号匹配。\n"
            "进度条会随每2~3条输出更新。会另存一个“xxx-已匹配版号.xlsx”副本。"
        )
        explanation_label.setWordWrap(True)

        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0,100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)

        self.layout.addLayout(header_layout)
        self.layout.addWidget(explanation_label)
        self.layout.addWidget(self.output_text_edit)
        self.layout.addWidget(self.progress_bar)

    def handle_upload(self):
        dlg=QFileDialog(self)
        dlg.setWindowTitle("选择Excel文件")
        dlg.setNameFilter("Excel Files (*.xlsx *.xls)")
        if dlg.exec():
            files=dlg.selectedFiles()
            if files:
                excel_path=files[0]
                self.thread=QThread()
                self.worker=VersionMatchWorker(excel_path)
                self.worker.moveToThread(self.thread)

                self.thread.started.connect(self.worker.run)
                self.worker.finished.connect(self.thread.quit)
                self.worker.finished.connect(self.worker.deleteLater)
                self.thread.finished.connect(self.thread.deleteLater)

                self.worker.progress.connect(self.on_progress)
                self.worker.progress_percent.connect(self.on_percent)

                self.upload_button.setEnabled(False)
                self.progress_bar.setVisible(True)
                self.progress_bar.setValue(0)

                self.thread.start()
                self.thread.finished.connect(lambda: self.upload_button.setEnabled(True))
                self.thread.finished.connect(self.on_match_finished)

    def on_progress(self, msg):
        self.output_text_edit.append(msg)

    def on_percent(self, val):
        self.progress_bar.setValue(val)
        self.progress_bar.setFormat(f"进度: {val}%")

    def on_match_finished(self):
        self.progress_bar.setValue(100)
        self.output_text_edit.append("\n版号匹配已完成！")
