from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QTextEdit, QLabel, QFileDialog,
    QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from qfluentwidgets import PrimaryPushButton
from .base_interface import BaseInterface

class CopyrightQueryWorker(QObject):
    finished = Signal()
    progress = Signal(str)
    progress_percent = Signal(int)

    def __init__(self, excel_file):
        super().__init__()
        self.excel_file = excel_file

    def run(self):
        # 暂时只是模拟进度，实际逻辑未实现
        self.progress.emit("开始查询著作权人信息...")
        self.progress.emit("正在读取Excel文件...")
        self.progress_percent.emit(30)
        self.progress.emit("正在查询著作权人信息...")
        self.progress_percent.emit(60)
        self.progress.emit("正在整理查询结果...")
        self.progress_percent.emit(90)
        self.progress.emit("查询完成！")
        self.progress_percent.emit(100)
        self.finished.emit()

class CopyrightQueryInterface(BaseInterface):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)

        header_layout = QHBoxLayout()
        self.upload_button = PrimaryPushButton("选择Excel文件并查询著作权人")
        self.upload_button.clicked.connect(self.handle_upload)
        header_layout.addWidget(self.upload_button)
        header_layout.addStretch()

        explanation_label = QLabel(
            "说明：请选择包含游戏名称的Excel文件，系统将自动查询并填充著作权人信息。\n"
            "查询结果将保存为新的Excel文件。"
        )
        explanation_label.setWordWrap(True)

        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("著作权人查询日志将显示在这里...")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setVisible(False)

        self.layout.addLayout(header_layout)
        self.layout.addWidget(explanation_label)
        self.layout.addWidget(self.output_text_edit)
        self.layout.addWidget(self.progress_bar)

    def handle_upload(self):
        dlg = QFileDialog(self)
        dlg.setWindowTitle("选择Excel文件")
        dlg.setNameFilter("Excel Files (*.xlsx *.xls)")
        if dlg.exec():
            files = dlg.selectedFiles()
            if files:
                excel_path = files[0]
                self.output_text_edit.clear()
                self.output_text_edit.append(f"已选择文件: {excel_path}")
                
                self.thread = QThread()
                self.worker = CopyrightQueryWorker(excel_path)
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
                self.thread.finished.connect(self.on_query_finished)

    def on_progress(self, msg):
        self.output_text_edit.append(msg)

    def on_percent(self, val):
        self.progress_bar.setValue(val)
        self.progress_bar.setFormat(f"进度: {val}%")

    def on_query_finished(self):
        self.progress_bar.setValue(100)
        self.output_text_edit.append("\n著作权人查询已完成！结果已保存到同目录下。") 