# project-01/interfaces/vocabulary_comparison_interface.py

from PySide6.QtCore import Qt, QThread
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QTextEdit
)
from qfluentwidgets import PrimaryPushButton
from .base_interface import BaseInterface
from utils.vocabulary_comparison import VocabularyComparisonProcessor
import os

class VocabularyComparisonInterface(BaseInterface):
    """词表对照界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.a_file_path = ""
        self.b_file_path = ""
        self.thread = None  # 初始化线程变量

    def init_ui(self):
        # 说明文本
        instruction_label = QLabel("说明：可分别选择A、B两个词表进行对比并查看结果，结果默认输出在词表A的地址中。")
        instruction_label.setWordWrap(True)
        #instruction_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")

        # 文件选择布局
        file_selection_layout = QHBoxLayout()

        # 选择A词表
        self.a_button = PrimaryPushButton("选择A词表")
        self.a_button.clicked.connect(self.select_a_file)
        self.a_label = QLabel("未选择A词表")
        self.a_label.setWordWrap(True)
        file_selection_layout.addWidget(self.a_button)
        file_selection_layout.addWidget(self.a_label)

        # 选择B词表
        self.b_button = PrimaryPushButton("选择B词表")
        self.b_button.clicked.connect(self.select_b_file)
        self.b_label = QLabel("未选择B词表")
        self.b_label.setWordWrap(True)
        file_selection_layout.addWidget(self.b_button)
        file_selection_layout.addWidget(self.b_label)

        # 比较按钮
        self.compare_button = PrimaryPushButton("开始对照")
        self.compare_button.clicked.connect(self.handle_compare)
        self.compare_button.setEnabled(False)  # 初始禁用

        # 结果输出区域
        self.result_text_edit = QTextEdit()
        self.result_text_edit.setReadOnly(True)
        self.result_text_edit.setPlaceholderText("结果输出区域")

        # 添加到主布局
        self.layout.addWidget(instruction_label)  # 添加说明文本
        self.layout.addLayout(file_selection_layout)
        self.layout.addWidget(self.compare_button)
        self.layout.addWidget(self.result_text_edit)

    def select_a_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择A词表文件", "", "文档文件 (*.docx *.xlsx *.xls *.txt)"
        )
        if file_path:
            self.a_file_path = file_path
            self.a_label.setText(file_path)
            self.check_files_selected()

    def select_b_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择B词表文件", "", "文档文件 (*.docx *.xlsx *.xls *.txt)"
        )
        if file_path:
            self.b_file_path = file_path
            self.b_label.setText(file_path)
            self.check_files_selected()

    def check_files_selected(self):
        """检查是否已选择两个文件"""
        if self.a_file_path and self.b_file_path:
            self.compare_button.setEnabled(True)

    def handle_compare(self):
        if not self.a_file_path or not self.b_file_path:
            QMessageBox.warning(self, "文件未选择", "请先选择A词表和B词表文件。")
            return

        # 输出目录为A词表所在目录
        output_dir = os.path.dirname(self.a_file_path)
        if not output_dir:
            QMessageBox.critical(self, "错误", "无法确定A词表文件的目录。")
            return

        # 禁用按钮，防止重复点击
        self.compare_button.setEnabled(False)
        self.result_text_edit.append("开始词表对照...")

        # 启动处理器线程
        self.thread = QThread()
        self.processor = VocabularyComparisonProcessor(self.a_file_path, self.b_file_path, output_dir)
        self.processor.moveToThread(self.thread)
        self.thread.started.connect(self.processor.run)
        self.processor.finished.connect(self.on_finished)
        self.processor.finished.connect(self.processor.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_finished(self, output_path, a_count, b_count, merged_count, a_missing_count, b_missing_count):
        if output_path.startswith("错误:"):
            QMessageBox.critical(self, "对照失败", output_path)
            self.result_text_edit.append(output_path)
        else:
            # 显示统计结果
            result_text = (
                f"A词表词汇数量: {a_count}\n"
                f"B词表词汇数量: {b_count}\n"
                f"A+B词表词汇数量: {merged_count}\n"
                f"A词表缺失的B词表词汇数量: {a_missing_count}\n"
                f"B词表缺失的A词表词汇数量: {b_missing_count}\n"
                f"对照结果已保存至: {output_path}"
            )
            self.result_text_edit.append(result_text)
            QMessageBox.information(self, "对照完成", "词表对照已完成并保存。")
        # 重新启用按钮
        self.compare_button.setEnabled(True)
        # 清理线程
        if self.thread is not None:
            self.thread.quit()
            self.thread.wait()
            self.thread = None
