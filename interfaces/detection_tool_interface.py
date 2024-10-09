# detection_tool_interface.py
import os
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QLabel, QHBoxLayout, QVBoxLayout, QTextEdit, QPushButton, QFileDialog, QMessageBox, QListWidget, QListWidgetItem
)
from .base_interface import BaseInterface
from utils.detection import (
    initialize_words, detect_language, update_words as update_words_func
)

class DetectionToolInterface(BaseInterface):
    """不文明用语检测工具界面"""
    def __init__(self, parent=None):
        # 初始化词汇列表
        self.violent_words, self.inducing_words = initialize_words()
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        # 顶部布局，包含两个输入区域
        top_layout = QHBoxLayout()

        # 左侧输入区域
        left_input_layout = QVBoxLayout()
        left_label = QLabel("血腥暴力词汇（用逗号分隔）")
        self.left_text_edit = QTextEdit()
        self.left_text_edit.setPlaceholderText("请输入血腥暴力词汇，用逗号分隔")
        self.left_text_edit.setText(','.join(self.violent_words))
        left_input_layout.addWidget(left_label)
        left_input_layout.addWidget(self.left_text_edit)

        # 右侧输入区域
        right_input_layout = QVBoxLayout()
        right_label = QLabel("不良诱导词汇（用逗号分隔）")
        self.right_text_edit = QTextEdit()
        self.right_text_edit.setPlaceholderText("请输入不良诱导词汇，用逗号分隔")
        self.right_text_edit.setText(','.join(self.inducing_words))
        right_input_layout.addWidget(right_label)
        right_input_layout.addWidget(self.right_text_edit)

        top_layout.addLayout(left_input_layout)
        top_layout.addLayout(right_input_layout)

        # 按钮布局
        buttons_layout = QHBoxLayout()
        self.update_button = QPushButton("更新并保存词汇列表")
        self.select_button = QPushButton("选择文件并检测")
        buttons_layout.addWidget(self.update_button)
        buttons_layout.addWidget(self.select_button)

        # 结果输出区域
        result_label = QLabel("结果输出区域")
        self.result_list_widget = QListWidget()

        # 添加说明标签
        explanation_label = QLabel(
            "说明：仅支持Word文档与Excel表格进行批量检测，其中【血腥暴力】风险词汇标记为红色，【不良诱导】风险词汇标记为绿色。"
        )

        # 将各个布局添加到主布局
        self.layout.addLayout(top_layout)
        self.layout.addLayout(buttons_layout)
        self.layout.addWidget(result_label)
        self.layout.addWidget(self.result_list_widget)
        self.layout.addWidget(explanation_label)  # 添加说明标签到UI布局

        # 连接按钮信号
        self.update_button.clicked.connect(self.handle_update)
        self.select_button.clicked.connect(self.handle_select_files)

    def handle_update(self):
        violent_input = self.left_text_edit.toPlainText().strip()
        inducing_input = self.right_text_edit.toPlainText().strip()

        if not violent_input or not inducing_input:
            QMessageBox.warning(self, "输入错误", "词汇列表不能为空！")
            return

        try:
            self.violent_words, self.inducing_words = update_words_func(violent_input, inducing_input)
            self.left_text_edit.setText(','.join(self.violent_words))
            self.right_text_edit.setText(','.join(self.inducing_words))
            QMessageBox.information(self, "词汇更新", "词汇列表已更新并保存。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"更新词汇列表时发生错误：{str(e)}")

    def handle_select_files(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择文件")
        file_dialog.setNameFilter("Word/Excel Files (*.docx *.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)  # 允许多文件选择
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.result_list_widget.clear()  # 清空之前的结果
                for file_path in selected_files:
                    try:
                        violent_count, inducing_count, total_word_count, new_file_path = detect_language(
                            file_path, self.violent_words, self.inducing_words
                        )
                        output_text = (
                            f"文件: {os.path.basename(file_path)}\n"
                            f"文字总字数: {total_word_count}\n"
                            f"血腥暴力词汇: {violent_count} 个\n"
                            f"不良诱导词汇: {inducing_count} 个\n"
                            f"已保存标记后的副本: {new_file_path}\n"
                        )
                        list_item = QListWidgetItem(output_text)
                        self.result_list_widget.addItem(list_item)
                    except Exception as e:
                        error_text = (
                            f"文件: {os.path.basename(file_path)}\n"
                            f"检测过程中发生错误：{str(e)}\n"
                        )
                        list_item = QListWidgetItem(error_text)
                        self.result_list_widget.addItem(list_item)
                QMessageBox.information(self, "完成", "所有文件检测完成！")
