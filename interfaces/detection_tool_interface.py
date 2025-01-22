# detection_tool_interface.py
import os
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import (
    QLabel, QHBoxLayout, QVBoxLayout, QTextEdit, QPushButton,
    QFileDialog, QMessageBox, QListWidget, QListWidgetItem, QProgressBar, QWidget
)
from .base_interface import BaseInterface
from utils.detection import (
    initialize_words, detect_language, update_words as update_words_func
)
from PySide6.QtGui import QTextCursor  # 如果需要其他文本处理功能

class DetectionWorker(QObject):
    """
    工作线程，用于批量检测文件中的风险词汇。
    """
    progress = Signal(str)
    progress_percent = Signal(int)
    finished = Signal(bool, list)  # success, list of result_dicts

    def __init__(self, file_paths, violent_words, inducing_words):
        super().__init__()
        self.file_paths = file_paths
        self.violent_words = violent_words
        self.inducing_words = inducing_words

    def run(self):
        results = []
        total_files = len(self.file_paths)
        if total_files == 0:
            self.progress.emit("没有选择任何文件。")
            self.finished.emit(False, results)
            return

        for idx, file_path in enumerate(self.file_paths, start=1):
            try:
                self.progress.emit(f"开始检测文件 {idx}/{total_files}: {os.path.basename(file_path)}")
                violent_count, inducing_count, total_word_count, new_file_path = detect_language(
                    file_path, self.violent_words, self.inducing_words
                )
                result = {
                    'file_name': os.path.basename(file_path),
                    'total_word_count': total_word_count,
                    'violent_count': violent_count,
                    'inducing_count': inducing_count,
                    'new_file_path': new_file_path
                }
                results.append(result)
                self.progress.emit(f"完成检测文件 {idx}/{total_files}: {os.path.basename(file_path)}")
            except Exception as e:
                self.progress.emit(f"文件 {os.path.basename(file_path)} 检测过程中发生错误：{str(e)}")
                result = {
                    'file_name': os.path.basename(file_path),
                    'error': str(e)
                }
                results.append(result)
            # 更新进度百分比
            percent = int((idx / total_files) * 100)
            self.progress_percent.emit(percent)

        self.finished.emit(True, results)

class DetectionToolInterface(BaseInterface):
    """不文明用语检测工具界面"""
    def __init__(self, parent=None):
        # 初始化词汇列表
        self.violent_words, self.inducing_words = initialize_words()
        super().__init__(parent)
        self.init_ui()
        self.thread = None  # 初始化线程引用

    def init_ui(self):
        # 设置整体背景色
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f6fa;
            }
        """)

        # 创建一个容器widget来包含所有内容
        container = QWidget()
        container.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 10px;
                padding: 15px;
            }
        """)
        container_layout = QVBoxLayout(container)
        container_layout.setSpacing(15)
        container_layout.setContentsMargins(15, 15, 15, 15)

        # 顶部布局，包含两个输入区域
        top_layout = QHBoxLayout()
        top_layout.setSpacing(15)

        # 左侧输入区域
        left_input_layout = QVBoxLayout()
        left_label = QLabel("血腥暴力词汇（用逗号分隔）")
        left_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                font-size: 14px;
                font-weight: bold;
                margin-bottom: 5px;
            }
        """)
        self.left_text_edit = QTextEdit()
        self.left_text_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
                background-color: #f8f9fa;
                color: #2c3e50;
                font-size: 12px;
                line-height: 1.4;
            }
        """)
        self.left_text_edit.setPlaceholderText("请输入血腥暴力词汇，用逗号分隔")
        self.left_text_edit.setText(','.join(self.violent_words))
        left_input_layout.addWidget(left_label)
        left_input_layout.addWidget(self.left_text_edit)

        # 右侧输入区域样式相同
        right_input_layout = QVBoxLayout()
        right_label = QLabel("不良诱导词汇（用逗号分隔）")
        right_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                font-size: 14px;
                font-weight: bold;
                margin-bottom: 5px;
            }
        """)
        self.right_text_edit = QTextEdit()
        self.right_text_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
                background-color: #f8f9fa;
                color: #2c3e50;
                font-size: 12px;
                line-height: 1.4;
            }
        """)
        self.right_text_edit.setPlaceholderText("请输入不良诱导词汇，用逗号分隔")
        self.right_text_edit.setText(','.join(self.inducing_words))
        right_input_layout.addWidget(right_label)
        right_input_layout.addWidget(self.right_text_edit)

        top_layout.addLayout(left_input_layout)
        top_layout.addLayout(right_input_layout)

        # 按钮布局
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)
        
        button_style = """
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 20px;
                font-weight: bold;
                font-size: 14px;
                min-width: 160px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2573a7;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """
        
        self.update_button = QPushButton("更新并保存词汇列表")
        self.update_button.setStyleSheet(button_style)
        self.select_button = QPushButton("选择文件并检测")
        self.select_button.setStyleSheet(button_style)
        buttons_layout.addWidget(self.update_button)
        buttons_layout.addWidget(self.select_button)

        # 进度条样式
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #e0e0e0;
                border-radius: 5px;
                text-align: center;
                background-color: #f8f9fa;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 5px;
            }
        """)
        self.progress_bar.setVisible(False)

        # 结果输出区域
        result_label = QLabel("结果输出区域")
        result_label.setStyleSheet("""
            QLabel {
                color: #2c3e50;
                font-size: 14px;
                font-weight: bold;
                margin-top: 10px;
            }
        """)
        
        self.result_list_widget = QListWidget()
        self.result_list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 5px;
                background-color: #f8f9fa;
                color: #2c3e50;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #e0e0e0;
            }
            QListWidget::item:last {
                border-bottom: none;
            }
        """)

        # 说明标签
        explanation_label = QLabel(
            "说明：仅支持Word文档与Excel表格进行批量检测，其中【血腥暴力】风险词汇标记为红色，【不良诱导】风险词汇标记为绿色。"
        )
        explanation_label.setStyleSheet("""
            QLabel {
                color: #7f8c8d;
                font-size: 12px;
                padding: 10px;
                background-color: #f8f9fa;
                border-radius: 8px;
            }
        """)
        explanation_label.setWordWrap(True)

        # 将各个布局添加到容器布局
        container_layout.addLayout(top_layout)
        container_layout.addLayout(buttons_layout)
        container_layout.addWidget(self.progress_bar)
        container_layout.addWidget(result_label)
        container_layout.addWidget(self.result_list_widget)
        container_layout.addWidget(explanation_label)

        # 将容器添加到主布局
        self.layout.addWidget(container)

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
                self.progress_bar.setValue(0)
                self.progress_bar.setVisible(True)
                self.progress_bar.setFormat("检测进度: 0%")
                
                # 禁用按钮，防止重复点击
                self.select_button.setEnabled(False)
                self.update_button.setEnabled(False)

                # 启动检测线程
                self.thread = QThread()
                self.worker = DetectionWorker(selected_files, self.violent_words, self.inducing_words)
                self.worker.moveToThread(self.thread)

                # 连接信号
                self.thread.started.connect(self.worker.run)
                self.worker.progress.connect(self.report_progress)
                self.worker.progress_percent.connect(self.update_progress_bar)
                self.worker.finished.connect(self.on_detection_finished)

                # 确保线程在工作完成后正确退出
                self.worker.finished.connect(self.thread.quit)
                self.worker.finished.connect(self.worker.deleteLater)
                self.thread.finished.connect(self.thread.deleteLater)

                # 启动线程
                self.thread.start()

    def report_progress(self, message):
        """
        接收工作线程发送的进度信息，并更新到界面上的 QListWidget。
        """
        if message:
            list_item = QListWidgetItem(message)
            self.result_list_widget.addItem(list_item)
            self.result_list_widget.scrollToBottom()

    def update_progress_bar(self, percent):
        """
        更新进度条的值。
        """
        self.progress_bar.setValue(percent)
        self.progress_bar.setFormat(f"检测进度: {percent}%")

    def on_detection_finished(self, success, results):
        """
        处理检测完成后的逻辑。
        显示结果，并重新启用按钮。
        """
        if success:
            for result in results:
                if 'error' in result:
                    output_text = (
                        f"文件: {result['file_name']}\n"
                        f"检测过程中发生错误：{result['error']}\n"
                    )
                else:
                    output_text = (
                        f"文件: {result['file_name']}\n"
                        f"文字总字数: {result['total_word_count']}\n"
                        f"血腥暴力词汇: {result['violent_count']} 个\n"
                        f"不良诱导词汇: {result['inducing_count']} 个\n"
                        f"已保存标记后的副本: {result['new_file_path']}\n"
                    )
                list_item = QListWidgetItem(output_text)
                self.result_list_widget.addItem(list_item)
            QMessageBox.information(self, "完成", "所有文件检测完成！")
        else:
            QMessageBox.warning(self, "错误", "检测过程中发生错误，请查看输出信息。")

        # 重置进度条和按钮状态
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("检测进度: 0%")
        self.select_button.setEnabled(True)
        self.update_button.setEnabled(True)
