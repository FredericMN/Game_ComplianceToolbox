# large_model_interface.py

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QTextEdit, QMessageBox, QFileDialog, QProgressBar, QWidget, QSizePolicy, QComboBox, QHBoxLayout, QListWidget, QListWidgetItem, QLineEdit, QGroupBox, QPushButton
)
from .base_interface import BaseInterface
from qfluentwidgets import PrimaryPushButton
from utils.large_model import (
    check_and_download_model, analyze_files_with_model, check_model_configured
)
from PySide6.QtGui import QTextCursor, QDoubleValidator  # 导入 QTextCursor 和 QDoubleValidator
import torch  # 导入torch库
import os


class DownloadModelWorker(QObject):
    """
    工作线程，用于下载大模型。
    """
    progress = Signal(str)
    progress_percent = Signal(int)
    finished = Signal(bool)

    def __init__(self, device='cpu'):
        super().__init__()
        self.device = device  # 添加设备属性

    def run(self):
        try:
            # 检查并下载大模型，传递 progress.emit 作为回调
            check_and_download_model(self.emit_progress)
            self.finished.emit(True)
        except Exception as e:
            self.progress.emit(f"发生错误: {str(e)}")
            self.finished.emit(False)

    def emit_progress(self, message):
        """
        处理进度信息，解析百分比并发射对应信号。
        """
        self.progress.emit(message)
        if "下载进度" in message:
            try:
                # 提取百分比
                percent_str = message.split("下载进度: ")[1].split("%")[0]
                percent = int(percent_str)
                self.progress_percent.emit(percent)
            except:
                pass  # 无法解析百分比则忽略


class AnalyzeFilesWorker(QObject):
    """
    工作线程，用于分析多个文件。
    """
    progress = Signal(str)
    progress_percent = Signal(int)
    finished = Signal(bool, list)  # success, list of result_dicts

    def __init__(self, file_paths, device='cpu', normal_threshold=0.8, other_threshold=0.1):
        super().__init__()
        self.file_paths = file_paths
        self.device = device  # 添加设备属性
        self.normal_threshold = normal_threshold
        self.other_threshold = other_threshold

    def run(self):
        try:
            results = analyze_files_with_model(
                self.file_paths, self.emit_progress, self.device,
                self.normal_threshold, self.other_threshold)
            self.finished.emit(True, results)
        except Exception as e:
            self.progress.emit(f"分析过程中发生错误：{str(e)}")
            self.finished.emit(False, [])

    def emit_progress(self, message):
        """
        处理进度信息，解析百分比并发射对应信号。
        """
        self.progress.emit(message)
        if "进度" in message:
            try:
                # 提取百分比
                percent_str = message.split("进度: ")[1].split("%")[0]
                percent = int(percent_str)
                self.progress_percent.emit(percent)
            except:
                pass  # 无法解析百分比则忽略


class LargeModelInterface(BaseInterface):
    """大模型语义分析界面"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.device = 'cpu'  # 默认设备为CPU
        self.init_ui()

    # 在 init_ui 方法中更新说明标签
    def init_ui(self):
        self.layout.setAlignment(Qt.AlignTop)

        # 设备选择布局
        device_layout = QHBoxLayout()
        device_label = QLabel("选择设备:")
        self.device_combo = QComboBox()
        self.device_combo.addItem("CPU")
        if torch.cuda.is_available():
            self.device_combo.addItem("GPU")
        else:
            self.device_combo.addItem("GPU (不可用)")
            self.device_combo.setItemData(1, "GPU不可用", Qt.ToolTipRole)
            self.device_combo.setEnabled(False)  # 如果GPU不可用，禁用选择

        self.device_combo.currentIndexChanged.connect(self.handle_device_change)
        device_layout.addWidget(device_label)
        device_layout.addWidget(self.device_combo)

        # 上部区域：下载和配置大模型
        top_layout = QVBoxLayout()
        top_layout.addLayout(device_layout)  # 添加设备选择到上部布局

        self.configure_button = PrimaryPushButton("检测并配置大模型环境")
        self.configure_button.clicked.connect(self.handle_configure)

        self.download_progress_bar = QProgressBar()
        self.download_progress_bar.setRange(0, 100)
        self.download_progress_bar.setValue(0)
        self.download_progress_bar.setTextVisible(True)

        self.progress_text_edit = QTextEdit()
        self.progress_text_edit.setReadOnly(True)
        self.progress_text_edit.setPlaceholderText("下载大模型进度信息输出区域")
        self.progress_text_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        top_layout.addWidget(self.configure_button)
        top_layout.addWidget(self.download_progress_bar)
        top_layout.addWidget(self.progress_text_edit)
        top_layout.setStretch(1, 1)
        top_layout.setStretch(2, 2)

        # 下部区域：选择文件并分析
        bottom_layout = QVBoxLayout()

        # 新增折叠栏
        self.threshold_group_box = QGroupBox("设置阈值(如非特殊需求，建议保持默认)")
        self.threshold_group_box.setCheckable(True)
        self.threshold_group_box.setChecked(False)
        threshold_layout = QHBoxLayout()

        # NORMAL_THRESHOLD 输入框
        normal_label = QLabel("正常阈值 (0-1):")
        self.normal_input = QLineEdit("0.8")
        self.normal_input.setValidator(QDoubleValidator(0.00, 1.00, 2))
        self.normal_input.setFixedWidth(60)

        # OTHER_THRESHOLD 输入框
        other_label = QLabel("其他阈值 (0-1):")
        self.other_input = QLineEdit("0.1")
        self.other_input.setValidator(QDoubleValidator(0.00, 1.00, 2))
        self.other_input.setFixedWidth(60)

        threshold_layout.addWidget(normal_label)
        threshold_layout.addWidget(self.normal_input)
        threshold_layout.addWidget(other_label)
        threshold_layout.addWidget(self.other_input)
        threshold_layout.addStretch()

        self.threshold_group_box.setLayout(threshold_layout)

        # 添加折叠栏到下部布局
        bottom_layout.addWidget(self.threshold_group_box)

        self.analyze_button = PrimaryPushButton("选择文件并检测")
        self.analyze_button.clicked.connect(self.handle_analyze)

        self.analysis_progress_bar = QProgressBar()
        self.analysis_progress_bar.setRange(0, 100)
        self.analysis_progress_bar.setValue(0)
        self.analysis_progress_bar.setTextVisible(True)

        self.analysis_progress_list_widget = QListWidget()

        # 添加说明标签
        explanation_label = QLabel(
            "说明：仅支持Word文档与Excel表格进行分析。\n"
            "分类标签及颜色标记如下：\n"
            "• 正常（无标记）\n"
            "• 低俗（绿色）\n"
            "• 色情（黄色）\n"
            "• 其他风险（红色）\n"
            "• 成人（蓝色）"
        )

        # 将两个布局添加到主布局
        self.layout.addLayout(top_layout)
        self.layout.addLayout(bottom_layout)
        self.layout.addWidget(explanation_label)
        self.layout.setStretch(0, 1)
        self.layout.setStretch(1, 4)
        self.layout.setStretch(2, 0)

        # 下部布局内容
        bottom_layout.addWidget(self.analyze_button)
        bottom_layout.addWidget(self.analysis_progress_bar)
        bottom_layout.addWidget(self.analysis_progress_list_widget)
        bottom_layout.setStretch(0, 1)
        bottom_layout.setStretch(1, 1)
        bottom_layout.setStretch(2, 4)  # 增加分析输出框的占比

    def handle_device_change(self, index):
        selected = self.device_combo.currentText()
        if selected == "CPU":
            self.device = 'cpu'
        elif selected == "GPU":
            if torch.cuda.is_available():
                self.device = 'cuda'
                QMessageBox.information(self, "提示", "恭喜，你可以使用GPU进行加速运算。")
            else:
                self.device = 'cpu'
                QMessageBox.warning(self, "提示", "抱歉，当前环境无法使用GPU进行运算。")
                self.device_combo.setCurrentText("CPU")  # 重置为CPU

    def handle_configure(self):
        """
        处理“检测并配置大模型环境”按钮点击事件。
        启动下载模型的工作线程，并连接信号。
        """
        # 检查模型是否已配置
        if check_model_configured():
            QMessageBox.information(self, "提示", "大模型已经配置，无需重复配置。")
            return

        # 禁用按钮，防止重复点击
        self.configure_button.setEnabled(False)
        self.progress_text_edit.clear()
        self.download_progress_bar.setValue(0)

        # 启动下载线程
        self.thread = QThread()
        self.worker = DownloadModelWorker(device=self.device)  # 传递设备信息
        self.worker.moveToThread(self.thread)

        # 连接信号
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.report_progress)
        self.worker.progress_percent.connect(self.download_progress_bar.setValue)
        self.worker.finished.connect(self.on_configure_finished)

        # 确保线程在工作完成后正确退出
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        # 启动线程
        self.thread.start()

    def report_progress(self, message):
        """
        接收工作线程发送的进度信息，并更新到界面上的 QTextEdit。
        避免添加过多空行。
        """
        message = message.strip()
        if message:
            self.progress_text_edit.append(message)
            self.progress_text_edit.moveCursor(QTextCursor.End)  # 修正 AttributeError

    def on_configure_finished(self, success):
        """
        处理下载完成后的逻辑。
        显示信息框，并重新启用按钮。
        """
        if success:
            QMessageBox.information(self, "完成", "大模型配置完成！")
        else:
            QMessageBox.warning(self, "错误", "大模型配置失败，请查看输出信息。")
        self.configure_button.setEnabled(True)

    def handle_analyze(self):
        """
        处理“选择文件并检测”按钮点击事件。
        先检查大模型是否已配置完成，如果完成则继续，否则提示用户先下载配置。
        """
        # 检查模型是否已配置
        if not check_model_configured():
            QMessageBox.warning(self, "提示", "大模型尚未配置，请先下载并配置大模型。")
            return

        # 获取阈值输入
        normal_threshold_text = self.normal_input.text()
        other_threshold_text = self.other_input.text()

        try:
            normal_threshold = float(normal_threshold_text) if normal_threshold_text else 0.95
            other_threshold = float(other_threshold_text) if other_threshold_text else 0.5

            # 确保阈值在0-1之间
            if not (0.0 <= normal_threshold <= 1.0) or not (0.0 <= other_threshold <= 1.0):
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "输入错误", "阈值必须是0到1之间的数字，最多两位小数。")
            return

        # 打开文件选择对话框
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择文件")
        file_dialog.setNameFilter("Word/Excel Files (*.docx *.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFiles)  # 允许多文件选择
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.analysis_progress_list_widget.clear()  # 清空之前的结果
                for file_path in selected_files:
                    list_item = QListWidgetItem(f"准备分析: {os.path.basename(file_path)}")
                    self.analysis_progress_list_widget.addItem(list_item)

                # 禁用按钮，防止重复点击
                self.analyze_button.setEnabled(False)
                self.analysis_progress_bar.setValue(0)

                # 启动分析线程
                self.analysis_thread = QThread()
                self.analysis_worker = AnalyzeFilesWorker(
                    selected_files, device=self.device,
                    normal_threshold=normal_threshold,
                    other_threshold=other_threshold
                )  # 传递设备信息和阈值
                self.analysis_worker.moveToThread(self.analysis_thread)

                # 连接信号
                self.analysis_thread.started.connect(self.analysis_worker.run)
                self.analysis_worker.progress.connect(self.report_analysis_progress)
                self.analysis_worker.progress_percent.connect(self.analysis_progress_bar.setValue)
                self.analysis_worker.finished.connect(self.on_analysis_finished)

                # 确保线程在工作完成后正确退出
                self.analysis_worker.finished.connect(self.analysis_thread.quit)
                self.analysis_worker.finished.connect(self.analysis_worker.deleteLater)
                self.analysis_thread.finished.connect(self.analysis_thread.deleteLater)

                # 启动线程
                self.analysis_thread.start()

    def report_analysis_progress(self, message):
        """
        接收工作线程发送的分析进度信息，并更新到界面上的 QListWidget。
        为了优化性能，仅在关键进度点更新。
        """
        message = message.strip()
        if message:
            # 根据消息内容决定如何更新
            if message.startswith("开始分析文件"):
                current_item = self.analysis_progress_list_widget.currentItem()
                if current_item:
                    current_item.setText(message)
            elif message.startswith("完成分析文件"):
                current_item = self.analysis_progress_list_widget.currentItem()
                if current_item:
                    current_item.setText(message)
            elif message.startswith("进度"):
                # 可以选择不更新，或者仅更新进度条
                pass
            else:
                # 其他消息可以选择性显示
                pass

    def on_analysis_finished(self, success, results):
        """
        处理分析完成后的逻辑。
        显示信息框，并重新启用按钮。
        """
        if success:
            for result in results:
                output_text = (
                    f"文件: {os.path.basename(result['file_path'])}\n"
                    f"文字总字数: {result['total_word_count']}。\n"
                    f"分析结果：\n"
                    f"正常内容 {result['normal_count']} 段，"
                    f"低俗内容 {result['low_vulgar_count']} 段，"
                    f"色情内容 {result['porn_count']} 段，"
                    f"其他风险内容 {result['other_risk_count']} 段，"
                    f"成人内容 {result['adult_count']} 段。\n"
                    f"已保存标记后的副本：{result['new_file_path']}\n"
                )
                list_item = QListWidgetItem(output_text)
                self.analysis_progress_list_widget.addItem(list_item)
            QMessageBox.information(self, "完成", "所有文件分析完成！结果已保存。")
        else:
            QMessageBox.warning(self, "错误", "分析过程中发生错误，请查看输出信息。")
        self.analyze_button.setEnabled(True)
