from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QTextEdit, QMessageBox, QFileDialog, QProgressBar, QWidget, QSizePolicy, QComboBox, QHBoxLayout
)
from .base_interface import BaseInterface
from qfluentwidgets import PrimaryPushButton
from utils.large_model import (
    check_and_download_model, analyze_file_with_model, check_model_configured
)
from PySide6.QtGui import QTextCursor  # 导入 QTextCursor 类
import torch  # 导入torch库

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


class AnalyzeFileWorker(QObject):
    """
    工作线程，用于分析文件。
    """
    progress = Signal(str)
    progress_percent = Signal(int)
    finished = Signal(bool, dict)  # success, result_dict

    def __init__(self, file_path, device='cpu'):
        super().__init__()
        self.file_path = file_path
        self.device = device  # 添加设备属性

    def run(self):
        try:
            # 使用大模型分析文件，传递 progress.emit 作为回调
            result = analyze_file_with_model(
                self.file_path, self.emit_progress, self.device)
            self.finished.emit(True, result)
        except Exception as e:
            self.progress.emit(f"发生错误: {str(e)}")
            self.finished.emit(False, {})

    def emit_progress(self, message):
        """
        处理进度信息，解析百分比并发射对应信号。
        """
        self.progress.emit(message)
        if "分析进度" in message:
            try:
                # 提取百分比
                percent_str = message.split("分析进度: ")[1].split("%")[0]
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
        self.analyze_button = PrimaryPushButton("选择文件并检测")
        self.analyze_button.clicked.connect(self.handle_analyze)

        self.analysis_progress_bar = QProgressBar()
        self.analysis_progress_bar.setRange(0, 100)
        self.analysis_progress_bar.setValue(0)
        self.analysis_progress_bar.setTextVisible(True)

        self.analysis_progress_text_edit = QTextEdit()
        self.analysis_progress_text_edit.setReadOnly(True)
        self.analysis_progress_text_edit.setPlaceholderText("大模型分析文档的进度展示")
        self.analysis_progress_text_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        bottom_layout.addWidget(self.analyze_button)
        bottom_layout.addWidget(self.analysis_progress_bar)
        bottom_layout.addWidget(self.analysis_progress_text_edit)
        bottom_layout.setStretch(0, 1)
        bottom_layout.setStretch(1, 1)
        bottom_layout.setStretch(2, 4)  # 增加分析输出框的占比

        # 添加说明标签
        explanation_label = QLabel(
            "说明：仅支持Word文档与Excel表格进行分析。\n"
            "分类标签及颜色标记如下：\n"
            "• 正常（无标记）"
            "• 低俗（绿色）"
            "• 色情（黄色）"
            "• 其他风险（红色）"
            "• 成人（蓝色）"
        )

        # 将两个布局添加到主布局
        self.layout.addLayout(top_layout)
        self.layout.addLayout(bottom_layout)
        self.layout.addWidget(explanation_label)
        self.layout.setStretch(0, 1)
        self.layout.setStretch(1, 4)
        self.layout.setStretch(2, 0)

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
        
        # 关键修复：确保线程在工作完成后正确退出
        self.worker.finished.connect(self.thread.quit)  # 添加这一行
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

        # 打开文件选择对话框
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择文件")
        file_dialog.setNameFilter("Word/Excel Files (*.docx *.xlsx)")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                file_path = selected_files[0]
                # 启动分析线程
                self.analyze_button.setEnabled(False)
                self.analysis_progress_text_edit.clear()
                self.analysis_progress_bar.setValue(0)

                self.analysis_thread = QThread()
                self.analysis_worker = AnalyzeFileWorker(file_path, device=self.device)  # 传递设备信息
                self.analysis_worker.moveToThread(self.analysis_thread)

                # 连接信号
                self.analysis_thread.started.connect(self.analysis_worker.run)
                self.analysis_worker.progress.connect(self.report_analysis_progress)
                self.analysis_worker.progress_percent.connect(self.analysis_progress_bar.setValue)
                self.analysis_worker.finished.connect(self.on_analysis_finished)
                
                # 关键修复：确保线程在工作完成后正确退出
                self.analysis_worker.finished.connect(self.analysis_thread.quit)  # 添加这一行
                self.analysis_worker.finished.connect(self.analysis_worker.deleteLater)
                self.analysis_thread.finished.connect(self.analysis_thread.deleteLater)

                # 启动线程
                self.analysis_thread.start()

    def report_analysis_progress(self, message):
        """
        接收工作线程发送的分析进度信息，并更新到界面上的 QTextEdit。
        避免添加过多空行。
        """
        message = message.strip()
        if message:
            self.analysis_progress_text_edit.append(message)
            self.analysis_progress_text_edit.moveCursor(QTextCursor.End)  # 修正 AttributeError

    def on_analysis_finished(self, success, result):
        """
        处理分析完成后的逻辑。
        显示信息框，并重新启用按钮。
        """
        if success:
            output_text = (
                f"识别完毕，文字总字数为 {result['total_word_count']}。\n"
                f"分析结果：\n"
                f"正常内容 {result['normal_count']} 段，"
                f"低俗内容 {result['low_vulgar_count']} 段，"
                f"色情内容 {result['porn_count']} 段，"
                f"其他风险内容 {result['other_risk_count']} 段，"
                f"成人内容 {result['adult_count']} 段。其中低俗、色情、其他风险、成人内容已进行颜色标记。\n"
                f"已保存标记后的副本：{result['new_file_path']}"
            )
            self.analysis_progress_text_edit.append(output_text)
            self.analysis_progress_text_edit.moveCursor(QTextCursor.End)
            QMessageBox.information(self, "完成", f"分析完成！结果已保存到 {result['new_file_path']}")
        else:
            QMessageBox.warning(self, "错误", "分析失败，请查看输出信息。")
        self.analyze_button.setEnabled(True)