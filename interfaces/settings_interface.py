# interfaces/settings_interface.py

from PySide6.QtCore import Qt, QTimer, QThread, Signal, QObject, QUrl, QSize
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QPushButton, QHBoxLayout, QFrame,
    QMessageBox, QTextEdit, QApplication, QSizePolicy, QDialog, QDialogButtonBox, QProgressBar
)
from PySide6.QtGui import QIcon, QDesktopServices, QFont
from .base_interface import BaseInterface
from utils.version import __version__
from utils.version_checker import VersionChecker, VersionCheckWorker, DownloadWorker
import os
import sys

GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
OWNER = "FredericMN"  # 替换为你的 GitHub 用户名
REPO = "Game_ComplianceToolbox"  # 替换为你的仓库名称

def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容开发和打包后的环境"""
    try:
        # PyInstaller 创建临时文件夹并将路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    return os.path.join(base_path, relative_path)

class VersionSelectionDialog(QDialog):
    """版本选择对话框"""
    def __init__(self, release_notes, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择下载版本")
        self.setModal(True)
        self.setFixedSize(500, 350)  # 增加高度以容纳说明文本
        layout = QVBoxLayout()

        # Release Notes 显示
        release_label = QLabel("更新内容：")
        release_label.setWordWrap(True)
        release_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(release_label)

        self.release_text = QTextEdit()
        self.release_text.setReadOnly(True)
        self.release_text.setPlainText(release_notes if release_notes else "无更新说明。")
        layout.addWidget(self.release_text)

        # 版本选择说明
        description_label = QLabel("拥有英伟达20系以上显卡且安装了最新显卡驱动的用户建议选择GPU版本。")
        description_label.setWordWrap(True)
        description_label.setStyleSheet("color: #555555;")
        layout.addWidget(description_label)

        self.button_box = QDialogButtonBox()
        self.cpu_button = QPushButton("下载CPU版")
        self.gpu_button = QPushButton("下载GPU版")
        self.cancel_button = QPushButton("取消更新")
        self.button_box.addButton(self.cpu_button, QDialogButtonBox.ActionRole)
        self.button_box.addButton(self.gpu_button, QDialogButtonBox.ActionRole)
        self.button_box.addButton(self.cancel_button, QDialogButtonBox.RejectRole)

        layout.addWidget(self.button_box)
        self.setLayout(layout)

        self.cpu_button.clicked.connect(self.accept_cpu)
        self.gpu_button.clicked.connect(self.accept_gpu)
        self.cancel_button.clicked.connect(self.reject)

        self.selected_version = None  # 'cpu', 'gpu', or None

    def accept_cpu(self):
        self.selected_version = 'cpu'
        self.accept()

    def accept_gpu(self):
        self.selected_version = 'gpu'
        self.accept()

class SettingsInterface(BaseInterface):
    """设定界面"""
    # 更新用的 GitHub 令牌
    ComplicanceToolbox_Update_Token = "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"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_version = __version__
        self.update_thread = None  # 更新检查线程
        self.download_thread = None  # 下载线程
        self.selected_version = None  # 'cpu' 或 'gpu'
        self.init_ui()
        self.load_local_update_logs()  # 加载本地更新日志

    def init_ui(self):
        # 创建主垂直布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # 创建“检查版本更新”按钮的水平布局
        update_button_layout = QHBoxLayout()
        update_button_layout.setAlignment(Qt.AlignLeft)

        # 创建“检查版本更新”按钮
        self.check_update_button = QPushButton("检查版本更新")
        self.check_update_button.setFixedHeight(40)  # 增加按钮高度
        self.check_update_button.setFixedWidth(200)   # 设置按钮宽度，不横跨整个界面
        # 设置按钮字体为加粗
        font = QFont()
        font.setBold(True)
        self.check_update_button.setFont(font)
        # 移除自定义样式，使用默认样式
        self.check_update_button.setStyleSheet("")
        # 连接按钮点击信号到处理函数
        self.check_update_button.clicked.connect(self.handle_check_update)

        # 将按钮添加到水平布局
        update_button_layout.addWidget(self.check_update_button)

        # 将按钮布局添加到主布局
        main_layout.addLayout(update_button_layout)

        # 信息输出区域
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("信息输出区域")

        # 添加输出区域到主布局
        main_layout.addWidget(self.output_text_edit)

        # 创建下载进度条
        self.download_progress_bar = QProgressBar()
        self.download_progress_bar.setVisible(False)  # 初始隐藏
        self.download_progress_bar.setMinimum(0)
        self.download_progress_bar.setMaximum(100)
        main_layout.addWidget(self.download_progress_bar)

        # 创建分割线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #CCCCCC;")
        separator.setFixedHeight(2)

        # 添加分割线到主布局
        main_layout.addWidget(separator)

        # 创建更新日志标签
        update_log_label = QLabel("更新日志：")
        update_log_label.setAlignment(Qt.AlignLeft)
        update_log_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333;")

        # 创建更新日志内容
        self.update_log_text_edit = QTextEdit()
        self.update_log_text_edit.setReadOnly(True)
        self.update_log_text_edit.setStyleSheet("font-size: 14px; color: #555555;")
        self.update_log_text_edit.setFixedHeight(150)

        # 添加更新日志标签和内容到主布局
        main_layout.addWidget(update_log_label)
        main_layout.addWidget(self.update_log_text_edit)

        # 添加Stretch以使后续内容位于底部
        main_layout.addStretch()

        # 创建开发者信息标签
        developer_label = QLabel("开发者：合规团队")
        developer_label.setAlignment(Qt.AlignLeft)
        developer_label.setStyleSheet("font-size: 14px; color: #555555;")

        # 创建当前版本标签
        version_label = QLabel(f"当前版本：{self.current_version}")
        version_label.setAlignment(Qt.AlignLeft)
        version_label.setStyleSheet("font-size: 14px; color: #555555;")

        # 添加信息标签到主布局
        main_layout.addWidget(developer_label)
        main_layout.addWidget(version_label)

        # 将主布局添加到 BaseInterface 的布局中
        self.layout.addLayout(main_layout)

    def load_local_update_logs(self):
        try:
            # 获取当前程序的目录
            if getattr(sys, 'frozen', False):
                # 如果是打包后的可执行文件
                base_path = sys._MEIPASS
                # 在打包后，CHANGELOG.md 位于 base_path 目录下
                changelog_path = os.path.join(base_path, 'CHANGELOG.md')
            else:
                # 如果是直接运行的脚本
                base_path = os.path.dirname(os.path.abspath(__file__))
                # 假设 CHANGELOG.md 位于脚本文件所在目录的上一级目录
                changelog_path = os.path.join(base_path, '..', 'CHANGELOG.md')

            # 检查文件是否存在
            if not os.path.exists(changelog_path):
                raise FileNotFoundError(f"未找到文件 {changelog_path}")

            with open(changelog_path, 'r', encoding='utf-8') as f:
                changelog_content = f.read()
                self.update_log_text_edit.setPlainText(changelog_content)
        except Exception as e:
            self.update_log_text_edit.setPlainText(f"无法读取更新日志：{str(e)}")

    def handle_check_update(self):
        self.check_update_button.setEnabled(False)
        self.output_text_edit.append("正在检查更新...")

        # 如果已有线程在运行，先停止
        if self.update_thread and self.update_thread.isRunning():
            self.update_thread.quit()
            self.update_thread.wait()

        self.update_thread = QThread()
        self.version_checker = VersionChecker()
        self.worker = VersionCheckWorker(self.version_checker)
        self.worker.moveToThread(self.update_thread)

        self.update_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_update_check_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.update_thread.finished.connect(self.update_thread.deleteLater)

        self.worker.progress.connect(self.report_progress)

        self.update_thread.start()

        self.worker.finished.connect(lambda: self.check_update_button.setEnabled(True))

    def report_progress(self, message):
        self.output_text_edit.append(message)

    def on_update_check_finished(self, is_new_version, latest_version, cpu_asset_id, cpu_asset_name, gpu_asset_id, gpu_asset_name, release_notes):
        """处理版本检测完成后的逻辑"""
        if is_new_version:
            # 使用自定义对话框让用户选择下载版本，并传递release_notes
            dialog = VersionSelectionDialog(release_notes, self)
            dialog.setWindowTitle("发现新版本")
            dialog.exec()

            self.selected_version = dialog.selected_version  # 存储用户选择的版本

            if self.selected_version == 'cpu' and cpu_asset_id:
                self.start_download(asset_id=cpu_asset_id, asset_name=cpu_asset_name)
            elif self.selected_version == 'gpu' and gpu_asset_id:
                self.start_download(asset_id=gpu_asset_id, asset_name=gpu_asset_name)
            else:
                self.output_text_edit.append("用户取消了更新。")
        elif latest_version:
            self.output_text_edit.append("当前已是最新版本。")
            QMessageBox.information(self, "已是最新版本", "当前已是最新版本。")
        else:
            self.output_text_edit.append("无法获取最新版本信息。")
            QMessageBox.warning(self, "更新失败", "无法获取最新版本信息。")
        self.check_update_button.setEnabled(True)

    def start_download(self, asset_id, asset_name):
        """开始下载指定的版本"""
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.quit()
            self.download_thread.wait()

        self.output_text_edit.append("开始下载更新文件...")
        self.download_progress_bar.setValue(0)
        self.download_progress_bar.setVisible(True)
        self.download_progress_bar.setFormat("下载进度：%p%")
        self.check_update_button.setEnabled(False)  # 禁用“检查版本更新”按钮

        self.download_thread = QThread()
        self.download_worker = DownloadWorker(
            asset_id=asset_id,
            owner=OWNER,
            repo=REPO,
            token=self.ComplicanceToolbox_Update_Token
        )
        self.download_worker.moveToThread(self.download_thread)

        self.download_thread.started.connect(self.download_worker.run)
        self.download_worker.progress.connect(self.report_download_progress)
        self.download_worker.finished.connect(lambda success, path: self.on_download_finished(success, path))
        self.download_worker.finished.connect(self.download_worker.deleteLater)
        self.download_thread.finished.connect(self.download_thread.deleteLater)

        self.download_thread.start()

    def report_download_progress(self, percent):
        """仅更新下载进度条"""
        self.download_progress_bar.setValue(percent)

    def on_download_finished(self, success, file_path):
        self.download_progress_bar.setVisible(False)
        self.download_progress_bar.setValue(0)
        if success:
            self.output_text_edit.append(f"下载完成，文件已保存到 {file_path}")
            QMessageBox.information(
                self,
                "下载完成",
                f"最新版本压缩包已下载到：\n{file_path}\n\n请关闭软件，解压并安装新版本。"
            )
            # 自动打开下载文件夹
            download_folder = os.path.dirname(file_path)
            QDesktopServices.openUrl(QUrl.fromLocalFile(download_folder))
        else:
            self.output_text_edit.append(f"下载失败：{file_path}")
            QMessageBox.warning(self, "下载失败", f"下载最新版本时出错：{file_path}\n请检查网络连接或文件写入权限。")
        self.check_update_button.setEnabled(True)
