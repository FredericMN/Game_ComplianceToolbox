# interfaces/settings_interface.py

from PySide6.QtCore import Qt, QTimer, QThread, Signal, QObject, QUrl, QSize
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QPushButton, QHBoxLayout, QFrame,
    QMessageBox, QTextEdit, QApplication, QSizePolicy, QDialog, QDialogButtonBox
)
from PySide6.QtGui import QIcon, QDesktopServices, QFont
from .base_interface import BaseInterface
from utils.version import __version__
from utils.version_checker import VersionChecker, VersionCheckWorker, DownloadWorker
import os
import sys
import py7zr  # 确保已安装 py7zr

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
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择下载版本")
        self.setModal(True)
        self.setFixedSize(400, 200)
        layout = QVBoxLayout()

        info_label = QLabel("请选择要下载的版本：")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

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
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_version = __version__
        self.update_thread = None  # 更新检查线程
        self.download_thread = None  # 下载线程
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

        # 创建 GitHub 按钮的水平布局
        github_button_layout = QHBoxLayout()
        github_button_layout.setAlignment(Qt.AlignLeft)

        # 创建 GitHub 按钮
        github_button = QPushButton("GitHub")
        github_button.setFixedHeight(40)  # 增加按钮高度
        github_button.setFixedWidth(120)  # 设置按钮宽度，不横跨整个界面
        # 设置按钮字体为加粗
        github_font = QFont()
        github_font.setBold(True)
        github_button.setFont(github_font)
        # 设置 GitHub 图标
        github_icon_path = resource_path(os.path.join('resources', 'githublogo.png'))
        if os.path.exists(github_icon_path):
            github_icon = QIcon(github_icon_path)
            github_button.setIcon(github_icon)
            github_button.setIconSize(QSize(24, 24))  # 调整图标大小为 24x24 像素
        else:
            # 如果图标未找到，保持文本显示并打印提示
            github_button.setText("GitHub")
            print(f"GitHub 图标未找到：{github_icon_path}")

        # 移除自定义样式，使用默认样式
        github_button.setStyleSheet("")

        # 连接按钮点击信号到处理函数
        github_button.clicked.connect(self.open_github_url)

        # 将按钮添加到水平布局
        github_button_layout.addWidget(github_button)

        # 将 GitHub 按钮布局添加到主布局
        main_layout.addLayout(github_button_layout)

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

    def open_github_url(self):
        github_url = QUrl("https://github.com/FredericMN/Game_ComplianceToolbox")
        if not QDesktopServices.openUrl(github_url):
            QMessageBox.warning(self, "无法打开链接", "无法在浏览器中打开 GitHub 链接。")

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

    def on_update_check_finished(self, is_new_version, latest_version, cpu_download_url, gpu_download_url, release_notes):
        """处理版本检测完成后的逻辑"""
        if is_new_version:
            msg = (
                f"当前版本: {self.current_version}\n"
                f"最新版本: {latest_version}\n\n"
                f"更新内容:\n{release_notes}\n\n"
                f"请选择下载的版本。"
            )
            # 使用自定义对话框让用户选择下载版本
            dialog = VersionSelectionDialog(self)
            dialog.setWindowTitle("发现新版本")
            dialog.exec()

            if dialog.selected_version == 'cpu' and cpu_download_url:
                self.start_download(cpu_download_url)
            elif dialog.selected_version == 'gpu' and gpu_download_url:
                self.start_download(gpu_download_url)
            else:
                self.output_text_edit.append("用户取消了更新。")
        elif latest_version:
            self.output_text_edit.append("当前已是最新版本。")
        else:
            self.output_text_edit.append("无法获取最新版本信息。")
        self.check_update_button.setEnabled(True)

    def start_download(self, download_url):
        """开始下载指定的版本"""
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.quit()
            self.download_thread.wait()

        self.output_text_edit.append("开始下载更新文件...")
        self.download_thread = QThread()
        self.download_worker = DownloadWorker(download_url)
        self.download_worker.moveToThread(self.download_thread)

        self.download_thread.started.connect(self.download_worker.run)
        self.download_worker.progress.connect(self.report_progress)
        self.download_worker.finished.connect(self.on_download_finished)
        self.download_worker.finished.connect(self.download_worker.deleteLater)
        self.download_thread.finished.connect(self.download_thread.deleteLater)

        self.download_thread.start()

    def on_download_finished(self, success, file_path):
        if success:
            self.output_text_edit.append(f"下载完成，文件已保存到 {file_path}")
            reply = QMessageBox.question(
                self,
                "下载完成",
                "下载完成，是否立即更新？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                # 开始更新
                self.output_text_edit.append("正在更新...")
                try:
                    self.update_application(file_path)
                except Exception as e:
                    QMessageBox.critical(self, "更新失败", f"更新过程中发生错误：{str(e)}")
            else:
                self.output_text_edit.append("用户取消了更新。")
        else:
            self.output_text_edit.append("下载失败，请检查网络连接。")
            QMessageBox.warning(self, "下载失败", "下载最新版本时出错，请稍后重试。")
        self.check_update_button.setEnabled(True)

    def update_application(self, downloaded_file_path):
        import shutil
        import subprocess

        # 创建临时目录
        temp_dir = os.path.join(os.getcwd(), "update_temp")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        # 根据文件类型解压
        filename = os.path.basename(downloaded_file_path)
        try:
            if filename.endswith('.zip'):
                import zipfile
                with zipfile.ZipFile(downloaded_file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
            elif filename.endswith('.7z'):
                with py7zr.SevenZipFile(downloaded_file_path, mode='r') as z:
                    z.extractall(path=temp_dir)
            else:
                raise Exception("未知的文件格式。")
        except Exception as e:
            raise Exception(f"解压缩失败：{str(e)}")

        try:
            app_dir = os.getcwd()
            update_script = os.path.join(app_dir, 'update.bat')
            executable_name = os.path.basename(sys.argv[0])
            zip_file_name = os.path.basename(downloaded_file_path)

            with open(update_script, 'w', encoding='utf-8') as f:
                f.write(f"""
@echo off
echo Updating, please wait...
:waitloop
tasklist /FI "IMAGENAME eq {executable_name}" 2>NUL | find /I /N "{executable_name}">NUL
if "%ERRORLEVEL%"=="0" (
    echo Waiting for application to close...
    timeout /t 2 > nul
    goto waitloop
)
xcopy /E /Y "{temp_dir}\\*" "{app_dir}\\" > nul
rd /S /Q "{temp_dir}"
del "{app_dir}\\{zip_file_name}"
start "" "{app_dir}\\{executable_name}"
del "%~f0"
                """)
            self.output_text_edit.append("更新完成，正在重启应用...")
            # 启动更新脚本
            subprocess.Popen([update_script], shell=True)
            # 退出当前应用
            QApplication.quit()
        except Exception as e:
            raise Exception(f"更新失败：{str(e)}")

    def restart_application(self):
        import sys
        import os
        import subprocess
        # 获取当前可执行文件路径
        executable = sys.executable
        args = sys.argv
        # 重启应用
        subprocess.Popen([executable] + args)
        # 退出当前应用
        QApplication.quit()
