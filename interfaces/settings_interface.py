# interfaces/settings_interface.py

from PySide6.QtCore import Qt, QTimer, QThread, Signal, QObject
from PySide6.QtWidgets import QLabel, QVBoxLayout, QPushButton, QFrame, QMessageBox, QTextEdit, QApplication
from .base_interface import BaseInterface
from utils.version import __version__
from utils.version_checker import VersionChecker, VersionCheckWorker, DownloadWorker
import os
import sys

class SettingsInterface(BaseInterface):
    """设定界面"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_version = __version__
        self.update_thread = None  # 重命名线程变量
        self.download_thread = None  # 初始化下载线程变量
        self.init_ui()
        self.load_local_update_logs()  # 加载本地更新日志

    def init_ui(self):
        # 创建主垂直布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # 创建“检查版本更新”按钮
        self.check_update_button = QPushButton("检查版本更新")
        self.check_update_button.setFixedHeight(40)
        self.check_update_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D7;
                color: white;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
        """)
        # 连接按钮点击信号到处理函数
        self.check_update_button.clicked.connect(self.handle_check_update)

        # 信息输出区域
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("信息输出区域")

        # 添加按钮和输出区域到主布局
        main_layout.addWidget(self.check_update_button)
        main_layout.addWidget(self.output_text_edit)

        # 创建分割线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #CCCCCC;")
        separator.setFixedHeight(2)

        # 添加分割线到主布局
        main_layout.addWidget(separator)

        # 创建开发者信息标签
        developer_label = QLabel("开发者：合规团队")
        developer_label.setAlignment(Qt.AlignCenter)
        developer_label.setStyleSheet("font-size: 14px; color: #555555;")

        # 创建当前版本标签
        version_label = QLabel(f"当前版本：{self.current_version}")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("font-size: 14px; color: #555555;")

        # 创建更新日志标签
        update_log_label = QLabel("更新日志：")
        update_log_label.setAlignment(Qt.AlignLeft)
        update_log_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333;")

        # 创建更新日志内容
        self.update_log_text_edit = QTextEdit()
        self.update_log_text_edit.setReadOnly(True)
        self.update_log_text_edit.setStyleSheet("font-size: 14px; color: #555555;")
        self.update_log_text_edit.setFixedHeight(150)

        # 添加信息标签到主布局
        main_layout.addWidget(developer_label)
        main_layout.addWidget(version_label)
        main_layout.addWidget(update_log_label)
        main_layout.addWidget(self.update_log_text_edit)

        # 添加Stretch以使内容居中
        main_layout.addStretch()

        # 将主布局添加到BaseInterface的布局中
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

    def report_progress(self, message):
        self.output_text_edit.append(message)

    def on_update_check_finished(self, is_new_version, latest_version, download_url, release_notes):
        """处理版本检测完成后的逻辑"""
        if is_new_version:
            msg = (
                f"当前版本: {self.current_version}\n"
                f"最新版本: {latest_version}\n\n"
                f"更新内容:\n{release_notes}\n\n"
                f"是否下载并更新到最新版本？"
            )
            reply = QMessageBox.question(
                self,
                "发现新版本",
                msg,
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes and download_url:
                # 开始下载

                # 如果已有下载线程在运行，先停止
                if self.download_thread and self.download_thread.isRunning():
                    self.download_thread.quit()
                    self.download_thread.wait()

                self.output_text_edit.append("开始下载最新版本...")
                self.download_thread = QThread()
                self.download_worker = DownloadWorker(download_url)
                self.download_worker.moveToThread(self.download_thread)

                self.download_thread.started.connect(self.download_worker.run)
                self.download_worker.progress.connect(self.report_progress)
                self.download_worker.finished.connect(self.on_download_finished)
                self.download_worker.finished.connect(self.download_worker.deleteLater)
                self.download_thread.finished.connect(self.download_thread.deleteLater)

                self.download_thread.start()
            else:
                self.output_text_edit.append("用户取消了更新。")
                self.check_update_button.setEnabled(True)
        elif latest_version:
            self.output_text_edit.append("当前已是最新版本。")
            self.check_update_button.setEnabled(True)
        else:
            self.output_text_edit.append("无法获取最新版本信息。")
            self.check_update_button.setEnabled(True)

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

    def update_application(self, zip_file_path):
        import zipfile
        import shutil
        import sys
        import os
        import subprocess

        # 解压缩到临时目录
        temp_dir = os.path.join(os.getcwd(), "update_temp")
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 创建更新脚本
        try:
            # 假设所有文件都在 temp_dir 下，复制到当前应用目录
            app_dir = os.getcwd()
            update_script = os.path.join(app_dir, 'update.bat')
            with open(update_script, 'w', encoding='utf-8') as f:
                f.write(f"""
            @echo off
            echo Updating, please wait...
            ping localhost -n 3 > nul
            xcopy /E /Y "{temp_dir}\\*" "{app_dir}" > nul
            rd /S /Q "{temp_dir}"
            del "{zip_file_path}"
            del "%~f0"
            start "" "{sys.executable}"
                """)
            self.output_text_edit.append("更新完成，正在重启应用...")
            # 启动更新脚本
            subprocess.Popen([update_script])
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
