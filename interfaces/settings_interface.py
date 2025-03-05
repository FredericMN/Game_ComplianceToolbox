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
import time
import json
import requests

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
        self.setFixedSize(500, 400)  # 稍微增加高度以更好地显示内容
        layout = QVBoxLayout()
        layout.setSpacing(15)  # 增加间距

        # 更新内容显示
        release_label = QLabel("更新信息：")
        release_label.setWordWrap(True)
        release_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(release_label)

        self.release_text = QTextEdit()
        self.release_text.setReadOnly(True)
        self.release_text.setPlainText(release_notes if release_notes else "无更新说明。")
        self.release_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        layout.addWidget(self.release_text)

        # 版本选择说明
        description_label = QLabel("请选择要下载的版本：")
        description_label.setWordWrap(True)
        description_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 12px;
                margin-top: 10px;
            }
        """)
        layout.addWidget(description_label)

        # GPU版本说明
        gpu_note = QLabel("* GPU版本：推荐配置为英伟达20系及以上显卡，且已安装最新显卡驱动")
        gpu_note.setWordWrap(True)
        gpu_note.setStyleSheet("color: #666666; font-size: 12px;")
        layout.addWidget(gpu_note)

        # CPU版本说明
        cpu_note = QLabel("* CPU版本：适用于所有用户，性能略低于GPU版本")
        cpu_note.setWordWrap(True)
        cpu_note.setStyleSheet("color: #666666; font-size: 12px;")
        layout.addWidget(cpu_note)

        # 按钮区域
        self.button_box = QDialogButtonBox()
        self.cpu_button = QPushButton("下载CPU版")
        self.gpu_button = QPushButton("下载GPU版")
        self.cancel_button = QPushButton("取消更新")
        
        # 设置按钮样式
        button_style = """
            QPushButton {
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """
        
        # 为每个按钮设置不同的样式
        self.cpu_button.setStyleSheet(button_style + """
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #2473a7;
            }
        """)
        
        self.gpu_button.setStyleSheet(button_style + """
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:pressed {
                background-color: #219a52;
            }
        """)
        
        self.cancel_button.setStyleSheet(button_style + """
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:pressed {
                background-color: #a93226;
            }
        """)

        # 设置按钮大小
        for btn in [self.cpu_button, self.gpu_button, self.cancel_button]:
            btn.setMinimumWidth(120)
            btn.setCursor(Qt.PointingHandCursor)  # 添加鼠标悬停效果
        
        self.button_box.addButton(self.cpu_button, QDialogButtonBox.ActionRole)
        self.button_box.addButton(self.gpu_button, QDialogButtonBox.ActionRole)
        self.button_box.addButton(self.cancel_button, QDialogButtonBox.RejectRole)

        # 设置按钮布局
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.button_box)
        button_layout.setAlignment(Qt.AlignCenter)  # 居中对齐按钮
        layout.addLayout(button_layout)

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
        self.download_worker = None  # 添加 download_worker 属性
        self.selected_version = None  # 'cpu' 或 'gpu'
        self.download_button_box = QDialogButtonBox()
        self._download_start_time = None
        self._last_update_time = None
        self._last_bytes = 0
        
        self.init_ui()
        self.load_local_update_logs()  # 加载本地更新日志

    def init_ui(self):
        # 创建主垂直布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # 创建"检查版本更新"按钮的水平布局
        update_button_layout = QHBoxLayout()
        update_button_layout.setAlignment(Qt.AlignLeft)

        # 创建"检查版本更新"按钮
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
        # 在开始前清理旧线程
        if self.update_thread:
            self.update_thread.quit()
            self.update_thread.wait()
            self.update_thread = None

        self.check_update_button.setEnabled(False)
        self.output_text_edit.append("正在检查更新...")

        self.update_thread = QThread()
        self.version_checker = VersionChecker()
        self.worker = VersionCheckWorker(self.version_checker)
        self.worker.moveToThread(self.update_thread)

        self.update_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_update_check_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.finished.connect(self.update_thread.quit)  # 确保线程退出
        self.update_thread.finished.connect(self.update_thread.deleteLater)
        self.update_thread.finished.connect(lambda: self._clear_thread_reference('update_thread'))

        self.worker.progress.connect(self.report_progress)

        self.update_thread.start()

    def _clear_thread_reference(self, thread_name):
        """线程清理时断开所有信号"""
        thread = getattr(self, thread_name)
        if thread:
            try:
                thread.disconnect()
            except:
                pass
        setattr(self, thread_name, None)

    def report_progress(self, message):
        self.output_text_edit.append(message)

    def on_update_check_finished(self, is_new_version, latest_version, cpu_download_url, gpu_download_url, release_notes):
        """处理版本检测完成后的逻辑"""
        if is_new_version:
            # 使用自定义对话框让用户选择下载版本，并在对话框中显示版本信息
            version_info = f"""当前版本: {self.current_version}
最新版本: {latest_version}

更新内容：
{release_notes}"""
            
            dialog = VersionSelectionDialog(version_info, self)
            dialog.setWindowTitle("发现新版本")
            dialog.exec()

            self.selected_version = dialog.selected_version  # 存储用户选择的版本

            if self.selected_version == 'cpu' and cpu_download_url:
                self.start_download(download_url=cpu_download_url)
            elif self.selected_version == 'gpu' and gpu_download_url:
                self.start_download(download_url=gpu_download_url)
            else:
                self.output_text_edit.append("用户取消了更新。")
        elif latest_version:
            self.output_text_edit.append("当前已是最新版本。")
            QMessageBox.information(self, "已是最新版本", "当前已是最新版本。")
        else:
            self.output_text_edit.append("无法获取最新版本信息。")
            QMessageBox.warning(self, "更新失败", "无法获取最新版本信息。")
        
        self.check_update_button.setEnabled(True)

    def start_download(self, download_url):
        """开始下载指定的版本"""
        # 检查文件是否已存在
        filename = os.path.basename(download_url)
        save_path = os.path.join(os.getcwd(), filename)
        
        if os.path.exists(save_path):
            reply = QMessageBox.question(
                self,
                "文件已存在",
                f"文件 {filename} 已存在，是否重新下载？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
            try:
                os.remove(save_path)  # 删除已存在的文件
            except Exception as e:
                QMessageBox.warning(self, "错误", f"无法删除已存在的文件：{str(e)}")
                return

        # 清理之前的下载线程和工作器
        self.cleanup_download_resources()
        
        # 创建并配置新的下载工作器
        self.download_thread = QThread()
        self.download_worker = DownloadWorker(download_url)
        self.download_worker.moveToThread(self.download_thread)
        
        # 连接信号
        self.download_thread.started.connect(self.download_worker.run)
        self.download_worker.progress.connect(self.report_download_progress)
        self.download_worker.finished.connect(self.on_download_finished)
        self.download_worker.finished.connect(lambda: self.cleanup_download_resources())
        
        # 更新UI状态
        self.download_progress_bar.setValue(0)
        self.download_progress_bar.setVisible(True)
        self.check_update_button.setEnabled(False)
        
        # 显示取消按钮
        self.download_button_box.clear()
        cancel_button = QPushButton("取消下载")
        cancel_button.clicked.connect(self.cancel_download)
        self.download_button_box.addButton(cancel_button, QDialogButtonBox.RejectRole)
        
        # 开始下载
        self.download_thread.start()
        self.output_text_edit.append(f"开始下载 {filename}...")

    def cleanup_download_resources(self):
        """清理下载相关资源"""
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.quit()
            self.download_thread.wait()
        
        if self.download_worker:
            try:
                self.download_worker.disconnect()
            except:
                pass
            self.download_worker = None
            
        if self.download_thread:
            try:
                self.download_thread.disconnect()
            except:
                pass
            self.download_thread = None
        
        # 重置下载状态
        self._download_start_time = None
        self._last_update_time = None
        self._last_bytes = 0

    def cancel_download(self):
        """取消下载"""
        if self.download_worker:
            self.download_worker.cancel()
            self.output_text_edit.append("正在取消下载...")
            self.download_progress_bar.setFormat("正在取消...")
            # 清理会在 on_download_finished 中完成

    def report_download_progress(self, percent, message):
        """更新下载进度"""
        self.download_progress_bar.setValue(percent)
        
        current_time = time.time()
        if not self._download_start_time:
            self._download_start_time = current_time
            self._last_update_time = current_time
            self._last_bytes = 0
            
        if current_time - self._last_update_time >= 0.5:  # 每0.5秒更新一次速度
            if hasattr(self.download_worker, 'downloaded_bytes'):
                bytes_diff = self.download_worker.downloaded_bytes - self._last_bytes
                time_diff = current_time - self._last_update_time
                speed = bytes_diff / time_diff
                
                # 格式化速度显示
                if speed < 1024 * 1024:  # < 1MB/s
                    speed_str = f"{speed/1024:.1f} KB/s"
                else:  # >= 1MB/s
                    speed_str = f"{speed/1024/1024:.1f} MB/s"
                
                self.download_progress_bar.setFormat(f"下载进度：%p% ({speed_str})")
                self._last_update_time = current_time
                self._last_bytes = self.download_worker.downloaded_bytes
        
        if message:
            self.output_text_edit.append(message)

    def on_download_finished(self, success, file_path):
        # 清理临时文件
        if not success and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                self.output_text_edit.append(f"清理临时文件失败: {str(e)}")
        
        # 重置下载状态
        self._download_start_time = None
        self._last_update_time = None
        self._last_bytes = 0

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
            self.output_text_edit.append("下载失败，请检查网络连接或文件写入权限。")
            QMessageBox.warning(self, "下载失败", "下载最新版本时出错，请检查网络连接或文件写入权限。")
        self.check_update_button.setEnabled(True)

    def verify_download(self, file_path, expected_size=None):
        """验证下载文件的完整性"""
        if not os.path.exists(file_path):
            return False
        if expected_size and os.path.getsize(file_path) != expected_size:
            return False
        return True

    def save_download_progress(self, url, downloaded_bytes):
        """保存下载进度"""
        try:
            with open('.download_progress', 'w') as f:
                json.dump({
                    'url': url,
                    'bytes': downloaded_bytes,
                    'time': time.time()
                }, f)
        except:
            pass

    def check_network(self):
        """检查网络连接状态"""
        try:
            requests.get('https://www.baidu.com', timeout=3)
            return True
        except:
            return False
