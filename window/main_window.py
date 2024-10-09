# window/main_window.py

from PySide6.QtWidgets import QApplication, QHBoxLayout, QStackedWidget, QMessageBox
from PySide6.QtGui import QIcon
from qfluentwidgets import (
    NavigationInterface, NavigationItemPosition, FluentIcon as FIF
)
from qframelesswindow import FramelessWindow, StandardTitleBar
import os
import sys

# 将本地库路径添加到 sys.path 以确保本地安装的库被优先导入
local_libs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'local_libs')
if os.path.exists(local_libs_path):
    sys.path.insert(0, local_libs_path)

import torch  # 导入torch库

from interfaces.welcome_interface import WelcomeInterface
from interfaces.detection_tool_interface import DetectionToolInterface
from interfaces.crawler_interface import CrawlerInterface
from interfaces.vocabulary_comparison_interface import VocabularyComparisonInterface
from interfaces.environment_config_interface import EnvironmentConfigInterface  # 导入新界面
from interfaces.empty_interface import EmptyInterface
from interfaces.settings_interface import SettingsInterface
from interfaces.version_matching_interface import VersionMatchingInterface
from interfaces.large_model_interface import LargeModelInterface
from utils.version_checker import VersionChecker, VersionCheckWorker
from PySide6.QtCore import QThread

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MainWindow(FramelessWindow):
    def __init__(self):
        super().__init__()
        self.setTitleBar(StandardTitleBar(self))
        self.setWindowTitle("合规工具箱")
        self.setWindowIcon(QIcon(resource_path('resources/logo.ico')))  # 确保 resource_path 在此之前定义
        self.update_thread = None  # 初始化更新线程

        # 主布局
        self.hBoxLayout = QHBoxLayout(self)
        self.navigationInterface = NavigationInterface(self, showMenuButton=True)
        self.stackWidget = QStackedWidget(self)

        # 创建界面
        self.welcomeInterface = WelcomeInterface(self)
        self.detectionToolInterface = DetectionToolInterface(self)
        self.crawlerInterface = CrawlerInterface(self)
        self.vocabularyComparisonInterface = VocabularyComparisonInterface(self)  # 实例化新的界面
        self.environmentConfigInterface = EnvironmentConfigInterface(self)  # 实例化新界面
        self.developingInterface = EmptyInterface(parent=self)  # 修改此行
        self.settingsInterface = SettingsInterface(self)
        self.versionMatchingInterface = VersionMatchingInterface(self)
        self.largeModelInterface = LargeModelInterface(self)

        # 初始化布局和导航栏
        self.init_layout()
        self.init_navigation()
        self.init_window()

        # 连接 WelcomeInterface 的信号以控制导航栏
        self.welcomeInterface.environment_check_started.connect(lambda: self.set_navigation_enabled(False))
        self.welcomeInterface.environment_check_finished.connect(self.on_environment_check_finished)

        # 启动环境检测
        self.welcomeInterface.run_environment_check()

    def init_layout(self):
        self.hBoxLayout.setSpacing(0)
        self.hBoxLayout.setContentsMargins(0, self.titleBar.height(), 0, 0)
        self.hBoxLayout.addWidget(self.navigationInterface)
        self.hBoxLayout.addWidget(self.stackWidget)
        self.hBoxLayout.setStretchFactor(self.stackWidget, 1)

    def init_navigation(self):
        # 添加欢迎页为首项
        self.add_sub_interface(
            self.welcomeInterface, FIF.HOME, "欢迎页"
        )

        # 添加其他主导航项
        self.add_sub_interface(
            self.detectionToolInterface, FIF.EDIT, "文档风险词汇批量检测"
        )

        self.add_sub_interface(
            self.crawlerInterface, FIF.GAME, "新游爬虫"
        )

        self.add_sub_interface(
            self.versionMatchingInterface, FIF.CAFE, "版号匹配"
        )

        self.add_sub_interface(
            self.vocabularyComparisonInterface, FIF.DOCUMENT, "词表对照"
        )

        self.add_sub_interface(
            self.largeModelInterface, FIF.PROJECTOR, "大模型语义分析"
        )

        self.add_sub_interface(
            self.environmentConfigInterface, FIF.BUS, "GPU加速环境配置"  # 添加新导航项
        )

        self.add_sub_interface(
            self.developingInterface, FIF.CODE, "新功能开发中"
        )

        # 在导航栏底部添加“设定”项
        self.navigationInterface.addSeparator()

        self.add_sub_interface(
            self.settingsInterface, FIF.SETTING, "设定",
            position=NavigationItemPosition.BOTTOM
        )

        # 设置默认显示界面为欢迎页
        self.stackWidget.setCurrentWidget(self.welcomeInterface)

    def init_window(self):
        self.resize(900, 700)
        self.setWindowIcon(QIcon('resources/logo.png'))

        # 居中显示窗口
        desktop = QApplication.primaryScreen().availableGeometry()
        w, h = desktop.width(), desktop.height()
        self.move((w - self.width()) // 2, (h - self.height()) // 2)

    def add_sub_interface(self, interface, icon, text: str, position=NavigationItemPosition.TOP, parent=None):
        """添加子界面到导航栏"""
        self.stackWidget.addWidget(interface)
        self.navigationInterface.addItem(
            routeKey=interface.__class__.__name__,
            icon=icon,
            text=text,
            onClick=lambda: self.switch_to(interface),
            position=position,
            tooltip=text,
            parentRouteKey=parent.__class__.__name__ if parent else None
        )

    def switch_to(self, widget):
        self.stackWidget.setCurrentWidget(widget)

    def set_navigation_enabled(self, enabled: bool):
        """启用或禁用导航栏"""
        self.navigationInterface.setEnabled(enabled)

    def on_environment_check_finished(self, has_errors):
        """处理环境检测完成后的逻辑"""
        self.set_navigation_enabled(True)
        if has_errors:
            QMessageBox.warning(self, "环境检测", "环境检测过程中存在问题，请根据提示进行处理。")
        else:
            QMessageBox.information(self, "环境检测", "恭喜，环境检测和配置完成！")
        # 在环境检测完成后，启动版本检测
        self.check_for_updates()

    def check_for_updates(self):
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

        self.update_thread.start()

    def on_update_check_finished(self, is_new_version, latest_version, download_url, release_notes):
        if is_new_version:
            msg = (
                f"当前版本: {self.version_checker.current_version}\n"
                f"最新版本: {latest_version}\n\n"
                f"更新内容:\n{release_notes}\n\n"
                f"是否前往更新？"
            )
            reply = QMessageBox.question(
                self,
                "发现新版本",
                msg,
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                # 切换到设置界面
                self.switch_to(self.settingsInterface)
                # 开始下载和更新过程
                self.settingsInterface.handle_check_update()
