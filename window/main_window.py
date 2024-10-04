# project-01/window/main_window.py

from PySide6.QtWidgets import QApplication, QHBoxLayout, QStackedWidget, QMessageBox  # 添加 QMessageBox
from PySide6.QtGui import QIcon
from qfluentwidgets import (
    NavigationInterface, NavigationItemPosition, FluentIcon as FIF
)
from qframelesswindow import FramelessWindow, StandardTitleBar
from PySide6.QtGui import QIcon
import os
import sys
from interfaces.welcome_interface import WelcomeInterface
from interfaces.detection_tool_interface import DetectionToolInterface
from interfaces.crawler_interface import CrawlerInterface
from interfaces.empty_interface import EmptyInterface
from interfaces.settings_interface import SettingsInterface
from interfaces.version_matching_interface import VersionMatchingInterface
from interfaces.vocabulary_comparison_interface import VocabularyComparisonInterface  # 新增导入

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

        # 主布局
        self.hBoxLayout = QHBoxLayout(self)
        self.navigationInterface = NavigationInterface(self, showMenuButton=True)
        self.stackWidget = QStackedWidget(self)

        # 创建界面
        self.welcomeInterface = WelcomeInterface(self)
        self.detectionToolInterface = DetectionToolInterface(self)
        self.crawlerInterface = CrawlerInterface(self)
        self.vocabularyComparisonInterface = VocabularyComparisonInterface(self)  # 实例化新的界面
        self.developingInterface = EmptyInterface(parent=self)  # 修改此行
        self.settingsInterface = SettingsInterface(self)
        self.versionMatchingInterface = VersionMatchingInterface(self)

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
            self.vocabularyComparisonInterface, FIF.DOCUMENT, "词表对照"  # 使用合适的图标，如 LIST
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
