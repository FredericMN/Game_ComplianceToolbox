# interfaces/welcome_interface.py

from PySide6.QtWidgets import QLabel, QVBoxLayout, QHBoxLayout, QWidget, QFrame, QPushButton, QTextEdit
from PySide6.QtCore import Qt, QThread, Signal, QRectF, QPointF
from .base_interface import BaseInterface
from PySide6.QtGui import QFont, QColor, QPainter, QPixmap
from utils.environment_checker import EnvironmentChecker

class WelcomeInterface(BaseInterface):
    """欢迎页界面"""
    # 定义信号
    environment_check_started = Signal()
    environment_check_finished = Signal(bool)  # 参数表示是否存在错误

    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()  # 添加这行，初始化 UI

    def init_ui(self):
        # 主垂直布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # 顶部区域
        top_widget = QWidget()
        top_layout = QVBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)

        # 欢迎标题
        welcome_label = QLabel("欢迎使用合规工具箱", top_widget)
        welcome_font = QFont("Arial", 24, QFont.Bold)
        welcome_label.setFont(welcome_font)
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_label.setStyleSheet("color: #333333;")  # 设置字体颜色

        top_layout.addWidget(welcome_label)
        top_widget.setLayout(top_layout)

        # 功能介绍区域
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout()
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(10)

        # 功能简介
        functions = [
            {"name": "文档风险词汇批量检测", "description": "检测并标记文档中的风险词汇。"},
            {"name": "新游爬虫", "description": "爬取TapTap上的新游信息并匹配版号。"},
            {"name": "版号匹配", "description": "匹配游戏的版号信息。"},
            {"name": "词表对照", "description": "对照两个词表的差异。"},
            {"name": "设定", "description": "配置工具的相关设置。"}
        ]

        for func in functions:
            func_layout = QHBoxLayout()
            func_layout.setContentsMargins(0, 0, 0, 0)
            func_layout.setSpacing(10)

            # 功能名称
            name_label = QLabel(func["name"])
            name_font = QFont("Arial", 12, QFont.Bold)
            name_label.setFont(name_font)
            name_label.setStyleSheet("color: #555555;")

            # 功能描述
            desc_label = QLabel(func["description"])
            desc_font = QFont("Arial", 12)
            desc_label.setFont(desc_font)
            desc_label.setStyleSheet("color: #777777;")
            desc_label.setWordWrap(True)

            func_layout.addWidget(name_label)
            func_layout.addWidget(desc_label)
            func_layout.setStretch(0, 1)
            func_layout.setStretch(1, 3)

            bottom_layout.addLayout(func_layout)

        bottom_widget.setLayout(bottom_layout)

        # 将顶部区域和功能介绍添加到主布局
        main_layout.addWidget(top_widget)
        main_layout.addWidget(bottom_widget)

        # 添加分割线
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setFrameShadow(QFrame.Sunken)
        separator1.setStyleSheet("color: #CCCCCC;")
        separator1.setFixedHeight(2)
        main_layout.addWidget(separator1)

        # 检测并配置运行环境按钮和说明
        env_layout = QHBoxLayout()
        self.check_env_button = QPushButton("检测并配置运行环境")
        self.check_env_button.clicked.connect(self.run_environment_check)
        description_label = QLabel("每次运行软件时会自动检测运行环境,需安装微软Edge浏览器。")
        description_label.setWordWrap(True)
        env_layout.addWidget(self.check_env_button)
        env_layout.addWidget(description_label)
        env_layout.addStretch()

        main_layout.addLayout(env_layout)

        # 信息输出区域
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("信息输出区域")
        main_layout.addWidget(self.output_text_edit)

        # 将主布局添加到界面
        self.layout.addLayout(main_layout)

        # 创建遮罩层
        self.overlay = QWidget(self)
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 150);")
        self.overlay_layout = QVBoxLayout(self.overlay)
        self.overlay_layout.setAlignment(Qt.AlignCenter)
        self.overlay_label = QLabel("正在检测运行环境，请稍候...")
        self.overlay_label.setStyleSheet("color: white; font-size: 24px;")
        self.overlay_layout.addWidget(self.overlay_label)
        self.overlay.hide()

    def resizeEvent(self, event):
        """在窗口大小改变时，调整遮罩层的大小"""
        super().resizeEvent(event)
        self.overlay.resize(self.size())

    def run_environment_check(self):
        """运行环境检查"""
        # 禁用按钮，防止重复点击
        self.check_env_button.setEnabled(False)
        # 显示遮罩层
        self.overlay.show()
        # Emit signal that environment check is started
        self.environment_check_started.emit()

        # 启动环境检查线程
        self.thread = QThread()
        self.environment_checker = EnvironmentChecker()
        self.environment_checker.moveToThread(self.thread)

        self.thread.started.connect(self.environment_checker.run)
        self.environment_checker.output_signal.connect(self.append_output)
        self.environment_checker.finished.connect(self.on_check_finished)
        self.environment_checker.finished.connect(self.environment_checker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def append_output(self, message):
        """在信息输出区域追加消息"""
        self.output_text_edit.append(message)

    def on_check_finished(self, has_errors):
        """环境检测完成后"""
        # 启用按钮
        self.check_env_button.setEnabled(True)
        # 隐藏遮罩层
        self.overlay.hide()
        # Emit signal that environment check is finished, passing whether there were errors
        self.environment_check_finished.emit(has_errors)
        # 根据检测结果给出提示
        if has_errors:
            self.append_output("环境检测过程中存在问题，请根据提示进行处理。")
        else:
            self.append_output("恭喜，环境检测和配置完成！")
        # 线程清理
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.environment_checker = None
