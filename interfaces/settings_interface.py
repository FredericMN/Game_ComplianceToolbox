# project-01/interfaces/settings_interface.py

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QLabel, QVBoxLayout, QPushButton, QFrame, QMessageBox
from .base_interface import BaseInterface

class SettingsInterface(BaseInterface):
    """设定界面"""
    def __init__(self, parent=None):
        super().__init__(parent)

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

        # 添加按钮到主布局
        main_layout.addWidget(self.check_update_button)

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
        version_label = QLabel("当前版本：1.0.0")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet("font-size: 14px; color: #555555;")

        # 创建更新日志标签
        update_log_label = QLabel("更新日志：")
        update_log_label.setAlignment(Qt.AlignLeft)
        update_log_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #333333;")

        # 创建更新日志内容
        update_log_content = QLabel(
            "2024-10-04 发布 1.0.0 版本\n"
            "新增 文档风险词汇批量检测功能\n"
            "新增 新游爬虫功能\n"
            "新增 版号匹配功能\n"
            "新增 词表对照功能"
        )
        update_log_content.setAlignment(Qt.AlignLeft)
        update_log_content.setStyleSheet("font-size: 14px; color: #555555;")
        update_log_content.setWordWrap(True)

        # 添加信息标签到主布局
        main_layout.addWidget(developer_label)
        main_layout.addWidget(version_label)
        main_layout.addWidget(update_log_label)
        main_layout.addWidget(update_log_content)

        # 添加Stretch以使内容居中
        main_layout.addStretch()

        # 将主布局添加到BaseInterface的布局中
        self.layout.addLayout(main_layout)

    def handle_check_update(self):
        """处理检查版本更新的逻辑"""
        # 延迟一秒后显示对话框
        QTimer.singleShot(1000, self.show_update_dialog)

    def show_update_dialog(self):
        """显示版本更新对话框"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("版本更新")
        msg_box.setText("当前已是最新版本！")
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec()
