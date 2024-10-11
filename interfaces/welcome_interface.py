# welcome_interface.py

import os
import sys
import json
from datetime import datetime
from PySide6.QtWidgets import QLabel, QVBoxLayout, QHBoxLayout, QWidget, QFrame, QPushButton, QTextEdit
from PySide6.QtCore import Qt, QThread, Signal, QStandardPaths
from .base_interface import BaseInterface
from PySide6.QtGui import QFont
from utils.environment_checker import EnvironmentChecker

class WelcomeInterface(BaseInterface):
    """欢迎页界面"""
    # 定义信号
    environment_check_started = Signal()
    environment_check_finished = Signal(bool)  # 参数表示是否存在错误

    def __init__(self, parent=None):
        super().__init__(parent)
        self.env_result_file = self.get_env_result_file_path()
        self.init_ui()  # 初始化 UI
        self.check_env_status()  # 检查环境状态

    def get_env_result_file_path(self):
        """获取环境检测结果文件的绝对路径"""
        # 尝试使用 QStandardPaths 获取标准配置目录
        config_dir = QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation)
        if not config_dir:
            # 如果无法获取配置目录，退回到当前文件所在目录
            config_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 确定当前运行的版本名称
        executable_name = os.path.basename(sys.argv[0]).lower()
        if 'cuda' in executable_name:
            version_suffix = 'cuda'
        else:
            version_suffix = 'standard'
        
        # 设置应用程序特定的目录
        app_specific_dir = os.path.join(config_dir, f"ComplianceToolbox_{version_suffix}")
        if not os.path.exists(app_specific_dir):
            os.makedirs(app_specific_dir)
        
        # 返回环境检测结果文件的路径
        return os.path.join(app_specific_dir, 'env_check_result.json')

    def check_env_status(self):
        """检查是否已记录环境检测结果"""
        if os.path.exists(self.env_result_file):
            try:
                with open(self.env_result_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                date_str = data.get('date')
                result = data.get('result')
                if date_str and (result is not None):
                    if result:
                        message = f"{date_str} 检测运行环境为通过，可直接使用，如遇问题可再次检测！"
                    else:
                        message = f"{date_str} 检测运行环境为不通过，建议再次检测或获取帮助！"
                    self.output_text_edit.append(message)
                else:
                    # 如果文件内容不完整，重新进行检测
                    self.run_environment_check()
            except Exception as e:
                # 如果读取文件出错，重新进行检测
                self.output_text_edit.append(f"读取检测结果失败，重新检测。错误信息：{str(e)}")
                self.run_environment_check()
        else:
            # 如果文件不存在，进行环境检测
            self.run_environment_check()

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
            {"name": "大模型语意分析", "description": "通过大模型判断文本内容态度。"},
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
        self.check_env_button.setFixedHeight(40)  # 增加按钮高度
        self.check_env_button.setFixedWidth(200)   # 设置按钮宽度，不横跨整个界面
        # 设置按钮字体为加粗
        font = QFont()
        font.setBold(True)
        self.check_env_button.setFont(font)

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
        # 获取当前日期
        current_date = datetime.now().strftime("%Y-%m-%d")
        # 记录结果到文件
        self.record_env_check_result(current_date, not has_errors)
        # 根据检测结果给出提示
        if has_errors:
            message = f"{current_date} 检测运行环境为不通过，建议再次检测或获取帮助！"
            self.append_output("环境检测过程中存在问题，请根据提示进行处理。")
        else:
            message = f"{current_date} 检测运行环境为通过，可直接使用，如遇问题可再次检测！"
            self.append_output("恭喜，环境检测和配置完成！")
        # 在信息输出区域显示提示
        self.output_text_edit.append(message)
        # 线程清理
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.environment_checker = None

    def record_env_check_result(self, date_str, result):
        """记录环境检测结果到文件"""
        data = {
            "date": date_str,
            "result": result  # True 表示通过，False 表示不通过
        }
        try:
            with open(self.env_result_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.append_output(f"记录环境检测结果失败：{str(e)}")
