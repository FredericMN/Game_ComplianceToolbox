# interfaces/welcome_interface.py

import os
import sys
import json
from datetime import datetime
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QHBoxLayout, QWidget, QFrame, QPushButton,
    QTextEdit, QTreeWidget, QTreeWidgetItem
)
from PySide6.QtCore import Qt, QThread, Signal
from .base_interface import BaseInterface
from PySide6.QtGui import QFont
from utils.environment_checker import EnvironmentChecker


class WelcomeInterface(BaseInterface):
    """欢迎页界面"""

    environment_check_started = Signal()
    environment_check_finished = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.env_result_file = self.get_env_result_file_path()
        self.init_ui()
        self.check_env_status()

    def get_current_dir(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        else:
            return os.path.dirname(os.path.abspath(__file__))

    def get_env_result_file_path(self):
        CURRENT_DIR = self.get_current_dir()
        return os.path.join(CURRENT_DIR, 'env_check_result.json')

    def check_env_status(self):
        """若已有检测结果则加载，否则执行新的检测"""
        if os.path.exists(self.env_result_file):
            try:
                with open(self.env_result_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                date_str = data.get('date')
                result = data.get('result')
                if date_str and (result is not None):
                    if result:
                        message = f"{date_str} 检测环境：通过。可直接使用。如遇问题可再次检测！"
                    else:
                        message = f"{date_str} 检测环境：不通过。建议再次检测或获取帮助！"
                    self.output_text_edit.append(message)
                else:
                    self.run_environment_check()
            except Exception as e:
                self.output_text_edit.append(f"读取检测结果失败，重新检测。错误信息：{str(e)}")
                self.run_environment_check()
        else:
            self.run_environment_check()

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # 顶部区域
        top_widget = QWidget()
        top_layout = QVBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)

        welcome_label = QLabel("欢迎使用合规工具箱", top_widget)
        welcome_font = QFont("Arial", 24, QFont.Bold)
        welcome_label.setFont(welcome_font)
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_label.setStyleSheet("color: #333333;")
        top_layout.addWidget(welcome_label)
        top_widget.setLayout(top_layout)

        # 中间功能简介
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout()
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setSpacing(10)

        functions = [
            {"name": "文档风险词汇批量检测", "description": "检测并标记文档中的风险词汇。"},
            {"name": "新游爬虫", "description": "爬取TapTap上的新游信息并匹配版号。"},
            {"name": "版号匹配", "description": "匹配游戏的版号信息。"},
            {"name": "词表对照", "description": "对照两个词表的差异。"},
            {"name": "大模型语义分析", "description": "通过大模型审核文本，标记高风险内容。"},
            {"name": "大模型文案正向优化", "description": "通过大模型输出语句的正向优化。"},
            {"name": "设定", "description": "配置工具的相关设置。"}
        ]

        for func in functions:
            func_layout = QHBoxLayout()
            func_layout.setContentsMargins(0, 0, 0, 0)
            func_layout.setSpacing(10)

            name_label = QLabel(func["name"])
            name_font = QFont("Arial", 12, QFont.Bold)
            name_label.setFont(name_font)
            name_label.setStyleSheet("color: #555555;")

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

        main_layout.addWidget(top_widget)
        main_layout.addWidget(bottom_widget)

        # 分割线
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setFrameShadow(QFrame.Sunken)
        separator1.setStyleSheet("color: #CCCCCC;")
        separator1.setFixedHeight(2)
        main_layout.addWidget(separator1)

        # 检测环境按钮和说明
        env_layout = QHBoxLayout()
        self.check_env_button = QPushButton("检测并配置运行环境")
        self.check_env_button.setFixedHeight(40)
        self.check_env_button.setFixedWidth(200)
        font = QFont()
        font.setBold(True)
        self.check_env_button.setFont(font)

        self.check_env_button.clicked.connect(self.run_environment_check)
        description_label = QLabel("每次运行软件时会自动检测运行环境，需要已安装Edge浏览器。")
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

        # 新增：使用QTreeWidget展示结构化检测结果
        self.result_tree = QTreeWidget()
        self.result_tree.setColumnCount(3)
        self.result_tree.setHeaderLabels(["检测项", "是否通过", "详情"])
        main_layout.addWidget(self.result_tree)

        # 遮罩层
        self.overlay = QWidget(self)
        self.overlay.setStyleSheet("background-color: rgba(0, 0, 0, 150);")
        self.overlay_layout = QVBoxLayout(self.overlay)
        self.overlay_layout.setAlignment(Qt.AlignCenter)
        self.overlay_label = QLabel("正在检测运行环境，请稍候...")
        self.overlay_label.setStyleSheet("color: white; font-size: 24px;")
        self.overlay_layout.addWidget(self.overlay_label)
        self.overlay.hide()

        self.layout.addLayout(main_layout)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.overlay.resize(self.size())

    def run_environment_check(self):
        """执行环境检测"""
        self.check_env_button.setEnabled(False)
        self.overlay.show()
        self.environment_check_started.emit()

        self.thread = QThread()
        self.environment_checker = EnvironmentChecker()
        self.environment_checker.moveToThread(self.thread)

        self.thread.started.connect(self.environment_checker.run)
        self.environment_checker.output_signal.connect(self.append_output)
        # 结构化结果信号
        self.environment_checker.structured_result_signal.connect(self.on_structured_results)
        self.environment_checker.finished.connect(self.on_check_finished)
        self.environment_checker.finished.connect(self.environment_checker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def append_output(self, message):
        self.output_text_edit.append(message)

    def on_structured_results(self, results):
        """将结果填入QTreeWidget"""
        self.result_tree.clear()
        for item_name, status_bool, detail in results:
            # 每个检测项生成一个QTreeWidgetItem
            row = QTreeWidgetItem(self.result_tree)
            row.setText(0, item_name)
            row.setText(1, "通过" if status_bool else "未通过")
            row.setText(2, detail if detail else "")
        self.result_tree.expandAll()

    def on_check_finished(self, has_errors):
        self.check_env_button.setEnabled(True)
        self.overlay.hide()
        self.environment_check_finished.emit(has_errors)

        # 获取当前日期
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.record_env_check_result(current_date, not has_errors)

        if has_errors:
            self.output_text_edit.append("环境检测存在问题，请根据提示进行处理。")
        else:
            self.output_text_edit.append("恭喜，环境检测和配置完成！")

        # 线程清理
        self.thread.quit()
        self.thread.wait()
        self.thread = None
        self.environment_checker = None

    def record_env_check_result(self, date_str, result):
        data = {
            "date": date_str,
            "result": result
        }
        try:
            with open(self.env_result_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.output_text_edit.append(f"记录环境检测结果失败：{str(e)}")
