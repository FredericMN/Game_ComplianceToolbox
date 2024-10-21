# interfaces/large_model_optimization_interface.py

from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtWidgets import (
    QLabel, QVBoxLayout, QTextEdit, QPushButton, QHBoxLayout, QMessageBox
)
from .base_interface import BaseInterface
import json
import os
import sys
from zhipuai import ZhipuAI

def resource_path(relative_path):
    """获取资源文件的绝对路径，兼容开发和打包后的环境"""
    try:
        # PyInstaller 创建临时文件夹，并把路径存储在 _MEIPASS 中
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class OptimizationWorker(QObject):
    """工作线程，用于调用大模型API"""
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, user_input, api_key, model_name, system_prompt):
        super().__init__()
        self.user_input = user_input
        self.api_key = api_key
        self.model_name = model_name
        self.system_prompt = system_prompt

    def run(self):
        try:
            client = ZhipuAI(api_key=self.api_key)
            response = client.chat.completions.create(
                model=self.model_name,
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": self.user_input},
                ],
            )
            # 修改这里，使用 .content 获取内容
            optimized_text = response.choices[0].message.content
            self.finished.emit(optimized_text)
        except Exception as e:
            self.error.emit(str(e))

class LargeModelOptimizationInterface(BaseInterface):
    """大模型文本正向优化界面"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.load_system_prompt()

    def init_ui(self):
        # 创建布局
        self.layout.setAlignment(Qt.AlignTop)
        main_layout = QHBoxLayout()

        # 左侧：用户输入
        left_layout = QVBoxLayout()
        input_label = QLabel("输入文本：")
        self.input_text_edit = QTextEdit()
        left_layout.addWidget(input_label)
        left_layout.addWidget(self.input_text_edit)

        # 优化按钮
        self.optimize_button = QPushButton("优化")
        self.optimize_button.clicked.connect(self.handle_optimize)
        left_layout.addWidget(self.optimize_button)

        # 右侧：输出
        right_layout = QVBoxLayout()
        output_label = QLabel("优化后的文本：")
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        right_layout.addWidget(output_label)
        right_layout.addWidget(self.output_text_edit)

        # 将布局添加到主布局
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)
        self.layout.addLayout(main_layout)

    def load_system_prompt(self):
        # 使用 resource_path 来定位提示词文件
        prompt_file = resource_path(os.path.join('resources', 'optimization_prompt.json'))
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            print("从 JSON 文件加载的数据：", data)  # 调试输出
            self.system_prompt = data.get('system_prompt', '')
            print("提取的 system_prompt：", self.system_prompt)  # 调试输出
            if not self.system_prompt:
                QMessageBox.warning(self, "警告", "系统提示词为空，请检查配置文件。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载系统提示词失败：{str(e)}")
            self.system_prompt = ''

    def handle_optimize(self):
        user_input = self.input_text_edit.toPlainText().strip()
        if not user_input:
            QMessageBox.warning(self, "提示", "请输入要优化的文本。")
            return

        if not self.system_prompt:
            QMessageBox.warning(self, "提示", "系统提示词未加载，无法进行优化。")
            return

        # 禁用按钮，防止重复点击
        self.optimize_button.setEnabled(False)
        self.output_text_edit.clear()

        # 启动工作线程
        self.thread = QThread()
        api_key = '5400397f85dcc13ed9a998b8f4d6468f.ixdbnojVlZJSxASg'  # 请确保在环境变量中设置了您的 API 密钥
        if not api_key:
            QMessageBox.critical(self, "错误", "API 密钥未设置，请配置环境变量 ZHIPUAI_API_KEY。")
            self.optimize_button.setEnabled(True)
            return

        model_name = 'glm-4-flash'  # 请替换为您要使用的模型名称
        self.worker = OptimizationWorker(user_input, api_key, model_name, self.system_prompt)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_optimization_finished)
        self.worker.error.connect(self.on_optimization_error)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_optimization_finished(self, optimized_text):
        self.output_text_edit.setPlainText(optimized_text)
        self.optimize_button.setEnabled(True)
        self.thread.quit()

    def on_optimization_error(self, error_message):
        QMessageBox.critical(self, "错误", f"优化过程中发生错误：{error_message}")
        self.optimize_button.setEnabled(True)
        self.thread.quit()
