# interfaces/environment_config_interface.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
import subprocess
import shutil
import sys
import os

class EnvironmentConfigWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool)  # success flag

    def __init__(self, target_dir):
        super().__init__()
        self.target_dir = target_dir

    def run(self):
        try:
            self.progress.emit("开始检测 NVIDIA GPU...")
            # 检查 NVIDIA GPU 是否存在
            if shutil.which("nvidia-smi") is None:
                self.progress.emit("未检测到 NVIDIA GPU。请确保安装了 NVIDIA 显卡和相关驱动。")
                self.finished.emit(False)
                return

            try:
                result = subprocess.run(
                    ["nvidia-smi"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=5
                )
                if result.returncode != 0:
                    self.progress.emit(f"nvidia-smi 执行失败: {result.stderr}")
                    self.finished.emit(False)
                    return
                else:
                    self.progress.emit("已检测到 NVIDIA GPU。")
            except subprocess.TimeoutExpired:
                self.progress.emit("检测 NVIDIA GPU 时超时。")
                self.finished.emit(False)
                return

            self.progress.emit("检测 CUDA 安装情况...")
            # 检查 CUDA 是否安装（通过检查 nvcc）
            if shutil.which("nvcc") is None:
                self.progress.emit("未检测到 CUDA。请安装 CUDA Toolkit 后重试。")
                self.finished.emit(False)
                return
            else:
                self.progress.emit("已检测到 CUDA 安装。")

            # 定义本地安装目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))
            local_libs = os.path.join(project_root, 'local_libs')

            if not os.path.exists(local_libs):
                os.makedirs(local_libs)
                self.progress.emit(f"已创建本地库目录: {local_libs}")
            else:
                self.progress.emit(f"本地库目录已存在: {local_libs}")

            # 定义 pip 安装命令（指定 CUDA 12.1）
            install_commands = [
                [
                    sys.executable, "-m", "pip", "install",
                    "torch==2.4.1+cu121",
                    "torchvision==0.19.1+cu121",
                    "torchaudio==2.4.1+cu121",
                    "--index-url", "https://download.pytorch.org/whl/cu121",
                    "--target", local_libs
                ]
            ]

            # 执行 pip 安装命令
            for cmd in install_commands:
                self.progress.emit(f"正在安装 {' '.join(cmd[3:7])} 到 {self.target_dir}...")
                process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True
                )
                for line in process.stdout:
                    self.progress.emit(line.strip())
                process.wait()
                if process.returncode != 0:
                    self.progress.emit(f"安装过程中出现错误，退出码 {process.returncode}。")
                    self.finished.emit(False)
                    return
                else:
                    self.progress.emit("安装完成。")

            self.progress.emit("CUDA 环境配置成功。")
            self.finished.emit(True)

        except Exception as e:
            self.progress.emit(f"发生异常: {str(e)}")
            self.finished.emit(False)

class EnvironmentConfigInterface(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setAlignment(Qt.AlignTop)
        self.setLayout(self.layout)

        # 说明标签
        explanation_text = (
            "说明：如运行该软件的电脑硬件配置包含英伟达独立显卡，可尝试进行cuda环境配置，"
            "配置后可使用GPU加速大模型运算。"
        )
        self.explanation_label = QLabel(explanation_text)
        self.explanation_label.setWordWrap(True)
        self.layout.addWidget(self.explanation_label)

        # 检测并配置按钮
        self.detect_button = QPushButton("检测并配置CUDA环境")
        self.detect_button.clicked.connect(self.handle_detect_and_configure)
        self.layout.addWidget(self.detect_button)

        # 结果输出栏
        self.output_text_edit = QTextEdit()
        self.output_text_edit.setReadOnly(True)
        self.output_text_edit.setPlaceholderText("结果输出区域")
        self.layout.addWidget(self.output_text_edit)

    def handle_detect_and_configure(self):
        self.detect_button.setEnabled(False)
        self.output_text_edit.clear()
        self.output_text_edit.append("开始检测并配置环境...")

        # 定义本地安装目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))
        local_libs = os.path.join(project_root, 'local_libs')

        # 启动工作线程
        self.thread = QThread()
        self.worker = EnvironmentConfigWorker(target_dir=local_libs)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_output)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def update_output(self, message):
        self.output_text_edit.append(message)

    def on_finished(self, success):
        if success:
            self.output_text_edit.append("环境配置完成。")
            # 提示用户重启软件以应用新的CUDA环境
            self.output_text_edit.append("请重新启动软件以应用新的CUDA环境。")
        else:
            self.output_text_edit.append("环境配置失败。请根据上述信息进行处理。")
        self.detect_button.setEnabled(True)
