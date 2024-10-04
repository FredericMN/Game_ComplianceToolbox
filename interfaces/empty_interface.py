# project-01/interfaces/empty_interface.py

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QHBoxLayout
from .base_interface import BaseInterface

class EmptyInterface(BaseInterface):
    """开发中界面"""
    def __init__(self, text="新功能开发中……", parent=None):  # 修改了默认的显示文字
        self.text = text
        super().__init__(parent)

    def init_ui(self):
        #print(f"Displaying text: {self.text}")  # 调试输出
        label = QLabel(self.text, self)
        label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(label)
