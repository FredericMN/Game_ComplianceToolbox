# interfaces/base_interface.py

from PySide6.QtWidgets import QWidget, QVBoxLayout

class BaseInterface(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        # 不在这里调用 self.init_ui()
