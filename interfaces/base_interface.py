from PySide6.QtWidgets import QWidget, QVBoxLayout

class BaseInterface(QWidget):
    """基础接口类，所有界面均继承自此类"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.init_ui()

    def init_ui(self):
        """初始化UI，子类重写"""
        pass
