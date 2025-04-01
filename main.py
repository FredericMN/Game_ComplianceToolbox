# project-01/main.py

import sys
import os
import psutil  # 导入 psutil 库
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from window.main_window import MainWindow
# 导入任务管理器
from utils.task_manager import task_manager

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def main():
    app = QApplication(sys.argv)
    
    # 连接 aboutToQuit 信号以在应用退出时清理
    app.aboutToQuit.connect(task_manager.cleanup_all_resources)
    
    # 设置内置主题样式
    app.setStyle("Fusion")

    window = MainWindow()
    window.setWindowIcon(QIcon(resource_path('resources/logo.ico')))  # 使用打包后资源路径
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
