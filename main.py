# project-01/main.py

import sys
import os
import psutil  # 导入 psutil 库
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from window.main_window import MainWindow

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def kill_msedgedriver():
    """终止所有 msedgedriver.exe 进程"""
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() == 'msedgedriver.exe':
            try:
                proc.terminate()
                proc.wait(timeout=5)
                print(f"Terminated msedgedriver.exe (PID: {proc.info['pid']})")
            except psutil.NoSuchProcess:
                print(f"Process msedgedriver.exe (PID: {proc.info['pid']}) does not exist.")
            except psutil.AccessDenied:
                print(f"Access denied when trying to terminate msedgedriver.exe (PID: {proc.info['pid']}).")
            except Exception as e:
                print(f"Failed to terminate msedgedriver.exe (PID: {proc.info['pid']}): {e}")

def main():
    app = QApplication(sys.argv)
    
    # 连接 aboutToQuit 信号以在应用退出时清理
    app.aboutToQuit.connect(kill_msedgedriver)
    
    # 设置内置主题样式
    app.setStyle("Fusion")

    window = MainWindow()
    window.setWindowIcon(QIcon(resource_path('resources/logo.ico')))  # 使用打包后资源路径
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
