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

def kill_msedgedriver():
    """终止所有 msedgedriver.exe 进程"""
    terminated_count = 0
    failed_count = 0
    
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            # 确保proc.info['name']存在且为msedgedriver.exe
            if proc.info['name'] and proc.info['name'].lower() == 'msedgedriver.exe':
                try:
                    # 先尝试温和地终止进程
                    proc.terminate()
                    # 等待最多3秒
                    gone, alive = psutil.wait_procs([proc], timeout=3)
                    
                    # 如果进程仍然存在，强制终止
                    if proc in alive:
                        proc.kill()
                    
                    terminated_count += 1
                    print(f"已终止 msedgedriver.exe (PID: {proc.info['pid']})")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                    failed_count += 1
                    print(f"终止 msedgedriver.exe (PID: {proc.info['pid']}) 时出错: {e}")
        except Exception as e:
            print(f"处理进程信息时出错: {e}")
    
    if terminated_count > 0 or failed_count > 0:
        print(f"清理完成: 已终止 {terminated_count} 个进程, 失败 {failed_count} 个进程")
    
    return terminated_count > 0

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
