# project-01/main.py

import sys
import os
import psutil  # 导入 psutil 库
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from window.main_window import MainWindow
# 导入任务管理器
from utils.task_manager import task_manager
import tempfile
import glob
import shutil
import subprocess
import time

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def cleanup_webdriver_temp_files():
    """启动时清理可能遗留的WebDriver临时文件夹"""
    try:
        print("正在清理WebDriver临时文件...")
        # 清理临时目录中的所有WebDriver相关目录
        temp_dir = tempfile.gettempdir()
        patterns = ["edge_driver_*", "edge_temp_*", "edge_port_*", "edge_alt_*", "scoped_dir*"]
        removed_count = 0
        
        # 首先彻底终止所有msedgedriver进程
        try:
            # 使用taskkill强制终止所有msedgedriver进程
            subprocess.call('taskkill /f /im msedgedriver.exe', shell=True, 
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            # 等待一小段时间确保进程真正终止
            time.sleep(0.5)
            
            # 再确认一次是否有残留进程并强制终止
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info.get('name') and 'msedgedriver' in proc.info.get('name').lower():
                        proc.kill()
                except:
                    pass
        except:
            pass
            
        # 也尝试清理可能存在的Edge浏览器进程
        try:
            subprocess.call('taskkill /f /im msedge.exe', shell=True, 
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except:
            pass
        
        # 等待一段时间确保进程完全释放文件锁
        time.sleep(1)
        
        # 现在清理临时文件
        for pattern in patterns:
            for path in glob.glob(os.path.join(temp_dir, pattern)):
                try:
                    if os.path.isdir(path):
                        shutil.rmtree(path, ignore_errors=True)
                        removed_count += 1
                except Exception as e:
                    print(f"清理临时目录失败: {path}, 错误: {str(e)}")
        
        # 再次确认是否有残留的msedgedriver进程
        terminated = False
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info.get('name') and 'msedgedriver' in proc.info.get('name').lower():
                    try:
                        proc.terminate()
                        proc.wait(timeout=1)
                        terminated = True
                    except:
                        try:
                            proc.kill()
                            terminated = True
                        except:
                            pass
            except:
                pass
        
        if removed_count > 0 or terminated:
            print(f"启动清理: 已移除 {removed_count} 个临时目录，驱动进程已清理。")
    except Exception as e:
        print(f"启动清理时出错: {str(e)}")

def main():
    # 应用启动时先执行一次全局清理
    cleanup_webdriver_temp_files()
    
    # 稍等一段时间以确保资源完全释放
    time.sleep(1)
    
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
