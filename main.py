# main.py

import sys
import os
import psutil  # 导入 psutil 库
import zipfile
import shutil  # 导入 shutil 用于文件操作
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from window.main_window import MainWindow
from window.initialization_window import InitializationWindow  # 导入 InitializationWindow
from PySide6.QtCore import Qt

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

def load_torch_cuda():
    """加载 CUDA 支持的 torch 和 torchvision，如果存在压缩包的话"""
    torch_cuda_zip = 'torch_cuda_package.zip'
    torch_cuda_dir = 'torch_cuda_package'

    if os.path.exists(torch_cuda_zip):
        # 解压缩到指定目录
        if not os.path.exists(torch_cuda_dir):
            try:
                with zipfile.ZipFile(torch_cuda_zip, 'r') as zip_ref:
                    zip_ref.extractall(torch_cuda_dir)
                print("成功解压 torch_cuda_package.zip")
            except zipfile.BadZipFile:
                print("压缩包损坏，无法解压。将使用 CPU 版本的 torch。")
                return False

        # 添加解压目录及其子目录到 PATH 环境变量
        for root, dirs, files in os.walk(torch_cuda_dir):
            for dir in dirs:
                lib_path = os.path.join(root, dir)
                os.environ['PATH'] = lib_path + ";" + os.environ.get('PATH', '')
        
        # 将解压后的目录添加到 sys.path
        sys.path.insert(0, os.path.abspath(torch_cuda_dir))

        # 动态导入 torch 和 torchvision
        try:
            import torch
            import torchvision
            # 确认 CUDA 可用
            if torch.cuda.is_available():
                print("已加载 CUDA 版本的 torch 和 torchvision。")
                # 将导入的模块赋值到 sys.modules 中，确保其他模块导入时使用的是 CUDA 版本
                sys.modules['torch'] = torch
                sys.modules['torchvision'] = torchvision
                return True
            else:
                print("CUDA 不可用，使用 CPU 版本的 torch。")
                return False
        except ImportError as e:
            print(f"从 torch_cuda_package 加载 torch 或 torchvision 失败: {e}")
            return False
    else:
        print("未找到 torch_cuda_package.zip，使用已安装的 CPU 版本的 torch。")
        return False

def main():
    app = QApplication(sys.argv)
    
    # 创建并显示初始化窗口
    init_window = InitializationWindow()
    init_window.show()

    # 进行环境检查和加载 torch
    from window.main_window import MainWindow  # 确保 MainWindow 被正确导入

    # 尝试加载 CUDA 版本的 torch 和 torchvision
    gpu_available = load_torch_cuda()

    # 如果 CUDA 加载失败，使用 CPU 版本的 torch
    if not gpu_available:
        try:
            import torch
            print("已加载 CPU 版本的 torch。")
            sys.modules['torch'] = torch
        except ImportError as e:
            print(f"加载 CPU 版本的 torch 失败: {e}")
            sys.exit(1)  # 退出程序，因为没有可用的 torch

    # 关闭初始化窗口
    init_window.close()
    
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
