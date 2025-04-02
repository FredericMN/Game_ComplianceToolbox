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
    """启动时清理可能遗留的WebDriver临时文件夹和驱动进程"""
    try:
        print("正在清理WebDriver临时文件和相关驱动进程...")
        # 清理临时目录中的所有WebDriver相关目录
        temp_dir = tempfile.gettempdir()
        patterns = ["edge_driver_*", "edge_temp_*", "edge_port_*", "edge_alt_*", "scoped_dir*"]
        removed_count = 0
        failed_count = 0
        
        # 第一阶段：彻底终止所有相关的驱动进程
        terminated_processes = set()
        
        # 1. 使用taskkill强制终止所有 msedgedriver 进程 (Windows特定)
        #    注意：不再终止 msedge.exe
        if sys.platform == 'win32':
            for process_name in ['msedgedriver.exe', 'EdgeWebDriver.exe']: # 移除了 'msedge.exe'
                try:
                    result = subprocess.run(
                        f'taskkill /f /im {process_name}', 
                        shell=True, 
                        stdout=subprocess.DEVNULL, 
                        stderr=subprocess.DEVNULL,
                        timeout=3  # 设置超时
                    )
                    if result.returncode == 0:
                        terminated_processes.add(process_name)
                except subprocess.TimeoutExpired:
                    print(f"终止进程 {process_name} 超时")
                except Exception as e:
                    pass
            
        # 2. 使用psutil更精确地查找和终止相关的驱动进程
        #    注意：不再查找 msedge 关键字
        process_keywords = ['msedgedriver', 'edgewebdriver'] # 移除了 'msedge'
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # 确保 proc.info['name'] 存在且包含关键字
                if proc.info.get('name') and any(keyword in proc.info['name'].lower() for keyword in process_keywords):
                    try:
                        proc_name = proc.info.get('name')
                        proc.terminate()  # 先尝试温和终止
                        try:
                            # 等待进程终止，最多2秒
                            gone, alive = psutil.wait_procs([proc], timeout=2)
                            if proc in alive:
                                proc.kill()  # 如果还活着，强制终止
                            
                            terminated_processes.add(proc_name)
                        except Exception:
                            proc.kill()  # 如果等待出错，直接强制终止
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        pass
            except Exception:
                pass
        
        # 等待一段时间确保进程完全释放文件锁
        if terminated_processes:
            print(f"已终止以下驱动进程: {', '.join(terminated_processes)}")
            time.sleep(1.5)  # 增加等待时间，确保资源完全释放
        
        # 第二阶段：智能清理临时文件
        # 1. 收集所有需要清理的目录，先按修改时间排序
        to_clean = []
        for pattern in patterns:
            for path in glob.glob(os.path.join(temp_dir, pattern)):
                try:
                    if os.path.isdir(path):
                        # 获取目录修改时间
                        mtime = os.path.getmtime(path)
                        to_clean.append((path, mtime))
                except Exception:
                    pass
        
        # 按修改时间从旧到新排序
        to_clean.sort(key=lambda x: x[1])
        
        # 2. 智能清理各个目录
        for path, _ in to_clean:
            try:
                # 检查目录是否可访问
                if not os.access(path, os.W_OK):
                    print(f"无法访问目录: {path}，跳过清理")
                    failed_count += 1
                    continue
                
                # 尝试使用安全的方式删除
                is_locked = False
                
                # 先检查目录中是否有锁定的文件
                for root, dirs, files in os.walk(path, topdown=False):
                    for name in files:
                        file_path = os.path.join(root, name)
                        try:
                            # 尝试以写入模式打开文件，检查是否被锁定
                            with open(file_path, 'a+'):
                                pass
                        except PermissionError:
                            # 文件被锁定
                            is_locked = True
                            break
                
                if is_locked:
                    print(f"目录 {path} 包含锁定的文件，尝试强制删除")
                    # 使用系统命令强制删除 (Windows特定)
                    if sys.platform == 'win32':
                        try:
                            subprocess.run(
                                f'rmdir /s /q "{path}"', 
                                shell=True, 
                                stdout=subprocess.DEVNULL, 
                                stderr=subprocess.DEVNULL,
                                timeout=3
                            )
                            # 检查目录是否已被删除
                            if not os.path.exists(path):
                                removed_count += 1
                            else:
                                failed_count += 1
                        except Exception:
                            failed_count += 1
                    else:
                        # 非Windows系统，使用shutil
                        shutil.rmtree(path, ignore_errors=True)
                        if not os.path.exists(path):
                            removed_count += 1
                        else:
                            failed_count += 1
                else:
                    # 目录未锁定，使用标准方法删除
                    shutil.rmtree(path, ignore_errors=True)
                    if not os.path.exists(path):
                        removed_count += 1
                    else:
                        failed_count += 1
            except Exception as e:
                print(f"清理目录失败: {path}, 错误: {str(e)}")
                failed_count += 1
        
        # 最后检查：再次确认是否有残留的驱动进程
        # 注意：不再检查 msedge.exe
        found_after_cleanup = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info.get('name') and any(keyword in proc.info.get('name').lower() for keyword in process_keywords):
                    found_after_cleanup.append(proc.info.get('name', f"PID-{proc.pid}"))
            except:
                pass
        
        if found_after_cleanup:
            print(f"警告: 清理后仍有以下驱动进程: {', '.join(found_after_cleanup)}")
        
        # 输出清理结果
        if removed_count > 0 or failed_count > 0:
            print(f"启动清理结果: 成功移除 {removed_count} 个临时目录, {failed_count} 个失败")
    except Exception as e:
        print(f"启动清理时出错: {str(e)}")
        # 记录异常但不影响程序启动

def check_system_resources():
    """检查系统资源状态，确保有足够资源运行应用"""
    try:
        print("正在检查系统资源...")
        warnings = []
        
        # 检查内存使用情况
        mem = psutil.virtual_memory()
        if mem.percent > 95:  # 如果内存使用超过95%
            warnings.append(f"系统内存使用率高: {mem.percent}%")
        
        # 检查CPU使用情况
        cpu_usage = psutil.cpu_percent(interval=0.5)
        if cpu_usage > 95:  # 如果CPU使用超过95%
            warnings.append(f"CPU使用率高: {cpu_usage}%")
        
        # 检查磁盘空间
        disk = psutil.disk_usage('/')
        if disk.percent > 95:  # 如果磁盘使用超过95%
            warnings.append(f"磁盘空间不足: 已使用 {disk.percent}%")
        
        # 输出警告信息
        if warnings:
            print("系统资源警告:")
            for warning in warnings:
                print(f" - {warning}")
            print("以上情况可能会影响应用性能，建议关闭其他占用资源的应用。")
            return False
        else:
            print("系统资源检查通过。")
            return True
    except Exception as e:
        print(f"检查系统资源时出错: {str(e)}")
        return True  # 出错时默认继续运行

def ensure_final_cleanup():
    """确保应用退出时进行最终清理"""
    print("正在执行最终资源清理...")
    try:
        # 1. 先清理任务管理器中的资源
        task_manager.cleanup_all_resources()
        
        # 2. 终止所有可能的残留驱动进程
        #    注意：不再终止 msedge.exe
        if sys.platform == 'win32':
            for process_name in ['msedgedriver.exe', 'EdgeWebDriver.exe']: # 移除了 'msedge.exe'
                try:
                    subprocess.run(
                        f'taskkill /f /im {process_name}', 
                        shell=True, 
                        stdout=subprocess.DEVNULL, 
                        stderr=subprocess.DEVNULL,
                        timeout=2
                    )
                except:
                    pass
        
        # 3. 使用psutil确保干净退出
        #    注意：不再查找 msedge
        process_keywords = ['msedgedriver', 'edgewebdriver'] # 移除了 'msedge'
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if (proc.info.get('name') and 
                    any(keyword in proc.info.get('name').lower() for keyword in process_keywords) and
                    proc.pid != os.getpid()):  # 确保不终止自己
                    try:
                        proc.terminate()
                        gone, alive = psutil.wait_procs([proc], timeout=1)
                        if proc in alive:
                            proc.kill()
                    except:
                        pass
            except:
                pass
        
        print("最终清理完成。")
    except Exception as e:
        print(f"最终清理时出错: {str(e)}")

def main():
    # 应用启动时先执行一次全局清理
    cleanup_webdriver_temp_files()
    
    # 检查系统资源状态
    resource_check_ok = check_system_resources()
    
    # 稍等一段时间以确保资源完全释放
    time.sleep(1)
    
    app = QApplication(sys.argv)
    
    # 连接 aboutToQuit 信号以在应用退出时进行全面清理
    app.aboutToQuit.connect(ensure_final_cleanup)
    
    # 设置内置主题样式
    app.setStyle("Fusion")

    window = MainWindow()
    window.setWindowIcon(QIcon(resource_path('resources/logo.ico')))  # 使用打包后资源路径
    
    # 如果资源检查有警告，可以在界面上显示提示
    if not resource_check_ok:
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.warning(
            window, 
            "系统资源警告",
            "检测到系统资源不足，可能会影响应用性能。\n建议关闭其他占用资源的应用后再使用。",
            QMessageBox.Ok
        )
    
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
