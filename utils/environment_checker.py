# utils/environment_checker.py

import shutil
import subprocess
import sys
import os
from PySide6.QtCore import QObject, Signal
import requests
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
import winreg
import concurrent.futures
import time
import psutil  # 添加psutil导入


class EnvironmentChecker(QObject):
    """环境检查类：并行执行若干检测项，并输出结果。"""
    # 输出文本消息
    output_signal = Signal(str)
    # 结构化结果信号，传递一个列表[(item_name, status_bool, detail), ...]
    # 其中 item_name 为检测项名称, status_bool 为是否通过, detail 为备注
    structured_result_signal = Signal(list)
    # 检测完成
    finished = Signal(bool)  # 参数表示是否有错误

    def __init__(self):
        super().__init__()
        self.edge_version = None
        self.has_errors = False
        # 确保初始化structured_results为空列表
        self.structured_results = []
        self.driver = None  # 跟踪WebDriver实例
        self.driver_future = None  # 跟踪异步任务

        # 自定义的环境检测项目列表，可按需扩展
        self.check_items = [
            ("网络连接检测", self.check_network),
            ("Edge浏览器检测", self.check_edge_browser),
            ("Edge WebDriver检测", self.check_edge_driver)
        ]

    def run(self):
        """依次执行所有检测项，可考虑并发，但量少时顺序即可。"""
        # 确保每次运行前清空结果列表
        self.structured_results = []
        
        try:
            for item_name, func in self.check_items:
                try:
                    self.output_signal.emit(f"开始检测：{item_name} ...")
                    ok, detail = func()  # 执行检测项，返回(bool, str)
                    if not ok:
                        self.has_errors = True
                    self.structured_results.append((item_name, ok, detail))
                except Exception as e:
                    self.has_errors = True
                    detail_msg = f"检测过程中出现异常: {str(e)}"
                    self.output_signal.emit(detail_msg)
                    self.structured_results.append((item_name, False, detail_msg))
        finally:
            # 确保在检测结束时清理所有资源
            self.cleanup_resources()
            
            # 发射结构化结果
            self.structured_result_signal.emit(self.structured_results)
            self.finished.emit(self.has_errors)

    def cleanup_resources(self):
        """清理所有资源，包括WebDriver和msedgedriver进程"""
        # 关闭WebDriver
        self.quit_driver()
        
        # 清理msedgedriver进程
        self.kill_msedgedriver()
        
        # 清理临时用户数据目录
        cleaned_dirs = 0
        failed_dirs = 0
        try:
            import tempfile
            import glob
            import shutil
            
            # 查找并删除所有edge_driver_*临时目录
            pattern = os.path.join(tempfile.gettempdir(), "edge_driver_*")
            for dir_path in glob.glob(pattern):
                if os.path.isdir(dir_path):
                    try:
                        shutil.rmtree(dir_path, ignore_errors=True)
                        cleaned_dirs += 1
                    except Exception as e:
                        failed_dirs += 1
                        self.output_signal.emit(f"清理临时目录失败: {dir_path}, 错误: {str(e)}")
            if cleaned_dirs > 0 or failed_dirs > 0:
                self.output_signal.emit(f"环境清理：已清理 {cleaned_dirs} 个临时目录，失败 {failed_dirs} 个。")
        except Exception as e:
            self.output_signal.emit(f"清理临时目录时出错: {str(e)}")

    def quit_driver(self):
        """安全地关闭WebDriver"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception as e:
                self.output_signal.emit(f"关闭WebDriver时出错: {str(e)}")
            finally:
                self.driver = None
        
        # 取消未完成的异步任务
        if self.driver_future and not self.driver_future.done():
            try:
                self.driver_future.cancel()
            except Exception:
                pass
            self.driver_future = None

    def kill_msedgedriver(self):
        """终止所有 msedgedriver.exe 进程"""
        terminated_count = 0
        failed_count = 0
        
        try:
            # 首先尝试使用系统命令强制终止所有msedgedriver进程
            if sys.platform == 'win32':
                try:
                    # 使用taskkill命令强制终止所有msedgedriver进程
                    subprocess.call('taskkill /f /im msedgedriver.exe', shell=True, 
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass
            
            # 然后使用psutil逐个终止
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
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                            failed_count += 1
                            self.output_signal.emit(f"终止 msedgedriver.exe (PID: {proc.info['pid']}) 时出错: {e}")
                except Exception as e:
                    pass
            
            if terminated_count > 0 or failed_count > 0:
                self.output_signal.emit(f"环境清理：已终止 {terminated_count} 个残留驱动进程，失败 {failed_count} 个。")
        except Exception as e:
            self.output_signal.emit(f"清理驱动进程时出现异常: {e}")

    def check_network(self):
        """检测网络连接"""
        try:
            response = requests.get("https://cn.bing.com/", timeout=3)
            if response.status_code == 200:
                msg = "网络连接正常。"
                self.output_signal.emit(msg)
                return True, msg
            else:
                msg = "网络异常，请检查网络连接。"
                self.output_signal.emit(msg)
                return False, msg
        except requests.RequestException:
            msg = "网络异常，请检查网络连接。"
            self.output_signal.emit(msg)
            return False, msg

    def check_edge_browser(self):
        """检测Edge浏览器是否安装，以及版本信息"""
        edge_path = self.find_edge_executable()
        if edge_path and os.path.exists(edge_path):
            try:
                version = self.get_edge_version_from_registry()
                if version:
                    self.output_signal.emit(f"检测到Edge浏览器，版本：{version}")
                    self.edge_version = version
                    return True, f"Edge浏览器已安装，版本：{version}"
                else:
                    msg = "无法获取Edge浏览器版本。"
                    self.output_signal.emit(msg)
                    return False, msg
            except Exception as e:
                msg = f"无法获取Edge浏览器版本: {str(e)}"
                self.output_signal.emit(msg)
                return False, msg
        else:
            msg = ("未检测到Edge浏览器，请先安装Edge浏览器。"
                   "下载链接: https://www.microsoft.com/edge")
            self.output_signal.emit(msg)
            return False, msg

    def check_edge_driver(self):
        """检测Edge WebDriver是否可用，否则尝试下载"""
        if not self.edge_version:
            msg = "Edge版本未知，需先安装或检测Edge浏览器后再检测WebDriver。"
            self.output_signal.emit(msg)
            return False, msg

        self.output_signal.emit("检测Edge WebDriver...")
        
        # 1. 在检测前彻底清理环境
        self.kill_all_browser_processes()
        
        import tempfile
        import uuid
        import time
        import random
        
        # 2. 尝试不同的启动策略
        strategies = [
            self._try_edge_without_user_data_dir,     # 不使用用户数据目录
            self._try_edge_with_temp_profile,         # 使用临时配置文件
            self._try_edge_with_random_port,          # 使用随机端口
            self._try_edge_with_alternative_service    # 使用替代服务启动方式
        ]
        
        # 3. 依次尝试不同策略
        for i, strategy in enumerate(strategies):
            try:
                self.output_signal.emit(f"尝试WebDriver启动策略 {i+1}/{len(strategies)}...")
                # 每次尝试前先睡眠一小段随机时间，避免冲突
                time.sleep(random.uniform(0.5, 1.5))
                
                # 执行策略
                result, driver = strategy()
                
                if result:
                    self.driver = driver
                    self.output_signal.emit("Edge WebDriver已正确配置。")
                    return True, "Edge WebDriver已正确配置。"
                    
            except Exception as e:
                self.output_signal.emit(f"启动Edge WebDriver遇到问题: {str(e)}")
                # 策略失败后清理环境
                self.kill_all_browser_processes()
                # 短暂延迟
                time.sleep(random.uniform(1.0, 2.0))
        
        # 如果所有策略都失败，尝试下载新的WebDriver
        try:
            self.output_signal.emit("当前Edge WebDriver无法启动，尝试自动下载...")
            driver_path = EdgeChromiumDriverManager().install()
            self.output_signal.emit(f"新的Edge WebDriver已下载: {driver_path}")
            
            # 再次尝试使用新下载的驱动和不同策略
            for i, strategy in enumerate(strategies):
                try:
                    time.sleep(random.uniform(0.5, 1.5))
                    result, driver = strategy(driver_path=driver_path)
                    
                    if result:
                        self.driver = driver
                        self.output_signal.emit("Edge WebDriver已自动下载并配置成功。")
                        return True, "Edge WebDriver下载并配置成功。"
                        
                except Exception as e:
                    self.output_signal.emit(f"使用新驱动启动Edge WebDriver失败: {str(e)}")
                    self.kill_all_browser_processes()
                    time.sleep(random.uniform(1.0, 2.0))
                
        except Exception as e:
            self.output_signal.emit(f"下载Edge WebDriver失败: {str(e)}")
            
        # 所有方法都失败
        self.output_signal.emit("无法配置Edge WebDriver，请检查Edge浏览器安装或网络连接。")
        return False, "所有WebDriver启动策略均失败。"

    def _try_edge_without_user_data_dir(self, driver_path=None):
        """策略1: 完全不使用用户数据目录"""
        options = Options()
        options.use_chromium = True
        options.add_argument("--headless")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-first-run")
        # 关键：明确禁用用户数据目录
        options.add_argument("--incognito")
        
        if driver_path:
            service = EdgeService(driver_path)
            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
        else:
            driver = self.create_edge_driver_with_timeout(options, service=None, timeout=10)
        
        return driver is not None, driver

    def _try_edge_with_temp_profile(self, driver_path=None):
        """策略2: 使用临时配置文件"""
        import tempfile
        import uuid
        import shutil
        import os
        
        temp_dir = os.path.join(tempfile.gettempdir(), f"edge_temp_{uuid.uuid4().hex}")
        os.makedirs(temp_dir, exist_ok=True)
        
        options = Options()
        options.use_chromium = True
        options.add_argument("--headless")
        options.add_argument(f"--user-data-dir={temp_dir}")
        options.add_argument("--profile-directory=TestProfile")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        
        if driver_path:
            service = EdgeService(driver_path)
            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
        else:
            driver = self.create_edge_driver_with_timeout(options, service=None, timeout=10)
        
        return driver is not None, driver

    def _try_edge_with_random_port(self, driver_path=None):
        """策略3: 使用随机端口"""
        import random
        import tempfile
        import uuid
        import os
        
        port = random.randint(9000, 9999)
        temp_dir = os.path.join(tempfile.gettempdir(), f"edge_port_{uuid.uuid4().hex}")
        os.makedirs(temp_dir, exist_ok=True)
        
        options = Options()
        options.use_chromium = True
        options.add_argument("--headless")
        options.add_argument(f"--user-data-dir={temp_dir}")
        options.add_argument(f"--remote-debugging-port={port}")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        
        if driver_path:
            service = EdgeService(driver_path)
            service.port = port  # 设置服务端口
            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
        else:
            from selenium.webdriver.edge.service import Service as EdgeService
            service = EdgeService(port=port)
            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
        
        return driver is not None, driver

    def _try_edge_with_alternative_service(self, driver_path=None):
        """策略4: 使用替代服务启动方式"""
        import subprocess
        import time
        import tempfile
        import uuid
        import os
        import random
        
        # 生成随机端口和唯一临时目录
        port = random.randint(9000, 9999)
        temp_dir = os.path.join(tempfile.gettempdir(), f"edge_alt_{uuid.uuid4().hex}")
        os.makedirs(temp_dir, exist_ok=True)
        
        try:
            # 手动启动msedgedriver进程
            if not driver_path:
                from webdriver_manager.microsoft import EdgeChromiumDriverManager
                driver_path = EdgeChromiumDriverManager().install()
            
            # 启动后台进程
            process = subprocess.Popen(
                [driver_path, f"--port={port}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            
            # 给服务一些启动时间
            time.sleep(2)
            
            # 配置WebDriver连接到已运行的服务
            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            options.add_argument(f"--user-data-dir={temp_dir}")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            
            # 直接连接到已运行的msedgedriver
            from selenium import webdriver
            driver = webdriver.Remote(
                command_executor=f'http://localhost:{port}',
                options=options
            )
            
            # 保存进程引用以便稍后清理
            driver._msedgedriver_process = process
            
            return True, driver
        except Exception as e:
            # 确保进程被终止
            if 'process' in locals():
                try:
                    process.kill()
                except:
                    pass
            raise e

    def kill_all_browser_processes(self):
        """彻底清理所有浏览器和WebDriver相关进程"""
        # 先清理msedgedriver进程
        self.kill_msedgedriver()
        
        # 也尝试清理Edge浏览器进程
        if sys.platform == 'win32':
            try:
                # 使用taskkill命令清理Edge进程
                subprocess.call('taskkill /f /im msedge.exe', shell=True, 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except:
                pass
        
        # 确保任何残留的WebDriver连接被关闭
        self.quit_driver()
        
        # 短暂暂停，确保进程完全终止
        import time
        time.sleep(1)

    def create_edge_driver_with_timeout(self, options, service=None, timeout=10):
        """并发启动Edge WebDriver，如超时则返回None"""
        def create_driver():
            try:
                if service:
                    return webdriver.Edge(service=service, options=options)
                else:
                    return webdriver.Edge(options=options)
            except Exception as e:
                self.output_signal.emit(f"创建Edge WebDriver时异常: {e}")
                return None

        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            try:
                self.driver_future = executor.submit(create_driver)
                driver = self.driver_future.result(timeout=timeout)
                self.driver_future = None
                return driver
            except concurrent.futures.TimeoutError:
                self.output_signal.emit(f"启动Edge WebDriver超时 ({timeout}秒)")
                return None
            except Exception as e:
                self.output_signal.emit(f"启动Edge WebDriver时发生错误: {e}")
                return None

    def find_edge_executable(self):
        """尝试找到Edge浏览器可执行文件"""
        possible_paths = [
            shutil.which('msedge'),
            shutil.which('edge'),
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
        for path in possible_paths:
            if path and os.path.exists(path):
                return path
        return None

    def get_edge_version_from_registry(self):
        """从Windows注册表获取Edge版本"""
        try:
            registry_paths = [
                r"SOFTWARE\Microsoft\Edge\BLBeacon",
                r"SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon"
            ]
            # 先从 HKEY_CURRENT_USER
            for reg_path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ) as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                except FileNotFoundError:
                    continue
            # 再从 HKEY_LOCAL_MACHINE
            for reg_path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ) as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                except FileNotFoundError:
                    continue
            return None
        except Exception as e:
            raise Exception(f"从注册表获取Edge版本失败: {str(e)}")
