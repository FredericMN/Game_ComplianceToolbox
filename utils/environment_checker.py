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
import uuid    # 导入uuid用于生成唯一ID
import tempfile
import glob
import random
import datetime


class EnvironmentChecker(QObject):
    """环境检查类：并行执行若干检测项，并输出结果。"""
    # 输出文本消息
    output_signal = Signal(str)
    # 结构化结果信号，传递一个列表[(item_name, status_bool, detail), ...]
    # 其中 item_name 为检测项名称, status_bool 为是否通过, detail 为备注
    structured_result_signal = Signal(list)
    # 检测完成
    finished = Signal(bool)  # 参数表示是否有错误
    
    # 设置最大保留的临时目录数量，防止过度膨胀
    MAX_TEMP_DIRS = 20

    def __init__(self):
        super().__init__()
        self.edge_version = None
        self.has_errors = False
        # 确保初始化structured_results为空列表
        self.structured_results = []
        self.driver = None  # 跟踪WebDriver实例
        self.driver_future = None  # 跟踪异步任务
        
        # 移除在构造函数中的预清理，避免阻塞UI
        # self.pre_cleanup()

        # 自定义的环境检测项目列表，可按需扩展
        self.check_items = [
            ("网络连接检测", self.check_network),
            ("Edge浏览器检测", self.check_edge_browser),
            ("Edge WebDriver检测", self.check_edge_driver)
        ]
        
        # 当前运行时创建的临时目录列表，用于在结束时清理
        self.current_temp_dirs = []
    
    def pre_cleanup(self):
        """应用启动时的预清理，确保没有残留的驱动进程和临时目录"""
        try:
            # 先清理所有msedgedriver进程
            self.kill_msedgedriver()
            
            # 清理所有可能存在的临时目录
            self.cleanup_temp_directories(enforce_limit=True)
            
        except Exception:
            pass

    def run(self):
        """依次执行所有检测项，可考虑并发，但量少时顺序即可。"""
        # 确保每次运行前清空结果列表
        self.structured_results = []
        self.current_temp_dirs = []  # 重置当前运行时创建的临时目录列表
        
        try:
            # 在后台线程中执行预清理，不阻塞UI
            self.pre_cleanup()
            
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

    def register_temp_dir(self, dir_path):
        """注册临时目录，以便在结束时清理"""
        if dir_path and os.path.isdir(dir_path):
            self.current_temp_dirs.append(dir_path)
            
    def cleanup_resources(self):
        """清理所有资源，包括WebDriver和msedgedriver进程"""
        # 关闭WebDriver
        self.quit_driver()
        
        # 清理msedgedriver进程
        self.kill_msedgedriver()
        
        # 清理临时用户数据目录
        cleaned_dirs = 0
        failed_dirs = 0
        
        # 先清理本次运行创建的目录
        for dir_path in self.current_temp_dirs:
            if dir_path and os.path.isdir(dir_path):
                try:
                    shutil.rmtree(dir_path, ignore_errors=True)
                    cleaned_dirs += 1
                except Exception as e:
                    failed_dirs += 1
                    self.output_signal.emit(f"清理临时目录失败: {dir_path}, 错误: {str(e)}")
        
        # 然后清理其他可能存在的临时目录，并限制总数
        try:
            self.cleanup_temp_directories(enforce_limit=True)
        except Exception as e:
            self.output_signal.emit(f"清理临时目录时出错: {str(e)}")
            
        if cleaned_dirs > 0 or failed_dirs > 0:
            self.output_signal.emit(f"环境清理：已清理 {cleaned_dirs} 个临时目录，失败 {failed_dirs} 个。")

    def cleanup_temp_directories(self, enforce_limit=False):
        """清理所有可能的临时目录，可选择限制总数"""
        try:
            temp_dir = tempfile.gettempdir()
            patterns = ["edge_driver_*", "edge_temp_*", "edge_port_*", "edge_alt_*", "scoped_dir*"]
            
            if enforce_limit:
                # 如果需要强制限制目录总数，收集所有匹配的目录及其修改时间
                all_dirs = []
                for pattern in patterns:
                    for path in glob.glob(os.path.join(temp_dir, pattern)):
                        try:
                            if os.path.isdir(path):
                                # 获取目录修改时间
                                mtime = os.path.getmtime(path)
                                all_dirs.append((path, mtime))
                        except Exception:
                            pass
                
                # 如果目录数超过限制，按修改时间排序，保留最新的
                if len(all_dirs) > self.MAX_TEMP_DIRS:
                    # 按修改时间从新到旧排序
                    all_dirs.sort(key=lambda x: x[1], reverse=True)
                    # 只保留最新的MAX_TEMP_DIRS个，删除其余的
                    for path, _ in all_dirs[self.MAX_TEMP_DIRS:]:
                        try:
                            shutil.rmtree(path, ignore_errors=True)
                        except Exception:
                            pass
            else:
                # 如果不需要限制总数，直接清理本次检测前存在的目录
                for pattern in patterns:
                    for path in glob.glob(os.path.join(temp_dir, pattern)):
                        try:
                            if os.path.isdir(path) and path not in self.current_temp_dirs:
                                shutil.rmtree(path, ignore_errors=True)
                        except Exception:
                            pass
        except Exception:
            pass

    def quit_driver(self):
        """安全地关闭WebDriver"""
        if self.driver:
            try:
                # 首先检查driver是否有关联的进程需要清理
                if hasattr(self.driver, '_msedgedriver_process'):
                    try:
                        self.driver._msedgedriver_process.kill()
                    except Exception:
                        pass
                
                # 尝试关闭浏览器窗口
                try:
                    self.driver.close()
                except Exception:
                    pass
                
                # 然后退出驱动
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
        
        # 额外清理 - 确保所有相关进程都被终止
        if sys.platform == 'win32':
            try:
                subprocess.call('taskkill /f /im msedgedriver.exe', shell=True, 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass

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
                    # self.output_signal.emit("已使用系统命令尝试终止所有msedgedriver进程") # 减少冗余输出
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
                            # self.output_signal.emit(f"已终止 msedgedriver.exe (PID: {proc.info['pid']})") # 减少冗余输出
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
        
        # 清理所有临时目录，避免目录锁定问题
        self.cleanup_temp_directories()
        
        # 添加更长的等待时间，确保之前的进程被清理
        time.sleep(2)  
        
        # 2. 尝试不同的启动策略
        strategies = [
            self._try_edge_headless_incognito,      # 使用无头浏览器和隐私模式
            self._try_edge_with_unique_profile,     # 使用唯一临时配置文件
            self._try_edge_with_random_port,        # 使用随机端口
            self._try_edge_with_alternative_service  # 使用替代服务启动方式
        ]
        
        # 3. 依次尝试不同策略
        for i, strategy in enumerate(strategies):
            try:
                # 每次尝试前都清理环境
                self.kill_all_browser_processes()
                time.sleep(1.5)  # 增加等待时间，确保进程真正被清理
                
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
                # 增加等待时间
                time.sleep(1.5)
        
        # 清理环境再尝试下载新的WebDriver
        self.kill_all_browser_processes()
        time.sleep(2)
        
        # 如果所有策略都失败，尝试下载新的WebDriver
        try:
            self.output_signal.emit("当前Edge WebDriver无法启动，尝试自动下载...")
            driver_path = EdgeChromiumDriverManager().install()
            self.output_signal.emit(f"新的Edge WebDriver已下载: {driver_path}")
            
            # 确保驱动程序可执行
            if sys.platform == 'win32':
                try:
                    os.chmod(driver_path, 0o755)
                except:
                    pass
            
            # 下载后再次彻底清理环境
            self.kill_all_browser_processes()
            time.sleep(2)
            
            # 再次尝试使用新下载的驱动和不同策略
            for i, strategy in enumerate(strategies):
                try:
                    # 每次尝试前都清理环境
                    self.kill_all_browser_processes()
                    time.sleep(1.5)
                    
                    result, driver = strategy(driver_path=driver_path)
                    
                    if result:
                        self.driver = driver
                        self.output_signal.emit("Edge WebDriver已自动下载并配置成功。")
                        return True, "Edge WebDriver下载并配置成功。"
                        
                except Exception as e:
                    self.output_signal.emit(f"使用新驱动启动Edge WebDriver失败: {str(e)}")
                    self.kill_all_browser_processes()
                    time.sleep(1.5)
                
        except Exception as e:
            self.output_signal.emit(f"下载Edge WebDriver失败: {str(e)}")
            
        # 所有方法都失败
        self.output_signal.emit("无法配置Edge WebDriver，请检查Edge浏览器安装或网络连接。")
        return False, "所有WebDriver启动策略均失败。"

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
            
            # 更彻底的进程清理
            for process_name in ['msedge.exe', 'msedgedriver.exe', 'EdgeWebDriver.exe']:
                try:
                    subprocess.call(f'taskkill /f /im {process_name}', shell=True, 
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except:
                    pass
        
        # 确保任何残留的WebDriver连接被关闭
        self.quit_driver()
        
        # 手动查找和终止所有Edge和EdgeDriver相关进程
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and ('msedge' in proc.info['name'].lower() or 'edge' in proc.info['name'].lower()):
                    try:
                        proc.kill()
                    except:
                        pass
            except:
                pass

    def create_edge_driver_with_timeout(self, options, service=None, timeout=10):
        """并发启动Edge WebDriver，如超时则返回None"""
        def create_driver():
            try:
                if service:
                    return webdriver.Edge(service=service, options=options)
                else:
                    return webdriver.Edge(options=options)
            except Exception as e:
                # self.output_signal.emit(f"创建Edge WebDriver时异常: {e}") # 改为在调用处处理
                raise e # 将异常抛出，由调用者处理

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
                self.output_signal.emit(f"启动Edge WebDriver时发生错误: {e}") # 统一在这里输出错误
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
            # self.output_signal.emit(f"从注册表获取Edge版本失败: {str(e)}") # 改为在调用处处理
            raise Exception(f"从注册表获取Edge版本失败: {str(e)}")

    def _try_edge_headless_incognito(self, driver_path=None):
        """策略1: 使用无头浏览器和隐私模式，不使用用户数据目录"""
        unique_id = uuid.uuid4().hex  # 生成唯一ID
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")  # 更精确的时间戳
        
        options = Options()
        options.use_chromium = True
        options.add_argument("--headless")
        options.add_argument("--incognito")  # 隐私模式
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-first-run")
        options.add_argument("--disable-application-cache")
        options.add_argument("--disable-web-security")
        # 第一个策略尝试不使用用户数据目录
        # options.add_argument(f"--user-agent=Mozilla/5.0 AppleWebKit/537.36 Chrome/{self.edge_version}")
        
        try:
            if driver_path:
                service = EdgeService(executable_path=driver_path)
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            else:
                service = EdgeService()
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            
            return driver is not None, driver
        except Exception as e:
            self.output_signal.emit(f"策略1启动失败: {str(e)}")
            return False, None

    def _try_edge_with_unique_profile(self, driver_path=None):
        """策略2: 使用唯一临时配置文件"""
        unique_id = uuid.uuid4().hex
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")  # 更精确的时间戳
        # 使用时间戳和UUID确保唯一性，并在系统临时目录内创建一个完全随机名称的子目录
        temp_base = os.path.join(tempfile.gettempdir(), f"edge_temp_{timestamp}_{unique_id}")
        
        # 确保父目录存在
        try:
            os.makedirs(os.path.dirname(temp_base), exist_ok=True)
        except Exception:
            pass
            
        # 使用绝对路径并确保路径不包含特殊字符
        temp_dir = os.path.abspath(temp_base)
        
        try:
            # 首先验证目录是否真的可用
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir, exist_ok=True)
            elif not os.access(temp_dir, os.W_OK):
                # 如果目录已存在但无法写入，生成新目录
                temp_dir = os.path.join(tempfile.gettempdir(), f"edge_temp_{timestamp}_{uuid.uuid4().hex}")
                os.makedirs(temp_dir, exist_ok=True)
                
            # 注册临时目录，以便在结束时清理
            self.register_temp_dir(temp_dir)
            
            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            # 明确指定用户数据目录为绝对路径
            options.add_argument(f"--user-data-dir={temp_dir}")
            options.add_argument("--profile-directory=TestProfile")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-application-cache")
            
            if driver_path:
                service = EdgeService(executable_path=driver_path)
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            else:
                service = EdgeService()
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            
            return driver is not None, driver
        except Exception as e:
            self.output_signal.emit(f"策略2启动失败: {str(e)}")
            # 确保在失败时清理
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass
            return False, None

    def _try_edge_with_random_port(self, driver_path=None):
        """策略3: 使用随机端口和唯一目录"""
        # 使用更大范围的随机端口，降低冲突可能性
        port = random.randint(10000, 32000)
        unique_id = uuid.uuid4().hex
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")  # 更精确的时间戳
        temp_dir = os.path.join(tempfile.gettempdir(), f"edge_port_{timestamp}_{unique_id}")
        
        try:
            # 验证目录是否真的可用
            if os.path.exists(temp_dir):
                # 如果目录已存在，生成新目录
                temp_dir = os.path.join(tempfile.gettempdir(), f"edge_port_{timestamp}_{uuid.uuid4().hex}")
                
            os.makedirs(temp_dir, exist_ok=True)
            # 注册临时目录，以便在结束时清理
            self.register_temp_dir(temp_dir)
            
            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            options.add_argument(f"--user-data-dir={temp_dir}")
            options.add_argument(f"--remote-debugging-port={port}")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            
            # 确保 EdgeService 已导入
            from selenium.webdriver.edge.service import Service as EdgeService
            
            if driver_path:
                service = EdgeService(executable_path=driver_path)
                service.port = port  # 设置服务端口
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            else:
                service = EdgeService(port=port)
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
            
            return driver is not None, driver
        except Exception as e:
            self.output_signal.emit(f"策略3启动失败: {str(e)}")
            # 确保在失败时清理
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass
            return False, None

    def _try_edge_with_alternative_service(self, driver_path=None):
        """策略4: 使用替代服务启动方式"""
        try:
            # 生成随机端口和唯一临时目录
            port = random.randint(10000, 32000)  # 使用更大范围端口
            unique_id = uuid.uuid4().hex
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")  # 更精确的时间戳
            temp_dir = os.path.join(tempfile.gettempdir(), f"edge_alt_{timestamp}_{unique_id}")
            
            # 检查目录是否可用
            if os.path.exists(temp_dir):
                # 如果目录已存在，生成新目录
                temp_dir = os.path.join(tempfile.gettempdir(), f"edge_alt_{timestamp}_{uuid.uuid4().hex}")
                
            os.makedirs(temp_dir, exist_ok=True)
            # 注册临时目录，以便在结束时清理
            self.register_temp_dir(temp_dir)
            
            # 手动启动msedgedriver进程
            if not driver_path:
                try:
                    from webdriver_manager.microsoft import EdgeChromiumDriverManager
                    driver_path = EdgeChromiumDriverManager().install()
                    
                    # 确保驱动程序可执行
                    if sys.platform == 'win32':
                        try:
                            os.chmod(driver_path, 0o755)
                        except:
                            pass
                except Exception as e:
                    self.output_signal.emit(f"获取驱动路径失败: {str(e)}")
                    return False, None
            
            # 启动后台进程，添加参数使其避免使用已存在的用户数据目录
            try:
                # 在命令行中添加禁用用户数据目录相关参数
                process = subprocess.Popen(
                    [driver_path, f"--port={port}", "--disable-dev-shm-usage", "--no-sandbox", 
                    "--disable-extensions", "--disable-application-cache"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE
                )
                
                # 给服务一些启动时间
                time.sleep(1.0)  # 增加等待时间
            except Exception as e:
                self.output_signal.emit(f"启动驱动服务失败: {str(e)}")
                return False, None
            
            # 配置WebDriver连接到已运行的服务
            try:
                options = Options()
                options.use_chromium = True
                options.add_argument("--headless")
                # 使用完全唯一的用户数据目录
                options.add_argument(f"--user-data-dir={temp_dir}")
                options.add_argument("--disable-extensions")
                options.add_argument("--disable-gpu")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                
                # 直接连接到已运行的msedgedriver
                driver = webdriver.Remote(
                    command_executor=f'http://localhost:{port}',
                    options=options
                )
                
                # 保存进程引用以便稍后清理
                driver._msedgedriver_process = process
                
                return True, driver
            except Exception as e:
                self.output_signal.emit(f"连接到驱动服务失败: {str(e)}")
                # 确保进程被终止
                if 'process' in locals():
                    try:
                        process.kill()
                    except Exception:
                        pass
                # 确保在失败时清理
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception:
                    pass
                return False, None
        except Exception as e:
            self.output_signal.emit(f"策略4整体失败: {str(e)}")
            return False, None
