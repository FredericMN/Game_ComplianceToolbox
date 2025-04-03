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

# 导入驱动管理器
try:
    from utils.driver_manager import driver_manager
except ImportError:
    # 如果导入失败，创建一个简易的管理器
    class SimpleDriverManager:
        def set_driver_path(self, driver_path, driver_version=None):
            return True
        def get_driver_path(self):
            return None
        def is_initialized(self):
            return False
    driver_manager = SimpleDriverManager()


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
        
        # 进度回调函数
        self.progress_callback = None

        # 自定义的环境检测项目列表，可按需扩展
        self.check_items = [
            ("网络连接检测", self.check_network),
            ("Edge浏览器检测", self.check_edge_browser),
            ("Edge WebDriver检测", self.check_edge_driver)
        ]
        
        # 当前运行时创建的临时目录列表，用于在结束时清理
        self.current_temp_dirs = []
    
    def set_progress_callback(self, callback):
        """设置进度回调函数"""
        self.progress_callback = callback
    
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
        """依次执行所有检测项，区分关键错误和警告。"""
        # 确保每次运行前清空结果列表
        self.structured_results = []
        self.current_temp_dirs = []  # 重置当前运行时创建的临时目录列表
        self.has_errors = False  # 重置错误标志
        network_ok = True  # 单独跟踪网络状态
        
        try:
            # 在后台线程中执行预清理，不阻塞UI
            self.output_signal.emit("正在清理环境...")
            self.pre_cleanup()
            
            # 如果驱动管理器已初始化，显示信息
            if driver_manager.is_initialized():
                driver_path = driver_manager.get_driver_path()
                driver_version = driver_manager.get_driver_version() or "未知"
                self.output_signal.emit(f"发现缓存的WebDriver: 路径={driver_path}, 版本={driver_version}")
            
            total_checks = len(self.check_items)
            for idx, (item_name, func) in enumerate(self.check_items):
                try:
                    progress_percent = int(25 + idx * 75 / total_checks)  # 进度从25%到100%
                    self.output_signal.emit(f"开始检测：{item_name} ... ({idx+1}/{total_checks})")
                    
                    # 发送当前检测项的进度
                    if self.progress_callback:
                        self.progress_callback(f"正在检测: {item_name}", progress_percent)
                    
                    ok, detail = func()  # 执行检测项，返回(bool, str)
                    
                    # 区分错误类型
                    if item_name == "网络连接检测":
                        network_ok = ok # 记录网络状态
                        # 网络问题不直接标记为 has_errors
                        self.structured_results.append((item_name, ok, detail))
                        if not ok:
                            self.output_signal.emit(f"警告：{detail}") # 输出警告
                    elif not ok:
                        # 其他关键检测项失败，标记错误
                        self.has_errors = True
                        self.structured_results.append((item_name, ok, detail))
                    else:
                        # 检测项成功
                        self.structured_results.append((item_name, ok, detail))
                    
                except Exception as e:
                    # 任何检测项出现异常都视为错误
                    self.has_errors = True
                    detail_msg = f"{item_name} 检测过程中出现异常: {str(e)}"
                    self.output_signal.emit(detail_msg)
                    self.structured_results.append((item_name, False, detail_msg))
        finally:
            # 确保在检测结束时清理所有资源
            self.cleanup_resources()
            
            # 发射结构化结果
            self.structured_result_signal.emit(self.structured_results)
            # finished 信号现在只反映关键错误
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
            # 获取缓存的WebDriver路径
            cached_driver_path = driver_manager.get_driver_path()
            
            temp_dir = tempfile.gettempdir()
            patterns = ["edge_driver_*", "edge_temp_*", "edge_port_*", "edge_alt_*", "scoped_dir*"]
            
            all_dirs = []
            for pattern in patterns:
                for path in glob.glob(os.path.join(temp_dir, pattern)):
                    try:
                        # 跳过当前运行时创建的临时目录
                        if path in self.current_temp_dirs:
                            continue
                            
                        # 跳过包含缓存WebDriver的目录
                        if cached_driver_path and os.path.exists(cached_driver_path):
                            driver_dir = os.path.dirname(cached_driver_path)
                            if path == driver_dir or driver_dir.startswith(path) or path.startswith(driver_dir):
                                self.output_signal.emit(f"跳过缓存WebDriver目录: {path}")
                                continue
                        
                        # 其余目录可清理
                        if os.path.isdir(path):
                            if enforce_limit:
                                # 收集目录及修改时间，稍后处理
                                all_dirs.append((path, os.path.getmtime(path)))
                            else:
                                # 直接清理
                                shutil.rmtree(path, ignore_errors=True)
                    except Exception:
                        pass
                
            # 如果启用限制，则保留最新的MAX_TEMP_DIRS个目录
            if enforce_limit and all_dirs:
                # 按修改时间从新到旧排序
                all_dirs.sort(key=lambda x: x[1], reverse=True)
                # 删除超出限制的旧目录
                for path, _ in all_dirs[self.MAX_TEMP_DIRS:]:
                    try:
                        shutil.rmtree(path, ignore_errors=True)
                    except Exception:
                        pass
        except Exception as e:
            self.output_signal.emit(f"清理临时目录时发生异常: {str(e)}")

    def quit_driver(self):
        """安全地关闭WebDriver"""
        if self.driver:
            try:
                # 首先检查driver是否有关联的进程需要清理
                if hasattr(self.driver, '_msedgedriver_process'):
                    try:
                        # 尝试获取进程并终止
                        proc = psutil.Process(self.driver.service.process.pid)
                        if proc.is_running() and proc.name().lower() == 'msedgedriver.exe':
                            proc.kill()
                    except (psutil.NoSuchProcess, AttributeError, Exception):
                        # 忽略进程不存在或属性错误等
                        pass
                
                # 尝试关闭浏览器窗口
                try:
                    self.driver.close() # 通常会关闭由driver启动的浏览器窗口
                except WebDriverException:
                    # 忽略关闭窗口时可能出现的错误 (比如窗口已不存在)
                    pass
                except Exception:
                    pass # 忽略其他可能的异常
                
                # 然后退出驱动服务
                try:
                    self.driver.quit()
                except WebDriverException:
                     # 忽略退出驱动时可能出现的错误 (比如服务已停止)
                     pass
                except Exception as e:
                    self.output_signal.emit(f"退出WebDriver时出错: {str(e)}")
            finally:
                self.driver = None
        
        # 取消未完成的异步任务
        if self.driver_future and not self.driver_future.done():
            try:
                self.driver_future.cancel()
            except Exception:
                pass
            self.driver_future = None
        
        # 再次确保驱动进程被终止 (作为最后的保障)
        self.kill_msedgedriver()

    def kill_msedgedriver(self):
        """终止所有 msedgedriver.exe 进程，不再终止 msedge.exe"""
        terminated_count = 0
        failed_count = 0
        driver_process_name = 'msedgedriver.exe' # 明确目标进程名
        
        try:
            # 首先尝试使用系统命令强制终止所有目标驱动进程
            if sys.platform == 'win32':
                try:
                    subprocess.call(f'taskkill /f /im {driver_process_name}', shell=True,
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass
            
            # 然后使用psutil逐个终止目标驱动进程
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    # 严格匹配进程名
                    if proc.info.get('name') and proc['name'].lower() == driver_process_name:
                        try:
                            proc.terminate()
                            gone, alive = psutil.wait_procs([proc], timeout=1) # 缩短等待时间
                            if proc in alive:
                                proc.kill()
                            terminated_count += 1
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                            failed_count += 1
                            # self.output_signal.emit(f"终止 {driver_process_name} (PID: {proc.info['pid']}) 时出错: {e}") # 减少冗余输出
                except Exception:
                    pass
            
            if terminated_count > 0 or failed_count > 0:
                 self.output_signal.emit(f"环境清理：已终止 {terminated_count} 个 {driver_process_name} 进程，失败 {failed_count} 个。")
        except Exception as e:
            self.output_signal.emit(f"清理驱动进程时出现异常: {e}")

    def check_network(self):
        """检测网络连接"""
        try:
            # 尝试访问国内和国外网站，增加可靠性
            urls_to_check = ["https://cn.bing.com/", "https://www.baidu.com/"]
            success_count = 0
            for url in urls_to_check:
                try:
                    response = requests.get(url, timeout=3)
                    if response.status_code == 200:
                        success_count += 1
                        break # 只要有一个成功就认为网络通畅
                except requests.RequestException:
                    continue
            
            if success_count > 0:
                msg = "网络连接正常。"
                # self.output_signal.emit(msg) # 减少冗余输出
                return True, msg
            else:
                msg = "网络连接可能存在问题，请检查网络设置。"
                # self.output_signal.emit(msg) # 网络问题作为警告，在run()中输出
                return False, msg
        except Exception as e:
            msg = f"网络检测时发生异常: {e}"
            # self.output_signal.emit(msg)
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
                    # 尝试从可执行文件获取版本（备用方案）
                    version_from_exe = self.get_edge_version_from_exe(edge_path)
                    if version_from_exe:
                         self.output_signal.emit(f"检测到Edge浏览器 (通过文件)，版本：{version_from_exe}")
                         self.edge_version = version_from_exe
                         return True, f"Edge浏览器已安装 (通过文件)，版本：{version_from_exe}"
                    else:
                        msg = "无法获取Edge浏览器版本（注册表和文件均失败）。"
                        self.output_signal.emit(msg)
                        return False, msg
            except Exception as e:
                msg = f"获取Edge浏览器版本时出错: {str(e)}"
                self.output_signal.emit(msg)
                return False, msg
        else:
            msg = ("未检测到Edge浏览器，请先安装Edge浏览器。 "
                   "下载链接: https://www.microsoft.com/edge")
            self.output_signal.emit(msg)
            return False, msg

    def check_edge_driver(self):
        """检测Edge WebDriver是否可用，否则尝试下载/更新"""
        if not self.edge_version:
            msg = "Edge版本未知，需先安装或检测Edge浏览器后再检测WebDriver。"
            self.output_signal.emit(msg)
            return False, msg

        # 在UI上只显示简单信息，不显示详细的缓存路径
        cache_dir = driver_manager.get_cache_dir()
        cache_file = driver_manager.get_cache_file_path()
        
        # 尝试加载缓存
        if not driver_manager.is_initialized():
            loaded = driver_manager.load_from_file()
            if loaded:
                self.output_signal.emit("已加载WebDriver缓存")
            # 不需要显示"未找到缓存"的消息

        # 首先检查驱动管理器是否已有可用驱动
        if driver_manager.is_initialized():
            driver_path = driver_manager.get_driver_path()
            
            # 检查版本是否匹配
            if driver_manager.version_matches(self.edge_version):
                self.output_signal.emit(f"使用已缓存的WebDriver配置")
                
                # 验证缓存的驱动是否可用
                result, driver = self._try_edge_with_unique_profile(driver_path=driver_path)
                if result:
                    self.driver = driver
                    self.output_signal.emit("缓存的WebDriver可用")
                    return True, "缓存的WebDriver可用。"
                else:
                    self.output_signal.emit("缓存的WebDriver不可用，需要重新配置")
            else:
                self.output_signal.emit(f"浏览器版本与WebDriver版本不匹配，需要更新")
            
            # 清理失败的尝试
            self.kill_msedgedriver()
            self.cleanup_temp_directories()
            # 如果版本不匹配或验证失败，重置缓存状态
            driver_manager.reset()
        
        self.output_signal.emit("检测Edge WebDriver...")
        
        # 1. 检测前清理驱动进程和旧目录
        self.kill_msedgedriver()
        self.cleanup_temp_directories() # 清理旧目录
        time.sleep(1) # 短暂等待

        driver_path = None
        # 尝试自动获取当前缓存的或已安装的驱动路径
        try:
            # 修复：EdgeChromiumDriverManager没有get_driver_path方法
            # 使用install方法获取已安装驱动的路径
            driver_path = EdgeChromiumDriverManager().install()
            if driver_path and os.path.exists(driver_path):
                 self.output_signal.emit("找到现有的WebDriver")
            else:
                 driver_path = None # 如果没找到或路径无效，则置空
        except Exception as e:
             self.output_signal.emit(f"查找WebDriver失败: {str(e)}")
             driver_path = None # 获取失败则置空

        # 2. 优先尝试核心策略：无头+唯一配置 (使用找到的或让Selenium自动找)
        try:
            result, driver = self._try_edge_with_unique_profile(driver_path=driver_path)
            if result:
                self.driver = driver
                self.output_signal.emit("Edge WebDriver已正确配置")
                
                # 将成功的驱动路径缓存到驱动管理器并保存
                if driver_path:
                    try:
                        success = driver_manager.set_driver_path(driver_path, self.edge_version)
                        if success:
                            self.output_signal.emit("WebDriver已缓存，下次将加速启动")
                        else:
                            self.output_signal.emit("缓存WebDriver失败")
                    except Exception as e:
                        self.output_signal.emit(f"保存WebDriver缓存失败: {str(e)}")
                
                return True, "Edge WebDriver已正确配置。"
            else:
                 # 如果第一次尝试失败，清理一下可能产生的临时文件和进程
                 self.kill_msedgedriver()
                 self.cleanup_temp_directories(enforce_limit=False) # 清理本次尝试的目录
                 time.sleep(0.5)
        except Exception as e:
            self.output_signal.emit(f"初始WebDriver启动尝试失败: {str(e)}")
            self.kill_msedgedriver() # 失败后清理
            self.cleanup_temp_directories(enforce_limit=False)
            time.sleep(0.5)

        # 3. 如果失败，尝试下载/更新驱动并重试核心策略
        try:
            self.output_signal.emit("尝试下载新的WebDriver...")
            # 强制重新下载最新的驱动
            try:
                driver_path = EdgeChromiumDriverManager(version=self.edge_version.split('.')[0]).install()
                self.output_signal.emit("WebDriver下载完成")
            except Exception as install_e:
                self.output_signal.emit(f"下载WebDriver出错，尝试使用默认版本")
                # 回退到无版本参数的下载方式
                driver_path = EdgeChromiumDriverManager().install()
                self.output_signal.emit("WebDriver下载完成")
            
            # 确保驱动程序可执行 (在非Windows上设置权限)
            if sys.platform != 'win32':
                try:
                    os.chmod(driver_path, 0o755)
                except Exception as chmod_e:
                     self.output_signal.emit(f"警告：设置WebDriver执行权限失败: {chmod_e}")

            # 下载后清理一次环境
            self.kill_msedgedriver()
            time.sleep(1)
            
            # 使用新驱动再次尝试核心策略
            result, driver = self._try_edge_with_unique_profile(driver_path=driver_path)
            if result:
                self.driver = driver
                self.output_signal.emit("WebDriver已更新并配置成功")
                
                # 将成功的驱动路径缓存到驱动管理器并保存
                try:
                    success = driver_manager.set_driver_path(driver_path, self.edge_version)
                    if success:
                        self.output_signal.emit("WebDriver已缓存，下次将加速启动")
                    else:
                        self.output_signal.emit("缓存WebDriver失败")
                except Exception as e:
                    self.output_signal.emit(f"保存WebDriver缓存出错")
                
                return True, "Edge WebDriver下载/更新并配置成功。"
            else:
                 # 如果使用新驱动的核心策略仍然失败，可以尝试一个备用策略（如随机端口）
                 self.output_signal.emit("尝试备用启动策略...")
                 self.kill_msedgedriver()
                 self.cleanup_temp_directories(enforce_limit=False)
                 time.sleep(0.5)
                 result, driver = self._try_edge_with_random_port(driver_path=driver_path)
                 if result:
                    self.driver = driver
                    self.output_signal.emit("WebDriver已使用备用策略配置成功")
                    
                    # 将成功的驱动路径缓存到驱动管理器并保存
                    try:
                        success = driver_manager.set_driver_path(driver_path, self.edge_version)
                        if success:
                            self.output_signal.emit("WebDriver已缓存，下次将加速启动")
                        else:
                            self.output_signal.emit("缓存WebDriver失败")
                    except Exception as e:
                        self.output_signal.emit("保存WebDriver缓存出错")
                    
                    return True, "Edge WebDriver下载/更新并使用备用策略配置成功。"
                 else:
                    # 如果备用策略也失败了，最后清理一次
                    self.kill_msedgedriver()
                    self.cleanup_temp_directories(enforce_limit=False)
                    
        except Exception as e:
            self.output_signal.emit(f"下载或使用新WebDriver失败: {str(e)}")
            # 失败后清理
            self.kill_msedgedriver()
            self.cleanup_temp_directories(enforce_limit=False)
            
        # 所有方法都失败
        self.output_signal.emit("无法配置Edge WebDriver。请确保Edge浏览器已正确安装并更新到最新版，或检查网络连接。")
        return False, "所有WebDriver启动策略均失败。"

    def kill_all_browser_processes(self):
        """彻底清理所有WebDriver相关进程，不再干预浏览器进程"""
        # 只清理msedgedriver进程
        self.kill_msedgedriver()
        
        # 确保任何残留的WebDriver连接被关闭
        self.quit_driver()

    def create_edge_driver_with_timeout(self, options, service=None, timeout=15): # 默认超时增加到15秒
        """并发启动Edge WebDriver，如超时则返回None"""
        def create_driver():
            driver_instance = None
            try:
                if service:
                    driver_instance = webdriver.Edge(service=service, options=options)
                else:
                    # 如果没有提供service，让selenium自动管理
                    driver_instance = webdriver.Edge(options=options)
                # 尝试访问一个简单页面，验证是否真正可用
                driver_instance.get("about:blank")
                return driver_instance
            except Exception as create_err:
                # 如果创建或访问页面失败，尝试清理可能启动的进程
                if driver_instance:
                    try:
                         driver_instance.quit()
                    except Exception:
                         pass
                # 抛出原始异常，由上层处理
                raise create_err

        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            future = None # 初始化future
            try:
                future = executor.submit(create_driver)
                driver = future.result(timeout=timeout)
                return driver
            except concurrent.futures.TimeoutError:
                self.output_signal.emit(f"启动Edge WebDriver超时 ({timeout}秒)")
                # 超时后尝试取消任务 (如果还在运行)
                if future and not future.done():
                    future.cancel()
                # 尝试强制清理可能的残留进程
                self.kill_msedgedriver()
                return None
            except Exception as e:
                # 将详细错误信息输出
                import traceback
                tb_str = traceback.format_exc()
                self.output_signal.emit(f"启动Edge WebDriver时发生错误: {e}\n详细信息:\n{tb_str}")
                # 出错后尝试强制清理可能的残留进程
                self.kill_msedgedriver()
                return None
            finally:
                 # 确保线程池关闭
                 executor.shutdown(wait=False)

    def find_edge_executable(self):
        """尝试找到Edge浏览器可执行文件"""
        # 优先使用 shutil.which
        edge_path = shutil.which('msedge') or shutil.which('edge')
        if edge_path and os.path.exists(edge_path):
            return edge_path
        
        # 如果which找不到，再尝试固定路径 (Windows特定)
        if sys.platform == 'win32':
            possible_paths = [
                os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Microsoft\\Edge\\Application\\msedge.exe"),
                os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft\\Edge\\Application\\msedge.exe"),
                os.path.join(os.environ.get("LocalAppData", ""), "Microsoft\\Edge\\Application\\msedge.exe") # 用户安装路径
            ]
            for path in possible_paths:
                if path and os.path.exists(path):
                    return path
        return None

    def get_edge_version_from_registry(self):
        """从Windows注册表获取Edge版本"""
        if sys.platform != 'win32':
             return None # 非Windows直接返回
        try:
            # 检查 HKEY_CURRENT_USER
            try:
                key_path = r"Software\\Microsoft\\Edge\\BLBeacon"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    if version:
                        return version
            except FileNotFoundError:
                pass
            except Exception as e_cu:
                 self.output_signal.emit(f"读取 HKEY_CURRENT_USER 注册表时出错: {e_cu}")

            # 检查 HKEY_LOCAL_MACHINE (64位视图)
            try:
                key_path = r"Software\\Microsoft\\Edge\\BLBeacon"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    if version:
                        return version
            except FileNotFoundError:
                pass
            except Exception as e_lm64:
                 self.output_signal.emit(f"读取 HKEY_LOCAL_MACHINE (64位) 注册表时出错: {e_lm64}")

            # 检查 HKEY_LOCAL_MACHINE (32位视图)
            try:
                key_path = r"Software\\Wow6432Node\\Microsoft\\Edge\\BLBeacon" # 注意路径变化
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY) as key:
                    version, _ = winreg.QueryValueEx(key, "version")
                    if version:
                        return version
            except FileNotFoundError:
                pass
            except Exception as e_lm32:
                 self.output_signal.emit(f"读取 HKEY_LOCAL_MACHINE (32位) 注册表时出错: {e_lm32}")

            # 如果上述都失败，尝试读取 Update 键 (更通用)
            try:
                 key_path = r"SOFTWARE\\Microsoft\\EdgeUpdate\\Clients\\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
                 # 先尝试 HKLM
                 try:
                      with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                           version, _ = winreg.QueryValueEx(key, "pv") # 读取 pv 值
                           if version:
                                return version
                 except FileNotFoundError:
                      pass
                 # 再尝试 HKCU
                 try:
                      with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                           version, _ = winreg.QueryValueEx(key, "pv")
                           if version:
                                return version
                 except FileNotFoundError:
                      pass
            except Exception as e_update:
                 self.output_signal.emit(f"读取 Edge Update 注册表键时出错: {e_update}")

            return None
        except Exception as e:
            raise Exception(f"从注册表获取Edge版本失败: {str(e)}")

    def get_edge_version_from_exe(self, exe_path):
         """尝试从Edge可执行文件获取版本信息 (Windows特定)"""
         if sys.platform != 'win32' or not exe_path:
             return None
         try:
             # 使用 PowerShell 命令获取文件版本信息
             command = f"(Get-Item \"{exe_path}\").VersionInfo.ProductVersion"
             result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
             version = result.stdout.strip()
             if version:
                 return version
         except (subprocess.CalledProcessError, FileNotFoundError, Exception) as e:
             # self.output_signal.emit(f"通过可执行文件获取版本失败: {e}") # 减少不必要的输出
             pass
         return None

    def _try_edge_with_unique_profile(self, driver_path=None):
        """策略: 使用唯一临时配置文件 (内部不再清理进程)"""
        unique_id = uuid.uuid4().hex
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        temp_dir_base = os.path.join(tempfile.gettempdir(), f"edge_temp_{timestamp}_{unique_id}")
        temp_dir = os.path.abspath(temp_dir_base) # 确保绝对路径
        profile_dir_created = False # 标记是否创建了目录

        try:
            # 确保父目录存在
            if not os.path.exists(os.path.dirname(temp_dir)):
                 try:
                     os.makedirs(os.path.dirname(temp_dir), exist_ok=True)
                 except OSError as e:
                     # 如果父目录创建失败，可能权限问题，记录错误并放弃
                     self.output_signal.emit(f"无法创建临时目录的父目录 {os.path.dirname(temp_dir)}: {e}")
                     return False, None

            # 尝试创建或验证目录
            retries = 3
            while retries > 0:
                try:
                    if not os.path.exists(temp_dir):
                        os.makedirs(temp_dir)
                        profile_dir_created = True
                        # self.output_signal.emit(f"创建临时目录: {temp_dir}")
                        break # 创建成功
                    elif os.access(temp_dir, os.W_OK):
                         # 目录存在且可写，可能被上次运行遗留，尝试使用
                         profile_dir_created = True # 标记以便清理
                         # self.output_signal.emit(f"使用已存在的临时目录: {temp_dir}")
                         break
                    else:
                        # 目录存在但不可写，尝试新名称
                         self.output_signal.emit(f"临时目录 {temp_dir} 不可写，尝试新名称...")
                         temp_dir = os.path.abspath(os.path.join(tempfile.gettempdir(), f"edge_temp_{timestamp}_{uuid.uuid4().hex}"))
                         retries -= 1
                except Exception as mkdir_e:
                     self.output_signal.emit(f"创建临时目录 {temp_dir} 时出错: {mkdir_e}, 重试...")
                     temp_dir = os.path.abspath(os.path.join(tempfile.gettempdir(), f"edge_temp_{timestamp}_{uuid.uuid4().hex}"))
                     retries -= 1
                     time.sleep(0.2)

            if not profile_dir_created:
                 self.output_signal.emit("无法创建可用的临时配置文件目录")
                 return False, None # 直接返回失败

            self.register_temp_dir(temp_dir) # 注册以便最终清理

            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            # 确保路径格式正确，特别是Windows
            options.add_argument(f'--user-data-dir={temp_dir}')
            # options.add_argument("--profile-directory=Default") # 使用Default可能更稳定，但也可能导致冲突，先注释掉
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage") # 添加这个参数
            options.add_argument("--disable-software-rasterizer") # 尝试禁用软件渲染
            options.add_argument("--log-level=3") # 减少浏览器日志级别
            options.add_argument("--output=/dev/null") # Linux/macOS 将日志输出到null
            if sys.platform == 'win32':
                 options.add_argument("--silent") # Windows 尝试静默模式

            # 排除可能导致问题的开关
            options.add_experimental_option("excludeSwitches", [
                "enable-logging", # 禁用冗余日志
                "enable-automation", # 有时会导致问题
                "disable-component-extensions-with-background-pages"
            ])
            options.add_experimental_option('useAutomationExtension', False) # 禁用自动化扩展

            service = None
            log_output = os.path.join(temp_dir, 'webdriver_service.log') # 使用注册的临时目录存放日志
            service_args = [
                '--log-level=WARNING', # 设置WebDriver服务日志级别
                f'--log-path={log_output}',
                '--append-log' # 追加日志而不是覆盖
            ]

            try:
                if driver_path and os.path.exists(driver_path):
                    # self.output_signal.emit(f"使用指定驱动路径: {driver_path}")
                    service = EdgeService(executable_path=driver_path, service_args=service_args)
                else:
                     # 如果没有指定路径或路径无效，让Selenium自动寻找或使用webdriver-manager缓存
                     # self.output_signal.emit("未指定有效驱动路径，让Selenium自动管理")
                     service = EdgeService(service_args=service_args)
            except WebDriverException as service_err:
                 # 如果服务初始化失败 (例如找不到驱动)，记录错误并返回
                 self.output_signal.emit(f"初始化WebDriver服务失败: {service_err}")
                 return False, None

            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=15) # 使用之前的超时设置

            return driver is not None, driver
        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            self.output_signal.emit(f"配置唯一配置文件策略时发生异常: {str(e)}\n详细信息:\n{tb_str}")
            # 不在此处清理临时目录，由外层调用者处理
            # 但要确保如果 driver 实例部分创建了，尝试 quit
            if 'driver' in locals() and locals()['driver']:
                 try:
                      locals()['driver'].quit()
                 except Exception:
                      pass
            return False, None

    def _try_edge_with_random_port(self, driver_path=None):
        """策略: 使用随机端口和唯一目录 (内部不再清理进程)"""
        port = random.randint(10000, 32000)
        unique_id = uuid.uuid4().hex
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        temp_dir_base = os.path.join(tempfile.gettempdir(), f"edge_port_{timestamp}_{unique_id}")
        temp_dir = os.path.abspath(temp_dir_base)
        port_dir_created = False

        try:
             # 确保父目录存在
            if not os.path.exists(os.path.dirname(temp_dir)):
                 try:
                     os.makedirs(os.path.dirname(temp_dir), exist_ok=True)
                 except OSError as e:
                     self.output_signal.emit(f"无法创建随机端口临时目录的父目录 {os.path.dirname(temp_dir)}: {e}")
                     return False, None

            # 尝试创建或验证目录 (与上一个策略类似)
            retries = 3
            while retries > 0:
                 try:
                     if not os.path.exists(temp_dir):
                         os.makedirs(temp_dir)
                         port_dir_created = True
                         # self.output_signal.emit(f"创建随机端口临时目录: {temp_dir}")
                         break
                     elif os.access(temp_dir, os.W_OK):
                          port_dir_created = True
                          # self.output_signal.emit(f"使用已存在的随机端口临时目录: {temp_dir}")
                          break
                     else:
                          self.output_signal.emit(f"随机端口临时目录 {temp_dir} 不可写，尝试新名称...")
                          temp_dir = os.path.abspath(os.path.join(tempfile.gettempdir(), f"edge_port_{timestamp}_{uuid.uuid4().hex}"))
                          retries -= 1
                 except Exception as mkdir_e:
                      self.output_signal.emit(f"创建随机端口临时目录 {temp_dir} 时出错: {mkdir_e}, 重试...")
                      temp_dir = os.path.abspath(os.path.join(tempfile.gettempdir(), f"edge_port_{timestamp}_{uuid.uuid4().hex}"))
                      retries -= 1
                      time.sleep(0.2)

            if not port_dir_created:
                 self.output_signal.emit("无法创建可用的随机端口临时目录")
                 return False, None

            self.register_temp_dir(temp_dir)

            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            options.add_argument(f'--user-data-dir={temp_dir}')
            # options.add_argument(f"--remote-debugging-port={port}") # 显式指定调试端口有时会与Service冲突，移除
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--log-level=3")
            options.add_argument("--output=/dev/null")
            if sys.platform == 'win32':
                 options.add_argument("--silent")
            options.add_experimental_option("excludeSwitches", [
                "enable-logging",
                "enable-automation",
                "disable-component-extensions-with-background-pages"
            ])
            options.add_experimental_option('useAutomationExtension', False)

            service = None
            log_output = os.path.join(temp_dir, 'webdriver_service_port.log')
            service_args=["--log-level=WARNING",
                 f'--log-path={log_output}',
                 '--append-log']

            try:
                if driver_path and os.path.exists(driver_path):
                    # self.output_signal.emit(f"使用指定驱动路径: {driver_path} 端口: {port}")
                    service = EdgeService(executable_path=driver_path, port=port, service_args=service_args)
                else:
                    # self.output_signal.emit(f"让Selenium自动管理驱动，端口: {port}")
                    service = EdgeService(port=port, service_args=service_args)
            except WebDriverException as service_err:
                 self.output_signal.emit(f"初始化WebDriver服务 (随机端口) 失败: {service_err}")
                 return False, None

            driver = self.create_edge_driver_with_timeout(options, service=service, timeout=15)

            return driver is not None, driver
        except Exception as e:
            import traceback
            tb_str = traceback.format_exc()
            self.output_signal.emit(f"配置随机端口策略时发生异常: {str(e)}\n详细信息:\n{tb_str}")
            # 清理
            if 'driver' in locals() and locals()['driver']:
                 try:
                      locals()['driver'].quit()
                 except Exception:
                      pass
            return False, None

    # 可以考虑移除 _try_edge_headless_incognito 和 _try_edge_with_alternative_service
    # 如果它们不再被 check_edge_driver 调用
    def _try_edge_headless_incognito(self, driver_path=None):
         # ... (此函数现在可能不再需要，可以删除或注释掉) ...
         # 如果保留，也需要移除内部的进程清理调用
         pass

    def _try_edge_with_alternative_service(self, driver_path=None):
         # ... (此函数现在可能不再需要，可以删除或注释掉) ...
         # 如果保留，也需要移除内部的进程清理调用
         pass
