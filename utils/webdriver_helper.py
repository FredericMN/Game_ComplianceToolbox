import sys
import os
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options

# 导入驱动管理器
try:
    from utils.driver_manager import driver_manager
except ImportError:
    # 如果导入失败，尝试相对导入
    try:
        from .driver_manager import driver_manager
    except ImportError:
        # 如果还是导入失败，创建一个简易的管理器
        class SimpleDriverManager:
            def set_driver_path(self, driver_path, driver_version=None):
                return True
            def get_driver_path(self):
                return None
            def is_initialized(self):
                return False
            def create_service(self, service_args=None):
                return None
            def load_from_file(self, file_path=None):
                return False
            def version_matches(self, browser_version):
                return False
        driver_manager = SimpleDriverManager()

class WebDriverHelper:
    """WebDriver辅助类：提供统一的WebDriver创建和管理功能"""
    
    _initialized = False
    
    @classmethod
    def init(cls):
        """初始化WebDriverHelper，加载驱动缓存"""
        if not cls._initialized:
            # 尝试加载缓存
            driver_manager.load_from_file()
            cls._initialized = True
            return driver_manager.is_initialized()
        return False
    
    @staticmethod
    def get_edge_version():
        """获取Edge浏览器版本"""
        try:
            # 尝试从注册表获取Edge版本
            if sys.platform == 'win32':
                import winreg
                try:
                    key_path = r"Software\\Microsoft\\Edge\\BLBeacon"
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                except:
                    pass
                    
                try:
                    key_path = r"Software\\Microsoft\\Edge\\BLBeacon"
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                except:
                    pass
            
            # 尝试使用命令行方法获取版本
            if sys.platform == 'win32':
                import subprocess
                try:
                    edge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
                    if not os.path.exists(edge_path):
                        edge_path = "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe"
                    
                    if os.path.exists(edge_path):
                        command = f"(Get-Item \"{edge_path}\").VersionInfo.ProductVersion"
                        result = subprocess.run(["powershell", "-Command", command], 
                                                capture_output=True, text=True, check=True)
                        version = result.stdout.strip()
                        if version:
                            return version
                except:
                    pass
        except:
            pass
            
        return None
    
    @staticmethod
    def create_driver(options=None, headless=True, progress_callback=None):
        """创建WebDriver实例
        
        Args:
            options: 自定义的WebDriver选项，如果为None则使用默认选项
            headless: 是否使用无头模式
            progress_callback: 进度回调函数，接收(message, percent)参数
            
        Returns:
            WebDriver实例或None
        """
        # 确保已初始化
        WebDriverHelper.init()
        
        def update_progress(message, percent=None):
            """更新进度信息"""
            if progress_callback:
                progress_callback(message, percent)
            else:
                print(message)
        
        # 创建默认选项
        if options is None:
            options = Options()
            options.use_chromium = True
            if headless:
                options.add_argument("--headless")
            options.add_argument("--disable-extensions")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            # 添加禁用自动化控制特征的选项
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
        
        # 尝试获取Edge版本
        edge_version = WebDriverHelper.get_edge_version()
        update_progress(f"检测到Edge浏览器版本: {edge_version or '未知'}", 8)
        
        # 尝试创建WebDriver
        try:
            update_progress("正在初始化Edge浏览器...", 5)
            
            # 显示缓存信息
            cache_dir = driver_manager.get_cache_dir()
            cache_file = driver_manager.get_cache_file_path()
            update_progress(f"WebDriver缓存目录: {cache_dir}", 6)
            update_progress(f"WebDriver缓存文件: {cache_file}", 7)
            
            # 从驱动管理器获取WebDriver服务
            edge_service = None
            if driver_manager.is_initialized():
                # 检查版本匹配
                if edge_version and not driver_manager.version_matches(edge_version):
                    update_progress(f"浏览器版本 ({edge_version}) 与缓存的WebDriver版本 ({driver_manager.get_driver_version()}) 不匹配，需要重新下载", 8)
                    # 重置缓存
                    driver_manager.reset()
                else:
                    update_progress("使用已缓存的WebDriver配置...", 10)
                    driver_path = driver_manager.get_driver_path()
                    service_args = [
                        '--log-level=WARNING',
                        '--append-log'
                    ]
                    edge_service = driver_manager.create_service(service_args)
                    update_progress(f"已获取缓存的WebDriver路径: {driver_path}", 15)
            
            if edge_service:
                # 使用缓存的服务创建WebDriver
                update_progress("使用缓存的WebDriver创建浏览器实例...", 20)
                start_time = time.time()
                driver = webdriver.Edge(service=edge_service, options=options)
                elapsed_time = time.time() - start_time
                update_progress(f"浏览器实例创建成功！用时 {elapsed_time:.2f} 秒", 30)
            else:
                # 如果没有缓存的服务，使用传统方法
                update_progress("未找到缓存的WebDriver，正在通过WebDriver Manager下载或配置...", 10)
                
                # 这里可以添加对其他驱动管理器方法的支持，如果需要
                from webdriver_manager.microsoft import EdgeChromiumDriverManager
                
                update_progress("正在下载或安装WebDriver...", 15)
                start_time = time.time()
                
                # 如果有获取到Edge版本，使用对应版本的驱动
                try:
                    if edge_version:
                        update_progress(f"尝试使用Edge版本 {edge_version} 下载匹配的WebDriver", 16)
                        # 使用正确的参数名
                        driver_path = EdgeChromiumDriverManager(version=edge_version.split('.')[0]).install()
                    else:
                        update_progress("使用默认版本下载WebDriver", 16)
                        driver_path = EdgeChromiumDriverManager().install()
                except Exception as install_e:
                    update_progress(f"下载WebDriver时出错: {str(install_e)}，尝试使用默认版本", 17)
                    # 回退到无参数的下载
                    driver_path = EdgeChromiumDriverManager().install()
                    
                elapsed_download_time = time.time() - start_time
                update_progress(f"WebDriver下载/安装完成! 用时 {elapsed_download_time:.2f} 秒，路径: {driver_path}", 20)
                
                # 创建WebDriver
                update_progress("创建浏览器实例...", 25)
                start_time = time.time()
                
                try:
                    driver = webdriver.Edge(
                        service=EdgeService(driver_path),
                        options=options
                    )
                    elapsed_time = time.time() - start_time
                    update_progress(f"浏览器实例创建成功（使用新配置）！总用时 {elapsed_time:.2f} 秒", 30)
                    
                    # 缓存成功创建的WebDriver路径
                    try:
                        driver_manager.set_driver_path(driver_path, edge_version)
                        update_progress(f"已缓存WebDriver配置，后续将加速启动。缓存路径: {driver_manager.get_cache_file_path()}", 35)
                    except Exception as cache_e:
                        update_progress(f"缓存WebDriver配置时出错: {str(cache_e)}", 35)
                except Exception as driver_e:
                    update_progress(f"创建WebDriver实例失败: {str(driver_e)}", 0)
                    # 尝试强制关闭所有msedgedriver进程
                    WebDriverHelper.kill_msedgedriver()
                    update_progress("已尝试清理msedgedriver进程", 0)
                    raise  # 重新抛出异常
            
            return driver
        except Exception as e:
            update_progress(f"创建WebDriver时出错: {str(e)}", 0)
            # 尝试强制关闭所有msedgedriver进程
            WebDriverHelper.kill_msedgedriver()
            return None
    
    @staticmethod
    def quit_driver(driver):
        """安全地关闭WebDriver"""
        if driver:
            try:
                driver.quit()
                return True
            except Exception:
                return False
        return False
    
    @staticmethod
    def kill_msedgedriver():
        """终止所有msedgedriver进程"""
        try:
            # 在Windows上使用taskkill
            if sys.platform == 'win32':
                try:
                    import subprocess
                    # 使用tasklist检查是否有msedgedriver进程
                    cmd = 'tasklist /FI "IMAGENAME eq msedgedriver.exe" /FO CSV /NH'
                    output = subprocess.check_output(cmd, shell=True, text=True)
                    if 'msedgedriver.exe' in output:
                        # 如果进程存在，强制终止
                        subprocess.call('taskkill /f /im msedgedriver.exe', 
                                    shell=True, stdout=subprocess.DEVNULL, 
                                    stderr=subprocess.DEVNULL)
                        print("已终止msedgedriver.exe进程 (通过taskkill)")
                except Exception as e:
                    print(f"使用taskkill终止msedgedriver进程时出错: {str(e)}")
            
            # 使用psutil逐个终止进程 (跨平台方案)
            try:
                import psutil
                terminated = 0
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        process_name = proc.info.get('name', '').lower()
                        if 'msedgedriver' in process_name:
                            proc.kill()
                            terminated += 1
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        pass
                
                if terminated > 0:
                    print(f"已终止 {terminated} 个msedgedriver进程 (通过psutil)")
            except Exception as e:
                print(f"使用psutil终止msedgedriver进程时出错: {str(e)}")
            
            return True
        except Exception as e:
            print(f"终止msedgedriver进程时出错: {str(e)}")
            return False
            
# 初始化WebDriverHelper
WebDriverHelper.init() 