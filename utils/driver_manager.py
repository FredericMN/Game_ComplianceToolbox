import os
import threading
import time
import json
import sys
import tempfile
from selenium.webdriver.edge.service import Service as EdgeService

class WebDriverManager:
    """WebDriver管理器：实现WebDriver路径的缓存和共享"""
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(WebDriverManager, cls).__new__(cls)
                cls._instance._driver_path = None
                cls._instance._driver_version = None
                cls._instance._last_check_time = 0
                cls._instance._is_initialized = False
                # 初始化时自动尝试加载缓存
                cls._instance.load_from_file()
            return cls._instance
    
    def set_driver_path(self, driver_path, driver_version=None):
        """设置WebDriver路径和版本"""
        if driver_path and os.path.exists(driver_path):
            self._driver_path = driver_path
            self._driver_version = driver_version
            self._is_initialized = True
            # 设置后自动保存到文件
            self.save_to_file()
            return True
        return False
    
    def get_driver_path(self):
        """获取已缓存的WebDriver路径"""
        if self._is_initialized and self._driver_path and os.path.exists(self._driver_path):
            return self._driver_path
        return None
    
    def get_driver_version(self):
        """获取已缓存的WebDriver版本"""
        return self._driver_version if self._is_initialized else None
    
    def is_initialized(self):
        """检查是否已初始化"""
        return self._is_initialized and self._driver_path is not None and os.path.exists(self._driver_path)
    
    def create_service(self, service_args=None):
        """创建EdgeService实例"""
        driver_path = self.get_driver_path()
        if driver_path:
            return EdgeService(executable_path=driver_path, service_args=service_args)
        return None
    
    def reset(self):
        """重置缓存状态"""
        self._driver_path = None
        self._driver_version = None
        self._is_initialized = False
        # 删除缓存文件
        try:
            cache_file = self.get_cache_file_path()
            if os.path.exists(cache_file):
                os.remove(cache_file)
        except Exception:
            pass
    
    def get_cache_dir(self):
        """获取缓存目录，优先使用程序根目录"""
        try:
            # 判断是否为打包后的程序
            if getattr(sys, 'frozen', False):
                # 打包后使用程序所在目录
                root_dir = os.path.dirname(sys.executable)
            else:
                # 开发环境使用当前工作目录
                root_dir = os.getcwd()
            
            # 在根目录创建.cache目录
            cache_dir = os.path.join(root_dir, '.webdriver_cache')  # 使用不同于临时文件的命名
            
            # 确保目录存在且可写
            try:
                if not os.path.exists(cache_dir):
                    os.makedirs(cache_dir, exist_ok=True)
                    print(f"Created cache directory: {cache_dir}")
                
                # 验证目录权限
                test_file = os.path.join(cache_dir, 'test_write.tmp')
                try:
                    with open(test_file, 'w') as f:
                        f.write('test')
                    os.remove(test_file)
                except (IOError, PermissionError) as e:
                    print(f"Cache directory not writable: {cache_dir}, error: {e}")
                    # 如果当前目录不可写，尝试使用临时目录
                    temp_cache = os.path.join(tempfile.gettempdir(), 'webdriver_cache_fixed')  # 使用固定命名，避开清理
                    os.makedirs(temp_cache, exist_ok=True)
                    print(f"Using temporary cache directory instead: {temp_cache}")
                    return temp_cache
            except (OSError, IOError, PermissionError) as e:
                print(f"Error creating cache directory {cache_dir}: {e}")
                # 回退到临时目录
                temp_cache = os.path.join(tempfile.gettempdir(), 'webdriver_cache_fixed')  # 使用固定命名，避开清理
                os.makedirs(temp_cache, exist_ok=True)
                print(f"Using temporary cache directory instead: {temp_cache}")
                return temp_cache
                
            return cache_dir
            
        except Exception as e:
            # 如果创建目录失败，回退到临时目录
            print(f"Error in get_cache_dir: {e}")
            try:
                temp_cache = os.path.join(tempfile.gettempdir(), 'webdriver_cache_fixed')  # 使用固定命名，避开清理
                os.makedirs(temp_cache, exist_ok=True)
                return temp_cache
            except Exception:
                # 最后的回退方案：使用模块所在目录
                module_dir = os.path.dirname(os.path.abspath(__file__))
                cache_dir = os.path.join(module_dir, '.cache')
                os.makedirs(cache_dir, exist_ok=True)
                return cache_dir
    
    def get_cache_file_path(self):
        """获取缓存文件的完整路径"""
        return os.path.join(self.get_cache_dir(), 'webdriver_cache.json')
    
    def save_to_file(self, file_path=None):
        """将WebDriver配置保存到文件"""
        if not file_path:
            file_path = self.get_cache_file_path()
        
        if not self._is_initialized or not self._driver_path:
            return False
            
        data = {
            'driver_path': self._driver_path,
            'driver_version': self._driver_version,
            'last_check_time': time.time()
        }
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"保存WebDriver缓存失败: {str(e)}")
            return False
    
    def load_from_file(self, file_path=None):
        """从文件加载WebDriver配置"""
        if not file_path:
            file_path = self.get_cache_file_path()
        
        if not os.path.exists(file_path):
            return False
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
            if 'driver_path' in data and os.path.exists(data['driver_path']):
                self._driver_path = data['driver_path'] 
                self._driver_version = data.get('driver_version')
                self._last_check_time = data.get('last_check_time', 0)
                self._is_initialized = True
                return True
        except Exception as e:
            print(f"加载WebDriver缓存失败: {str(e)}")
        
        return False
    
    def version_matches(self, browser_version):
        """检查缓存的WebDriver版本是否与浏览器版本匹配"""
        if not self._is_initialized or not self._driver_version or not browser_version:
            return False
        
        # 主要比较主版本号
        try:
            cached_major = self._driver_version.split('.')[0]
            browser_major = browser_version.split('.')[0]
            return cached_major == browser_major
        except Exception:
            return False
    
    def needs_update(self, browser_version, max_age_days=30):
        """检查是否需要更新WebDriver
        
        基于版本匹配和缓存时间判断
        """
        # 如果版本不匹配，需要更新
        if not self.version_matches(browser_version):
            return True
            
        # 检查缓存是否过期
        if self._last_check_time:
            current_time = time.time()
            cache_age_days = (current_time - self._last_check_time) / (24 * 3600)
            if cache_age_days > max_age_days:
                return True
        
        return False

# 全局单例实例
driver_manager = WebDriverManager() 