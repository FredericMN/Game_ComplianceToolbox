# utils/version_checker.py

import requests
from utils.version import __version__
from PySide6.QtCore import QObject, Signal
import os
import threading
from packaging import version  # 新增标准版本库
import time
import json  # 导入json模块处理配置文件
import datetime  # 用于处理缓存时间

GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
OWNER = "FredericMN"  # 替换为你的 GitHub 用户名
REPO = "Game_ComplianceToolbox"  # 替换为你的仓库名称
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'version_cache.json')

# 默认GitHub令牌
DEFAULT_TOKEN = "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"

# 缓存有效期（小时）
CACHE_LIFETIME = 1

def save_config(config):
    """保存配置到文件"""
    try:
        with open(CONFIG_PATH, 'w') as f:
            json.dump(config, f)
        return True
    except Exception as e:
        print(f"保存配置失败: {str(e)}")
        return False

def load_config():
    """从文件加载配置"""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载配置失败: {str(e)}")
    # 创建默认配置
    default_config = {"github_token": DEFAULT_TOKEN}
    save_config(default_config)
    return default_config

def save_version_cache(version_data):
    """保存版本信息到缓存"""
    try:
        cache_data = {
            "timestamp": datetime.datetime.now().isoformat(),
            "data": version_data
        }
        with open(CACHE_PATH, 'w') as f:
            json.dump(cache_data, f)
        return True
    except Exception as e:
        print(f"保存版本缓存失败: {str(e)}")
        return False

def load_version_cache():
    """从缓存加载版本信息"""
    if not os.path.exists(CACHE_PATH):
        return None
        
    try:
        with open(CACHE_PATH, 'r') as f:
            cache = json.load(f)
            
        # 处理新旧两种格式的缓存
        if "data" in cache and isinstance(cache["data"], dict):
            # 新格式
            data = cache["data"]
            timestamp_str = cache.get("timestamp", "")
        else:
            # 旧格式 - 整个对象就是数据
            data = cache
            timestamp_str = data.get("timestamp", "")
            
        # 无论哪种格式，确保返回数据包含所需字段
        if not all(key in data for key in ["latest_version", "assets"]):
            print("缓存数据格式不完整，重新获取")
            return None
            
        # 检查缓存是否过期
        try:
            timestamp = datetime.datetime.fromisoformat(timestamp_str)
            now = datetime.datetime.now()
            diff = now - timestamp
            
            # 如果缓存未过期，返回数据
            if diff.total_seconds() < CACHE_LIFETIME * 3600:
                return data
            else:
                print(f"缓存已过期 ({diff.total_seconds()//3600:.1f}小时)，重新获取")
        except:
            print("无法解析缓存时间戳，重新获取")
            
    except Exception as e:
        print(f"加载版本缓存失败: {str(e)}")
        try:
            # 如果缓存文件损坏，尝试删除
            os.remove(CACHE_PATH)
            print("已删除损坏的缓存文件")
        except:
            pass
    
    return None

class VersionChecker:
    def __init__(self):
        self.current_version = __version__
        self.latest_version = None
        self.assets = []
        self.release_notes = None
        self.config = load_config()

    def check_latest_version(self, force_check=False):
        """检查最新版本，支持强制检查和缓存"""
        # 如果不是强制检查，尝试从缓存加载
        if not force_check:
            cached_data = load_version_cache()
            if cached_data:
                print("从缓存加载版本信息")
                self.latest_version = cached_data.get("latest_version")
                self.assets = cached_data.get("assets", [])
                self.release_notes = cached_data.get("release_notes")
                return True
        
        # 诊断网络连接
        try:
            print("诊断: 测试网络连接...")
            test_response = requests.get("https://api.github.com", timeout=(3, 10))
            print(f"诊断: GitHub API 可访问性: {test_response.status_code}")
        except Exception as e:
            print(f"诊断: 无法连接到GitHub: {str(e)}")
        
        url = GITHUB_API_URL.format(owner=OWNER, repo=REPO)
        print(f"诊断: 尝试访问 {url}")
        
        for attempt in range(3):
            try:
                # 构建请求头，加入令牌
                token = self.config.get("github_token", "")
                headers = {
                    'Accept': 'application/vnd.github+json'
                }
                
                if token:
                    print(f"诊断: 使用令牌 {token[:5]}...{token[-5:]}")
                    headers['Authorization'] = f'token {token}'
                else:
                    print("诊断: 未找到有效令牌")
                
                print(f"诊断: 请求 (第{attempt+1}次尝试)")
                response = requests.get(url, timeout=(5, 30),
                                      headers=headers,
                                      params={'per_page': 1})
                
                print(f"诊断: 响应状态码: {response.status_code}")
                
                # 检查API限制状态
                remaining = int(response.headers.get('X-RateLimit-Remaining', '60'))
                reset_time = int(response.headers.get('X-RateLimit-Reset', '0'))
                current_time = time.time()
                minutes_to_reset = max(0, (reset_time - current_time) // 60)
                
                print(f"诊断: API限制剩余: {remaining}, 重置时间: {minutes_to_reset}分钟后")
                
                if remaining < 10:  # 如果剩余次数较少，记录警告
                    print(f"GitHub API 剩余请求次数较少: {remaining}，将在{minutes_to_reset}分钟后重置")
                
                # 不直接使用 raise_for_status，改为检查状态码
                if response.status_code >= 400:
                    # 记录错误但不抛出异常
                    error_detail = response.text[:200] if response.text else "无详细信息"
                    print(f"版本检查请求失败，状态码: {response.status_code}, 详情: {error_detail}")
                    
                    if response.status_code == 401:
                        print("诊断: 令牌认证失败，请检查令牌是否有效")
                    elif response.status_code == 403 and 'X-RateLimit-Remaining' in response.headers and int(response.headers['X-RateLimit-Remaining']) == 0:
                        print("已达到GitHub API请求限制")
                    elif response.status_code == 404:
                        print("诊断: 资源不存在，请检查仓库名称和所有者是否正确")
                        
                    if attempt == 2:  # 最后一次尝试
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                # 解析数据前检查内容是否为JSON
                try:
                    data = response.json()
                    print("诊断: 成功解析JSON响应")
                except ValueError:
                    print(f"诊断: 服务器响应不是有效的JSON格式: {response.text[:200]}")
                    if attempt == 2:
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                # 增加字段校验
                if not all(key in data for key in ['tag_name', 'assets', 'body']):
                    missing_keys = [k for k in ['tag_name', 'assets', 'body'] if k not in data]
                    print(f"诊断: 无效的API响应格式，缺少必要字段: {missing_keys}")
                    print(f"诊断: 响应内容: {str(data)[:200]}")
                    if attempt == 2:
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                self.latest_version = data['tag_name'].lstrip('v')
                self.assets = data['assets']
                self.release_notes = data['body'][:2000]  # 限制长度防止内存溢出
                
                print(f"诊断: 成功获取版本信息: {self.latest_version}")
                
                # 保存到缓存
                cache_data = {
                    "latest_version": self.latest_version,
                    "assets": self.assets,
                    "release_notes": self.release_notes
                }
                save_version_cache(cache_data)
                
                return True
                
            except requests.exceptions.RequestException as e:
                print(f"诊断: 请求异常: {str(e)}, 类型: {type(e).__name__}")
                if attempt == 2:
                    return False
                time.sleep(2 ** attempt)  # 指数退避重试
            except Exception as e:
                print(f"诊断: 版本检查过程中发生未预期的异常: {str(e)}, 类型: {type(e).__name__}")
                import traceback
                print(f"诊断: 异常堆栈: {traceback.format_exc()}")
                if attempt == 2:
                    return False
                time.sleep(2 ** attempt)
        
        return False  # 所有尝试都失败

    def is_new_version_available(self, force_check=False):
        try:
            if not self.latest_version:
                success = self.check_latest_version(force_check)
                if not success:
                    return False  # 如果检查版本失败，直接返回没有新版本
                
            return version.parse(self.latest_version) > version.parse(self.current_version)
        except Exception as e:
            print(f"版本比较错误: {str(e)}")
            return False

    def compare_versions(self, v1, v2):
        def parse_version(v):
            return [int(x) for x in v.split('.')]
        v1_nums = parse_version(v1)
        v2_nums = parse_version(v2)
        # 补齐版本号长度，例如将 [1, 0] 补齐为 [1, 0, 0]
        length = max(len(v1_nums), len(v2_nums))
        v1_nums.extend([0] * (length - len(v1_nums)))
        v2_nums.extend([0] * (length - len(v2_nums)))
        # 比较每一部分
        for a, b in zip(v1_nums, v2_nums):
            if a > b:
                return 1
            elif a < b:
                return -1
        return 0

    def get_download_urls(self):
        """获取标准版和CUDA版的下载链接"""
        cpu_url = None
        gpu_url = None
        for asset in self.assets:
            if asset['name'] == "ComplianceToolbox_standard.zip":
                cpu_url = asset['browser_download_url']
            elif asset['name'] == "ComplianceToolbox_cuda.7z":
                gpu_url = asset['browser_download_url']
        return cpu_url, gpu_url

class VersionCheckWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool, str, str, str, str)  # is_new_version, latest_version, cpu_download_url, gpu_download_url, release_notes

    def __init__(self, version_checker: VersionChecker, force_check=False):
        super().__init__()
        self.version_checker = version_checker
        self.force_check = force_check

    def run(self):
        try:
            self.progress.emit("正在获取最新版本信息...")
            
            # 显示是否使用强制检查模式
            if self.force_check:
                print("使用强制检查模式，忽略缓存")
            
            # 增加超时控制
            result = self._run_with_timeout(60)  # 增加超时时间
            self.finished.emit(*result)
            
        except Exception as e:
            error_msg = f"版本检查工作线程异常: {str(e)}"
            print(error_msg)  # 添加日志以便调试
            
            # 添加更详细的异常信息
            import traceback
            print(f"异常堆栈: {traceback.format_exc()}")
            
            self.progress.emit(f"检查更新失败: {str(e)}")
            self.finished.emit(False, None, None, None, None)

    def _run_with_timeout(self, timeout):
        """带超时的执行方法"""
        result = [False, None, None, None, None]
        event = threading.Event()
        
        def target():
            try:
                print("开始版本检查...")
                success = self.version_checker.check_latest_version(self.force_check)
                
                if not success:
                    # 如果版本检查失败，直接设置结果并返回
                    error_msg = "版本检查失败，返回默认结果"
                    print(error_msg)
                    self.progress.emit(f"无法获取最新版本信息: {error_msg}")
                    result[:] = [False, None, None, None, None]
                    return
                
                # 检查是否有新版本
                has_new_version = self.version_checker.is_new_version_available()
                if has_new_version:
                    print(f"发现新版本: {self.version_checker.latest_version}")
                    cpu_url, gpu_url = self.version_checker.get_download_urls()
                    result[:] = [True, self.version_checker.latest_version, 
                                cpu_url, gpu_url, self.version_checker.release_notes]
                else:
                    print(f"已是最新版本: {self.version_checker.current_version}")
                    self.progress.emit(f"已是最新版本: {self.version_checker.current_version}")
                    result[:] = [False, self.version_checker.latest_version, None, None, None]
                
            except Exception as e:
                # 捕获并记录异常，但不再抛出
                error_msg = f"版本检查线程内部异常: {str(e)}"
                print(error_msg)
                import traceback
                print(f"异常堆栈: {traceback.format_exc()}")
                self.progress.emit(f"检查更新出错: {str(e)}")
                result[:] = [False, None, None, None, None]
            finally:
                event.set()
                
        thread = threading.Thread(target=target)
        thread.daemon = True  # 设置为守护线程，避免主程序退出时线程还在运行
        thread.start()
        
        # 等待线程完成或超时
        is_set = event.wait(timeout)
        if not is_set:
            error_msg = f"版本检查超时，超过{timeout}秒没有响应"
            print(error_msg)
            self.progress.emit(f"获取版本信息超时: {error_msg}")
            # 如果超时，不抛出异常，只返回默认结果
            return [False, None, None, None, None]
            
        return result

class DownloadWorker(QObject):
    progress = Signal(int, str)
    finished = Signal(bool, str)
    
    def __init__(self, download_url, parent=None):
        super().__init__(parent)
        self.download_url = download_url
        self._cancel_flag = False
        self.downloaded_bytes = 0  # 添加已下载字节计数
        self.config = load_config()  # 加载配置获取令牌
        
    def cancel(self):
        self._cancel_flag = True
        
    def run(self):
        try:
            for attempt in range(3):
                if self._cancel_flag:
                    break
                
                try:
                    # 检查是否为GitHub URL
                    is_github_url = "github.com" in self.download_url or "githubusercontent.com" in self.download_url
                    
                    # 构建请求头
                    headers = {'Cache-Control': 'no-cache'}
                    # 如果是GitHub URL且有令牌，添加认证
                    if is_github_url and self.config.get("github_token"):
                        headers['Authorization'] = f'token {self.config.get("github_token", "")}'
                    
                    # 添加断点续传支持
                    filename = os.path.basename(self.download_url)
                    save_path = os.path.join(os.getcwd(), filename)
                    if os.path.exists(save_path):
                        file_size = os.path.getsize(save_path)
                        if file_size > 0:
                            headers['Range'] = f'bytes={file_size}-'
                            self.downloaded_bytes = file_size
                    
                    response = requests.get(
                        self.download_url, 
                        stream=True,
                        timeout=(3.05, 60),  # 增加下载超时时间
                        headers=headers
                    )
                    response.raise_for_status()
                    
                    # 检查是否支持断点续传
                    is_resume = response.status_code == 206
                    
                    total_length = int(response.headers.get('content-length', 0))
                    if is_resume:
                        total_length += self.downloaded_bytes
                    else:
                        # 如果服务器不支持断点续传，从头开始下载
                        self.downloaded_bytes = 0
                        save_path = os.path.join(os.getcwd(), filename)
                    
                    mode = 'ab' if is_resume else 'wb'
                    chunk_size = 8192 * 16  # 增加块大小以提高下载速度
                    
                    with open(save_path, mode) as f:
                        for chunk in response.iter_content(chunk_size=chunk_size):
                            if self._cancel_flag:
                                raise Exception("用户取消下载")
                                
                            if chunk:
                                f.write(chunk)
                                if not is_resume:
                                    self.downloaded_bytes += len(chunk)
                                else:
                                    self.downloaded_bytes = file_size + len(chunk)
                                
                                if total_length > 0:
                                    percent = min(99, int(self.downloaded_bytes * 100 / total_length))
                                else:
                                    percent = 0
                                    
                                self.progress.emit(
                                    percent,
                                    f"下载中... {self.downloaded_bytes//1024//1024:.1f}MB"
                                    f"{f'/{total_length//1024//1024:.1f}MB' if total_length else ''}"
                                )
                                
                    if total_length > 0 and self.downloaded_bytes != total_length:
                        raise Exception("文件下载不完整")
                        
                    self.finished.emit(True, save_path)
                    return
                    
                except requests.exceptions.RequestException as e:
                    if attempt == 2:  # 最后一次尝试失败
                        raise
                    time.sleep(2 ** attempt)  # 指数退避
                    
        except Exception as e:
            error_msg = f"下载失败：{str(e)}"
            try:
                if 'save_path' in locals():
                    os.remove(save_path)
            except:
                pass
            self.progress.emit(0, error_msg)
            self.finished.emit(False, None)
