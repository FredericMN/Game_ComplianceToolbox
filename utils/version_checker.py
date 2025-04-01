# utils/version_checker.py

import requests
from utils.version import __version__
from PySide6.QtCore import QObject, Signal
import os
import threading
from packaging import version  # 新增标准版本库
import time

GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
OWNER = "FredericMN"  # 替换为你的 GitHub 用户名
REPO = "Game_ComplianceToolbox"  # 替换为你的仓库名称

class VersionChecker:
    def __init__(self):
        self.current_version = __version__
        self.latest_version = None
        self.assets = []
        self.release_notes = None

    def check_latest_version(self):
        url = GITHUB_API_URL.format(owner=OWNER, repo=REPO)
        for attempt in range(3):
            try:
                response = requests.get(url, timeout=(3.05, 30),
                                      headers={'Accept': 'application/vnd.github+json'},
                                      params={'per_page': 1})
                # 不直接使用 raise_for_status，改为检查状态码
                if response.status_code >= 400:
                    # 记录错误但不抛出异常
                    print(f"版本检查请求失败，状态码: {response.status_code}")
                    if attempt == 2:  # 最后一次尝试
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                # 解析数据前检查内容是否为JSON
                try:
                    data = response.json()
                except ValueError:
                    print("服务器响应不是有效的JSON格式")
                    if attempt == 2:
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                # 增加字段校验
                if not all(key in data for key in ['tag_name', 'assets', 'body']):
                    print("无效的API响应格式，缺少必要字段")
                    if attempt == 2:
                        return False
                    time.sleep(2 ** attempt)
                    continue
                
                self.latest_version = data['tag_name'].lstrip('v')
                self.assets = data['assets']
                self.release_notes = data['body'][:2000]  # 限制长度防止内存溢出
                return True
                
            except requests.exceptions.RequestException as e:
                print(f"请求异常: {str(e)}")
                if attempt == 2:
                    return False
                time.sleep(2 ** attempt)  # 指数退避重试
            except Exception as e:
                print(f"版本检查过程中发生未预期的异常: {str(e)}")
                if attempt == 2:
                    return False
                time.sleep(2 ** attempt)
        
        return False  # 所有尝试都失败

    def is_new_version_available(self):
        try:
            if not self.latest_version:
                success = self.check_latest_version()
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

    def __init__(self, version_checker: VersionChecker):
        super().__init__()
        self.version_checker = version_checker

    def run(self):
        try:
            self.progress.emit("正在获取最新版本信息...")
            
            # 增加超时控制
            result = self._run_with_timeout(30)
            self.finished.emit(*result)
            
        except Exception as e:
            print(f"版本检查工作线程异常: {str(e)}")  # 添加日志以便调试
            self.progress.emit(f"检查更新失败: {str(e)}")
            self.finished.emit(False, None, None, None, None)

    def _run_with_timeout(self, timeout):
        """带超时的执行方法"""
        result = [False, None, None, None, None]
        event = threading.Event()
        
        def target():
            try:
                success = self.version_checker.check_latest_version()
                if not success:
                    # 如果版本检查失败，直接设置结果并返回
                    print("版本检查失败，返回默认结果")
                    result[:] = [False, None, None, None, None]
                    return
                
                if self.version_checker.is_new_version_available():
                    cpu_url, gpu_url = self.version_checker.get_download_urls()
                    result[:] = [True, self.version_checker.latest_version, 
                                cpu_url, gpu_url, self.version_checker.release_notes]
            except Exception as e:
                # 捕获并记录异常，但不再抛出
                print(f"版本检查线程内部异常: {str(e)}")
                result[:] = [False, None, None, None, None]
            finally:
                event.set()
                
        thread = threading.Thread(target=target)
        thread.daemon = True  # 设置为守护线程，避免主程序退出时线程还在运行
        thread.start()
        
        # 等待线程完成或超时
        is_set = event.wait(timeout)
        if not is_set:
            print("版本检查超时")
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
        
    def cancel(self):
        self._cancel_flag = True
        
    def run(self):
        try:
            for attempt in range(3):
                if self._cancel_flag:
                    break
                
                try:
                    response = requests.get(
                        self.download_url, 
                        stream=True,
                        timeout=(3.05, 30),
                        headers={'Cache-Control': 'no-cache'}
                    )
                    response.raise_for_status()
                    
                    total_length = int(response.headers.get('content-length', 0))
                    filename = os.path.basename(self.download_url)
                    save_path = os.path.join(os.getcwd(), filename)
                    
                    self.downloaded_bytes = 0
                    chunk_size = 8192 * 16  # 增加块大小以提高下载速度
                    
                    with open(save_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=chunk_size):
                            if self._cancel_flag:
                                raise Exception("用户取消下载")
                                
                            if chunk:
                                f.write(chunk)
                                self.downloaded_bytes += len(chunk)
                                
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
