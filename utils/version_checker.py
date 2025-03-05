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
                response.raise_for_status()
                
                data = response.json()
                # 增加字段校验
                if not all(key in data for key in ['tag_name', 'assets', 'body']):
                    raise ValueError("无效的API响应格式")
                
                self.latest_version = data['tag_name'].lstrip('v')
                self.assets = data['assets']
                self.release_notes = data['body'][:2000]  # 限制长度防止内存溢出
                return True
                
            except requests.exceptions.RequestException as e:
                if attempt == 2:
                    raise
                time.sleep(2 ** attempt)  # 指数退避重试

    def is_new_version_available(self):
        try:
            if not self.latest_version:
                self.check_latest_version()
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
            self.progress.emit(f"检查更新失败: {str(e)}")
            self.finished.emit(False, None, None, None, None)

    def _run_with_timeout(self, timeout):
        """带超时的执行方法"""
        result = [False, None, None, None, None]
        event = threading.Event()
        
        def target():
            try:
                self.version_checker.check_latest_version()
                if self.version_checker.is_new_version_available():
                    cpu_url, gpu_url = self.version_checker.get_download_urls()
                    result[:] = [True, self.version_checker.latest_version, 
                                cpu_url, gpu_url, self.version_checker.release_notes]
            except Exception as e:
                raise
            finally:
                event.set()
                
        thread = threading.Thread(target=target)
        thread.start()
        event.wait(timeout)
        
        if not event.is_set():
            raise TimeoutError("操作超时")
            
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
