# utils/version_checker.py

import requests
from utils.version import __version__
from PySide6.QtCore import QObject, Signal
import os
import threading
import cgi  # 引入 cgi 模块用于解析 Content-Disposition

# GitHub 相关配置
GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
OWNER = "FredericMN"  # 替换为你的 GitHub 用户名
REPO = "Game_ComplianceToolbox"  # 替换为你的仓库名称

# 更新用的 GitHub 令牌
ComplicanceToolbox_Update_Token = "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"

class VersionChecker:
    def __init__(self):
        self.current_version = __version__
        self.latest_version = None
        self.assets = []
        self.release_notes = None

    def check_latest_version(self):
        url = GITHUB_API_URL.format(owner=OWNER, repo=REPO)
        headers = {
            "Authorization": f"token {ComplicanceToolbox_Update_Token}"
        }
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                self.latest_version = data['tag_name'].lstrip('v')  # 去除 'v' 前缀
                self.assets = data['assets']
                self.release_notes = data['body']
                return True
            else:
                raise Exception(f"无法获取最新版本信息。状态码: {response.status_code}")
        except requests.RequestException as e:
            raise Exception(f"网络请求失败：{str(e)}")

    def is_new_version_available(self):
        if not self.latest_version:
            self.check_latest_version()
        return self.compare_versions(self.latest_version, self.current_version) > 0

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

    def get_download_asset_ids_and_names(self):
        """获取标准版和CUDA版的资产ID及名称"""
        cpu_asset_id = None
        cpu_asset_name = None
        gpu_asset_id = None
        gpu_asset_name = None
        for asset in self.assets:
            if asset['name'] == "ComplianceToolbox_standard.zip":
                cpu_asset_id = asset['id']
                cpu_asset_name = asset['name']
            elif asset['name'] == "ComplianceToolbox_cuda.7z":
                gpu_asset_id = asset['id']
                gpu_asset_name = asset['name']
        return cpu_asset_id, cpu_asset_name, gpu_asset_id, gpu_asset_name

class VersionCheckWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool, str, int, str, int, str, str)  # is_new_version, latest_version, cpu_asset_id, cpu_asset_name, gpu_asset_id, gpu_asset_name, release_notes

    def __init__(self, version_checker: VersionChecker):
        super().__init__()
        self.version_checker = version_checker

    def run(self):
        try:
            self.progress.emit("正在获取最新版本信息...")
            self.version_checker.check_latest_version()
            if self.version_checker.is_new_version_available():
                cpu_asset_id, cpu_asset_name, gpu_asset_id, gpu_asset_name = self.version_checker.get_download_asset_ids_and_names()
                if not cpu_asset_id and not gpu_asset_id:
                    self.progress.emit("未找到可用的更新文件。")
                    self.finished.emit(False, None, None, None, None, None, None)
                    return
                self.progress.emit(f"发现新版本：{self.version_checker.latest_version}")
                self.finished.emit(True, self.version_checker.latest_version, cpu_asset_id, cpu_asset_name, gpu_asset_id, gpu_asset_name, self.version_checker.release_notes)
            else:
                self.finished.emit(False, self.version_checker.latest_version, None, None, None, None, None)
        except Exception as e:
            self.progress.emit(f"检查更新失败: {str(e)}")
            self.finished.emit(False, None, None, None, None, None, None)

class DownloadWorker(QObject):
    progress = Signal(int)  # percent
    finished = Signal(bool, str)  # success, file_path

    def __init__(self, asset_id, owner, repo, token):
        super().__init__()
        self.asset_id = asset_id
        self.owner = owner
        self.repo = repo
        self.token = token

    def run(self):
        try:
            download_url = f"https://api.github.com/repos/{self.owner}/{self.repo}/releases/assets/{self.asset_id}"
            headers = {
                "Authorization": f"token {self.token}",
                "Accept": "application/octet-stream"
            }
            response = requests.get(download_url, headers=headers, stream=True, timeout=30, allow_redirects=True)

            if response.status_code != 200:
                self.progress.emit(0)
                self.finished.emit(False, f"下载失败，状态码: {response.status_code}")
                return

            total_length = response.headers.get('content-length')
            if total_length is None:
                self.progress.emit(0)
                self.finished.emit(False, "无法获取文件大小。")
                return
            total_length = int(total_length)

            # 从 Content-Disposition 头中提取文件名
            content_disposition = response.headers.get('Content-Disposition')
            if content_disposition:
                value, params = cgi.parse_header(content_disposition)
                filename = params.get('filename') or params.get('filename*')
                if not filename:
                    filename = f"downloaded_asset_{self.asset_id}"
            else:
                filename = f"downloaded_asset_{self.asset_id}"
            save_path = os.path.join(os.getcwd(), filename)
            with open(save_path, 'wb') as f:
                downloaded = 0
                for data in response.iter_content(chunk_size=4096):
                    if data:
                        f.write(data)
                        downloaded += len(data)
                        percent = int(downloaded * 100 / total_length)
                        self.progress.emit(percent)
            self.finished.emit(True, save_path)
        except requests.RequestException as e:
            self.progress.emit(0)
            self.finished.emit(False, f"下载失败：{str(e)}")
        except PermissionError:
            self.progress.emit(0)
            self.finished.emit(False, "文件写入权限不足。")
        except Exception as e:
            self.progress.emit(0)
            self.finished.emit(False, f"下载过程中发生错误：{str(e)}")
