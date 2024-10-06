# utils/version_checker.py

import requests
from utils.version import __version__
from PySide6.QtCore import QObject, Signal
import os
import threading

GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
OWNER = "FredericMN"  # 替换为你的 GitHub 用户名
REPO = "Game_ComplianceToolbox"       # 替换为你的仓库名称

class VersionChecker:
    def __init__(self):
        self.current_version = __version__
        self.latest_version = None
        self.assets = []
        self.release_notes = None

    def check_latest_version(self):
        url = GITHUB_API_URL.format(owner=OWNER, repo=REPO)
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                self.latest_version = data['tag_name'].lstrip('v')  # 去除 'v' 前缀
                self.assets = data['assets']
                self.release_notes = data['body']
                return True
            else:
                raise Exception("无法获取最新版本信息。")
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

    def get_download_url(self):
        # 查找名为 "ComplianceToolbox.zip" 的资产
        for asset in self.assets:
            if asset['name'] == "ComplianceToolbox.zip":
                return asset['browser_download_url']
        return None

class VersionCheckWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool, str, str, str)  # is_new_version, latest_version, download_url, release_notes

    def __init__(self, version_checker: VersionChecker):
        super().__init__()
        self.version_checker = version_checker

    def run(self):
        try:
            self.progress.emit("正在获取最新版本信息...")
            self.version_checker.check_latest_version()
            if self.version_checker.is_new_version_available():
                download_url = self.version_checker.get_download_url()
                if not download_url:
                    self.progress.emit("未找到可用的更新文件。")
                    self.finished.emit(False, None, None, None)
                    return
                self.progress.emit(f"发现新版本：{self.version_checker.latest_version}")
                self.finished.emit(True, self.version_checker.latest_version, download_url, self.version_checker.release_notes)
            else:
                self.progress.emit("当前已是最新版本。")
                self.finished.emit(False, self.version_checker.latest_version, None, None)
        except Exception as e:
            self.progress.emit(f"检查更新失败: {str(e)}")
            self.finished.emit(False, None, None, None)

class DownloadWorker(QObject):
    progress = Signal(str)
    finished = Signal(bool, str)  # success, file_path

    def __init__(self, download_url):
        super().__init__()
        self.download_url = download_url

    def run(self):
        try:
            response = requests.get(self.download_url, stream=True, timeout=30)
            total_length = response.headers.get('content-length')
            if total_length is None:
                self.progress.emit("无法获取文件大小。")
                self.finished.emit(False, None)
                return
            total_length = int(total_length)
            filename = os.path.basename(self.download_url)
            save_path = os.path.join(os.getcwd(), filename)
            with open(save_path, 'wb') as f:
                downloaded = 0
                for data in response.iter_content(chunk_size=4096):
                    if data:
                        f.write(data)
                        downloaded += len(data)
                        percent = int(downloaded * 100 / total_length)
                        self.progress.emit(f"下载进度：{percent}%")
            self.finished.emit(True, save_path)
        except requests.RequestException as e:
            self.progress.emit(f"下载失败：{str(e)}")
            self.finished.emit(False, None)
