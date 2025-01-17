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

        # 自定义的环境检测项目列表，可按需扩展
        self.check_items = [
            ("网络连接检测", self.check_network),
            ("Edge浏览器检测", self.check_edge_browser),
            ("Edge WebDriver检测", self.check_edge_driver)
        ]

        # 用于收集结构化检查结果
        self.structured_results = []

    def run(self):
        """依次执行所有检测项，可考虑并发，但量少时顺序即可。"""
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

        # 发射结构化结果
        self.structured_result_signal.emit(self.structured_results)
        self.finished.emit(self.has_errors)

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
        driver = None
        try:
            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")

            # 尝试启动
            driver = self.create_edge_driver_with_timeout(options, service=None, timeout=10)
            if driver:
                self.output_signal.emit("Edge WebDriver已正确配置。")
                return True, "Edge WebDriver已正确配置。"
            else:
                msg = "Edge WebDriver启动超时或失败。"
                self.output_signal.emit(msg)
                return False, msg
        except WebDriverException:
            # 尝试安装
            self.output_signal.emit("未配置Edge WebDriver，尝试下载中...")
            try:
                driver_path = EdgeChromiumDriverManager().install()
                service = EdgeService(driver_path)
                driver = self.create_edge_driver_with_timeout(options, service=service, timeout=10)
                if driver:
                    msg = f"Edge WebDriver下载并配置成功：{driver_path}"
                    self.output_signal.emit(msg)
                    return True, msg
                else:
                    msg = "Edge WebDriver下载后仍无法启动。"
                    self.output_signal.emit(msg)
                    return False, msg
            except Exception as download_e:
                msg = (f"Edge WebDriver下载失败: {str(download_e)}，请手动配置。"
                       "下载链接: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
                self.output_signal.emit(msg)
                return False, msg
        finally:
            if driver:
                driver.quit()

    def create_edge_driver_with_timeout(self, options, service=None, timeout=10):
        """并发启动Edge WebDriver，如超时则返回None"""
        def create_driver():
            if service:
                return webdriver.Edge(service=service, options=options)
            else:
                return webdriver.Edge(options=options)

        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(create_driver)
            try:
                driver = future.result(timeout=timeout)
                return driver
            except concurrent.futures.TimeoutError:
                return None
            except Exception as e:
                self.output_signal.emit(f"启动Edge WebDriver时异常: {e}")
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
