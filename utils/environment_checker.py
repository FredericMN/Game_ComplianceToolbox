# project-01/utils/environment_checker.py

import shutil
import subprocess
import sys
import os
from PySide6.QtCore import QObject, Signal, QTimer, QEventLoop
import requests
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
import winreg
import ctypes
import ctypes.wintypes

class EnvironmentChecker(QObject):
    output_signal = Signal(str)
    finished = Signal(bool)  # 修改为传递一个布尔值

    def __init__(self):
        super().__init__()
        self.edge_version = None
        self.has_errors = False  # 新增属性，记录是否有错误

    def run(self):
        try:
            self.step1_check_network()
            self.step2_check_edge_browser()
            self.step3_check_edge_driver()
        except Exception as e:
            self.output_signal.emit(f"环境检测过程中出现异常: {str(e)}")
            self.has_errors = True
        finally:
            self.finished.emit(self.has_errors)  # 传递 has_errors

    def step1_check_network(self):
        self.output_signal.emit("01-检测网络连接...")
        loop = QEventLoop()
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(loop.quit)
        timer.start(5000)  # 设置5秒超时

        try:
            response = requests.get("https://cn.bing.com/", timeout=3)
            if response.status_code == 200:
                self.output_signal.emit("01-检测网络连接-网络连接正常。")
            else:
                self.output_signal.emit("01-检测网络连接-网络异常，请检查网络连接。")
                self.has_errors = True
        except requests.RequestException:
            self.output_signal.emit("01-检测网络连接-网络异常，请检查网络连接。")
            self.has_errors = True
        finally:
            loop.quit()

    def step2_check_edge_browser(self):
        self.output_signal.emit("02-检测Edge浏览器安装情况...")
        loop = QEventLoop()
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(loop.quit)
        timer.start(5000)  # 设置5秒超时

        try:
            edge_path = self.find_edge_executable()
            if edge_path:
                self.output_signal.emit("02-检测Edge浏览器安装情况-检测到Edge浏览器已安装。")
                try:
                    version = self.get_edge_version_from_registry()
                    if version:
                        self.output_signal.emit(f"02-检测Edge浏览器安装情况-Edge浏览器版本: {version}")
                        self.edge_version = version
                    else:
                        self.output_signal.emit("02-检测Edge浏览器安装情况-无法获取Edge浏览器版本。")
                        self.has_errors = True
                except Exception as e:
                    self.output_signal.emit(f"02-检测Edge浏览器安装情况-无法获取Edge浏览器版本: {str(e)}")
                    self.has_errors = True
            else:
                self.output_signal.emit("02-检测Edge浏览器安装情况-未检测到Edge浏览器，请安装Edge浏览器。")
                self.output_signal.emit("02-检测Edge浏览器安装情况-下载链接: https://www.microsoft.com/edge")
                self.has_errors = True
                self.output_signal.emit("03-Edge浏览器WebDriver配置情况-需Edge浏览器安装后检测。")
        finally:
            loop.quit()

    def step3_check_edge_driver(self):
        if not self.edge_version:
            self.output_signal.emit("03-Edge浏览器WebDriver配置情况-需安装Edge浏览器后检测。")
            self.has_errors = True
            return

        self.output_signal.emit("03-Edge浏览器WebDriver配置情况-检测Edge WebDriver...")
        loop = QEventLoop()
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(loop.quit)
        timer.start(10000)  # 设置10秒超时

        driver = None  # 初始化 driver
        try:
            options = Options()
            options.use_chromium = True
            options.add_argument("--headless")
            # 尝试实例化webdriver.Edge以检查WebDriver是否正确配置
            driver = webdriver.Edge(options=options)
            self.output_signal.emit("03-Edge浏览器WebDriver配置情况-Edge WebDriver已正确配置。")
        except WebDriverException as e:
            self.output_signal.emit("03-Edge浏览器WebDriver配置情况-未正确配置Edge WebDriver。尝试下载并配置WebDriver...")
            try:
                driver_path = EdgeChromiumDriverManager().install()
                self.output_signal.emit(f"03-Edge浏览器WebDriver配置情况-Edge WebDriver已下载到: {driver_path}")
                # 设置Edge WebDriver服务
                service = EdgeService(driver_path)
                # 测试下载的WebDriver
                driver = webdriver.Edge(service=service, options=options)
                self.output_signal.emit("03-Edge浏览器WebDriver配置情况-Edge WebDriver已成功配置。")
            except Exception as download_e:
                self.output_signal.emit(f"03-Edge浏览器WebDriver配置情况-Edge WebDriver下载失败: {str(download_e)}")
                self.output_signal.emit("03-Edge浏览器WebDriver配置情况-请手动下载并配置WebDriver。下载链接: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/")
                self.has_errors = True
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception as quit_e:
                    self.output_signal.emit(f"03-Edge浏览器WebDriver配置情况-无法正常关闭WebDriver: {quit_e}")
            loop.quit()

    def find_edge_executable(self):
        """尝试找到Edge浏览器的可执行文件路径"""
        # Windows上常见的Edge安装路径
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
        """从Windows注册表获取Edge浏览器的版本信息"""
        try:
            registry_paths = [
                r"SOFTWARE\Microsoft\Edge\BLBeacon",
                r"SOFTWARE\WOW6432Node\Microsoft\Edge\BLBeacon"
            ]
            for reg_path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ) as key:
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                except FileNotFoundError:
                    continue
            # 尝试从 HKEY_LOCAL_MACHINE
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
