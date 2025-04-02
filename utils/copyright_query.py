import os
import pandas as pd
import time
import re
import random
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 导入任务管理器（如果不存在需要先导入）
try:
    from utils.task_manager import task_manager
except ImportError:
    # 简易任务管理器实现，避免导入错误
    class SimpleTaskManager:
        def register_webdriver(self, driver):
            pass
        
        def unregister_webdriver(self, driver):
            pass
    
    task_manager = SimpleTaskManager()

class CopyrightQuery:
    """著作权人查询工具"""
    
    def __init__(self):
        # 用户代理列表，用于模拟不同浏览器
        self.user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:90.0) Gecko/20100101 Firefox/90.0",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59"
        ]
        
    def random_delay(self, min_sec=1.0, max_sec=3.0):
        """避免爬取过快"""
        time.sleep(random.uniform(min_sec, max_sec))
        
    def identify_game_name_column(self, excel_path):
        """识别Excel中的游戏名称列"""
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)
            
            # 打印所有列名以便调试
            print("Excel文件中的所有列名:")
            for i, col in enumerate(df.columns):
                print(f"{i+1}. {col}")
            
            # 尝试自动识别游戏名称列
            # 策略1: 查找可能包含"游戏"或"名称"的列
            possible_name_columns = []
            for col in df.columns:
                if any(keyword in str(col) for keyword in ["游戏名称", "名称", "名字"]):
                    possible_name_columns.append(col)
            
            # 如果找到可能的列，返回第一个
            if possible_name_columns:
                chosen_column = possible_name_columns[0]
                print(f"自动识别到可能的游戏名称列: {chosen_column}")
                
                # 返回列名和对应的数据
                return chosen_column, df[chosen_column].dropna().tolist()
            
            # 如果没有找到，提示用户手动选择
            else:
                print("未能自动识别游戏名称列，请手动选择:")
                for i, col in enumerate(df.columns):
                    sample_values = df[col].dropna().head(3).tolist()
                    print(f"{i+1}. {col} (样例值: {sample_values})")
                
                choice = int(input("请输入列号: ")) - 1
                chosen_column = df.columns[choice]
                print(f"已选择: {chosen_column}")
                
                # 返回列名和对应的数据
                return chosen_column, df[chosen_column].dropna().tolist()
                
        except Exception as e:
            print(f"读取Excel文件出错: {str(e)}")
            return None, []
    
    def test_selenium_browser(self, game_name):
        """使用Selenium测试对企查查和天眼查的请求"""
        # 对游戏名称进行URL编码
        encoded_name = quote(game_name)
        
        # 企查查搜索URL
        qcc_url = f"https://www.qcc.com/web/search?key={encoded_name}"
        # 天眼查搜索URL
        tianyancha_url = f"https://www.tianyancha.com/nsearch?key={encoded_name}"
        
        print(f"\n测试请求搜索: {game_name}")
        print(f"编码后的搜索词: {encoded_name}")
        
        # 创建Edge浏览器驱动
        opt = webdriver.EdgeOptions()
        
        # 添加一些性能优化选项
        opt.add_argument('--disable-extensions')
        opt.add_argument('--no-sandbox')
        opt.add_argument('--disable-dev-shm-usage')
        opt.add_argument('--disable-gpu')
        
        # 随机选择User-Agent
        user_agent = random.choice(self.user_agents)
        opt.add_argument(f'--user-agent={user_agent}')
        
        driver = None
        try:
            print("正在初始化Edge浏览器...")
            driver = webdriver.Edge(
                service=EdgeService(EdgeChromiumDriverManager().install()),
                options=opt
            )
            
            # 注册WebDriver到任务管理器
            task_manager.register_webdriver(driver)
            
            # 设置页面加载超时
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            
            # 测试企查查
            print("\n===== 企查查请求测试 =====")
            try:
                print(f"正在访问: {qcc_url}")
                driver.get(qcc_url)
                
                # 等待页面加载
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                
                print(f"页面标题: {driver.title}")
                print(f"当前URL: {driver.current_url}")
                
                # 检查是否有验证码或登录页面
                if "验证" in driver.page_source or "登录" in driver.page_source:
                    print("检测到验证码或需要登录")
                
                # 尝试截图保存
                screenshot_path = "qcc_screenshot.png"
                driver.save_screenshot(screenshot_path)
                print(f"已保存页面截图到: {screenshot_path}")
                
                # 等待一下确保页面完全加载
                self.random_delay(2, 3)
                
                # 尝试查找搜索结果
                try:
                    search_results = driver.find_elements(By.CSS_SELECTOR, '.search-result-single')
                    if search_results:
                        print(f"找到搜索结果: {len(search_results)} 个")
                        # 打印第一个结果的基本信息
                        if len(search_results) > 0:
                            try:
                                first_result = search_results[0]
                                title_elem = first_result.find_element(By.CSS_SELECTOR, ".title")
                                print(f"第一个结果标题: {title_elem.text}")
                            except Exception as e:
                                print(f"提取结果信息出错: {str(e)}")
                    else:
                        print("未找到明确的搜索结果")
                except Exception as e:
                    print(f"查找搜索结果出错: {str(e)}")
                
            except Exception as e:
                print(f"请求企查查出错: {str(e)}")
            
            # 等待一段时间，避免过快发送请求
            self.random_delay(3, 5)
            
            # 测试天眼查
            print("\n===== 天眼查请求测试 =====")
            try:
                print(f"正在访问: {tianyancha_url}")
                driver.get(tianyancha_url)
                
                # 等待页面加载
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                
                print(f"页面标题: {driver.title}")
                print(f"当前URL: {driver.current_url}")
                
                # 检查是否有验证码或登录页面
                if "验证" in driver.page_source or "登录" in driver.page_source:
                    print("检测到验证码或需要登录")
                
                # 尝试截图保存
                screenshot_path = "tianyancha_screenshot.png"
                driver.save_screenshot(screenshot_path)
                print(f"已保存页面截图到: {screenshot_path}")
                
                # 等待一下确保页面完全加载
                self.random_delay(2, 3)
                
                # 尝试查找搜索结果
                try:
                    search_results = driver.find_elements(By.CSS_SELECTOR, '.search-result .result-list .content')
                    if search_results:
                        print(f"找到搜索结果: {len(search_results)} 个")
                        # 打印第一个结果的基本信息
                        if len(search_results) > 0:
                            try:
                                first_result = search_results[0]
                                title_elem = first_result.find_element(By.CSS_SELECTOR, ".header")
                                print(f"第一个结果标题: {title_elem.text}")
                            except Exception as e:
                                print(f"提取结果信息出错: {str(e)}")
                    else:
                        print("未找到明确的搜索结果")
                except Exception as e:
                    print(f"查找搜索结果出错: {str(e)}")
                
            except Exception as e:
                print(f"请求天眼查出错: {str(e)}")
                
            print("\n两个网站测试完成")
            
        except Exception as e:
            print(f"浏览器操作过程中出错: {str(e)}")
        finally:
            # 关闭浏览器
            if driver:
                # 取消注册WebDriver
                task_manager.unregister_webdriver(driver)
                try:
                    driver.quit()
                    print("浏览器已关闭")
                except Exception as e:
                    print(f"关闭浏览器出错: {str(e)}")
        
    def process_excel(self, excel_path):
        """处理Excel文件并查询著作权人信息"""
        # 识别游戏名称列
        column_name, game_names = self.identify_game_name_column(excel_path)
        
        if not game_names:
            print("没有找到有效的游戏名称数据")
            return
            
        print(f"\n找到 {len(game_names)} 个游戏名称，准备测试请求")
        
        # 测试第一个游戏的请求
        if game_names:
            test_game = game_names[0]
            print(f"\n使用第一个游戏 '{test_game}' 进行请求测试")
            self.test_selenium_browser(test_game)


# 主程序入口
if __name__ == "__main__":
    print("著作权人查询工具测试脚本")
    print("="*50)
    
    # 创建查询实例
    query = CopyrightQuery()
    
    # 手动设置Excel文件路径
    # TODO: 在这里修改为您的Excel文件路径
    excel_file_path = r"D:\git-01\著作权人测试.xlsx"  # 替换为实际的Excel文件路径
    
    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        print(f"错误: 文件不存在 - {excel_file_path}")
        excel_file_path = input("请输入正确的Excel文件路径: ").strip('"').strip("'")
        
    # 处理Excel文件
    query.process_excel(excel_file_path) 