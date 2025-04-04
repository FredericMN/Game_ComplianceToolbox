import os
import pandas as pd
import time
import re
import random
from urllib.parse import quote
import shutil  # 用于复制文件
import openpyxl
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException

# 导入WebDriver辅助类
try:
    from utils.webdriver_helper import WebDriverHelper
except ImportError:
    # 如果导入失败，尝试相对导入
    try:
        from .webdriver_helper import WebDriverHelper
    except ImportError:
        # 如果还是导入失败，创建一个简易的辅助类
        class WebDriverHelper:
            @staticmethod
            def create_driver(options=None, headless=True, progress_callback=None):
                # 默认使用原始方法创建浏览器
                from webdriver_manager.microsoft import EdgeChromiumDriverManager
                edge_options = options or webdriver.EdgeOptions()
                if headless:
                    edge_options.add_argument("--headless")
                return webdriver.Edge(
                    service=EdgeService(EdgeChromiumDriverManager().install()),
                    options=edge_options
                )
            
            @staticmethod
            def quit_driver(driver):
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass

# 导入驱动管理器
try:
    from utils.driver_manager import driver_manager
except ImportError:
    pass  # 简化导入处理，WebDriverHelper中已经有备选方案

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
        # 初始化回调函数
        self.progress_callback = None
        self.login_confirm_callback = None
        # 记录连续失败次数
        self.continuous_failures = 0
        # 最大允许的连续失败次数，超过则认为遇到了反爬机制
        self.max_failures = 3
        # 当前处理的游戏索引，用于恢复进度
        self.current_game_index = 0
        # 暂停标志
        self.paused = False
        # 暂停原因
        self.paused_reason = ""
        
    def set_progress_callback(self, callback):
        """设置进度回调函数"""
        self.progress_callback = callback
        
    def set_login_confirm_callback(self, callback):
        """设置登录确认回调函数"""
        self.login_confirm_callback = callback
    
    def update_progress(self, message, percent=None):
        """更新进度信息"""
        if self.progress_callback:
            return self.progress_callback(message, percent)
        else:
            print(message)
            return True
    
    def filter_webdriver_message(self, message, percent=None):
        """过滤WebDriver信息，只保留关键状态"""
        # 定义需要保留的关键信息关键词
        key_messages = [
            "使用已缓存的WebDriver", 
            "未找到缓存", 
            "正在下载", 
            "下载/安装完成",
            "创建浏览器实例",
            "浏览器实例创建成功",
            "缓存WebDriver配置",
            "创建WebDriver时出错"
        ]
        
        # 检查消息是否包含关键信息
        show_message = False
        for key in key_messages:
            if key in message:
                show_message = True
                break
        
        # 简化消息，删除过多细节
        if "缓存目录" in message or "缓存文件" in message:
            return None  # 忽略缓存路径信息
        
        if "检测到Edge浏览器版本" in message:
            return None  # 忽略版本检测信息
        
        # 只显示关键节点信息
        if show_message:
            # 进一步简化消息内容
            if "使用已缓存的WebDriver" in message:
                return self.update_progress("使用已缓存的WebDriver配置", percent)
            elif "未找到缓存的WebDriver" in message:
                return self.update_progress("未找到可用WebDriver缓存，需要下载配置", percent)
            elif "正在下载" in message:
                return self.update_progress("正在下载WebDriver", percent)
            elif "下载/安装完成" in message:
                return self.update_progress("WebDriver下载完成", percent)
            elif "创建浏览器实例" in message and "成功" not in message:
                return self.update_progress("正在启动浏览器", percent)
            elif "浏览器实例创建成功" in message:
                return self.update_progress("浏览器启动成功", percent)
            elif "缓存WebDriver配置" in message:
                return self.update_progress("WebDriver配置已缓存，下次将加速启动", percent)
            elif "创建WebDriver时出错" in message:
                error_detail = message.split(':', 1)[1].strip() if ':' in message else ''
                return self.update_progress(f"启动浏览器失败: {error_detail}", percent)
        
        # 默认不显示其他WebDriver消息
        return None
    
    def random_delay(self, min_sec=1.0, max_sec=3.0):
        """避免爬取过快"""
        time.sleep(random.uniform(min_sec, max_sec))
    
    def create_excel_copy_with_new_columns(self, excel_path):
        """创建Excel副本并添加新列"""
        try:
            # 创建带"-已匹配"后缀的副本文件名
            base_name, ext = os.path.splitext(excel_path)
            new_excel_path = f"{base_name}-已匹配{ext}"
            
            # 复制原文件
            shutil.copyfile(excel_path, new_excel_path)
            self.update_progress(f"已创建Excel副本: {new_excel_path}", 10)
            
            # 加载副本并添加新列
            wb = openpyxl.load_workbook(new_excel_path)
            ws = wb.active
            
            # 获取当前最后一列的索引
            last_col_idx = len(list(ws.columns))
            
            # 定义新增加的列名
            new_columns = [
                "匹配著作权人", 
                "搜索的结果数量", 
                "当前结果与游戏简称是否一致", 
                "当前结果与运营单位是否一致", 
                "是否建议人工排查"
            ]
            
            # 添加新列
            for i, col_name in enumerate(new_columns):
                ws.cell(row=1, column=last_col_idx + i + 1, value=col_name)
            
            # 保存文件
            wb.save(new_excel_path)
            self.update_progress("已在副本中添加新列", 15)
            
            return new_excel_path
        except Exception as e:
            self.update_progress(f"创建Excel副本并添加新列时出错: {str(e)}")
            return None
        
    def identify_game_name_column(self, excel_path):
        """识别Excel中的游戏名称列和运营单位列，返回游戏名称和运营单位的键值对"""
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)
            
            # 打印所有列名以便调试
            self.update_progress("Excel文件中的所有列名:")
            for i, col in enumerate(df.columns):
                self.update_progress(f"{i+1}. {col}")
            
            # 查找游戏名称列
            game_name_col = None
            for col in df.columns:
                if any(keyword in str(col) for keyword in ["游戏名称", "名称", "名字"]):
                    game_name_col = col
                    self.update_progress(f"找到游戏名称列: {game_name_col}")
                    break
            
            if not game_name_col:
                self.update_progress("未自动找到游戏名称列，请选择:")
                columns_info = []
                for i, col in enumerate(df.columns):
                    sample_values = df[col].dropna().head(3).tolist()
                    sample_text = ", ".join([str(val) for val in sample_values])
                    columns_info.append(f"{i+1}. {col} (样例值: {sample_text})")
                
                self.update_progress("\n".join(columns_info))
                
                # 调用UI请求选择列
                # (在实际实现中，这里需要一个对话框让用户选择)
                # 暂时使用第一列作为默认
                game_name_col = df.columns[0]
                self.update_progress(f"已选择游戏名称列: {game_name_col}")
            
            # 查找运营单位列
            operator_col = None
            for col in df.columns:
                if "运营单位" in str(col):
                    operator_col = col
                    self.update_progress(f"找到运营单位列: {operator_col}")
                    break
            
            if not operator_col:
                self.update_progress("未自动找到运营单位列，请选择:")
                columns_info = []
                for i, col in enumerate(df.columns):
                    sample_values = df[col].dropna().head(3).tolist()
                    sample_text = ", ".join([str(val) for val in sample_values])
                    columns_info.append(f"{i+1}. {col} (样例值: {sample_text})")
                
                self.update_progress("\n".join(columns_info))
                
                # 调用UI请求选择列
                # 暂时使用默认值
                # 如果找不到运营单位列，使用空列
                operator_col = None
                if len(df.columns) > 1:
                    operator_col = df.columns[1]
                self.update_progress(f"已选择运营单位列: {operator_col}")
            
            # 创建游戏名称和运营单位的键值对
            game_operator_dict = {}
            row_indices = {}  # 存储每个游戏对应的行索引
            
            for idx, row in df.iterrows():
                game_name = row[game_name_col]
                operator = row[operator_col] if operator_col and pd.notna(row[operator_col]) else ""
                if pd.notna(game_name):
                    game_operator_dict[game_name] = operator
                    row_indices[game_name] = idx  # 记录行索引
            
            self.update_progress(f"找到 {len(game_operator_dict)} 个游戏名称和运营单位", 20)
            
            # 返回游戏名称列名、运营单位列名、游戏-运营单位键值对和行索引映射
            return game_name_col, operator_col, game_operator_dict, row_indices
                
        except Exception as e:
            self.update_progress(f"读取Excel文件出错: {str(e)}")
            return None, None, {}, {}
    
    def extract_number_from_text(self, text):
        """从文本中提取数字"""
        if not text:
            return 0
        
        # 使用正则表达式匹配数字
        match = re.search(r'\d+', text)
        if match:
            return int(match.group())
        return 0
    
    def wait_for_page_load(self, driver, timeout=10):
        """等待页面加载完成"""
        try:
            # 等待页面完全加载
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            # 增加一个短暂的固定等待，确保动态内容加载
            time.sleep(0.5)
            return True
        except TimeoutException:
            print("页面加载超时")
            return False
    
    def click_filter_checkbox(self, driver, label_text):
        """使用WebDriverWait点击包含特定文本的筛选复选框标签"""
        try:
            # 构建XPath，查找包含特定文本的label元素
            xpath = f"//label[contains(@class, 'fcheck')][contains(., '{label_text}')]"
            self.update_progress(f"尝试查找并点击复选框: {label_text} (XPath: {xpath})")
            
            # 等待元素可见且可点击
            wait = WebDriverWait(driver, 15) # 增加等待时间
            checkbox_label = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            
            # 尝试多种点击方式
            try:
                # 方式1: 直接点击
                checkbox_label.click()
                self.update_progress(f"成功点击复选框: {label_text} (直接点击)")
                self.continuous_failures = 0  # 成功后重置失败计数
                return True
            except ElementNotInteractableException:
                self.update_progress("直接点击失败，尝试JavaScript点击")
                # 方式2: 使用JavaScript点击
                driver.execute_script("arguments[0].click();", checkbox_label)
                self.update_progress(f"成功点击复选框: {label_text} (JavaScript点击)")
                self.continuous_failures = 0  # 成功后重置失败计数
                return True
                
        except TimeoutException:
            self.update_progress(f"查找或点击复选框超时: {label_text}")
            self.save_debug_info(driver, f"checkbox_timeout_{label_text.replace(' ', '_')}")
            # 检查是否遇到了反爬机制
            return self.check_anti_crawl(f"查找或点击复选框超时: {label_text}")
        except Exception as e:
            self.update_progress(f"点击复选框时发生未知错误 ({label_text}): {str(e)}")
            self.save_debug_info(driver, f"checkbox_error_{label_text.replace(' ', '_')}")
            # 检查是否遇到了反爬机制
            return self.check_anti_crawl(f"点击复选框异常: {str(e)}")

    def get_search_result_count(self, driver):
        """获取搜索结果数量"""
        try:
            # 等待包含结果数量的span元素可见
            wait = WebDriverWait(driver, 15) # 增加等待时间
            span_xpath = "//h4[contains(., '为您找到')]/span[@class='text-danger']"
            self.update_progress(f"等待搜索结果数量元素可见 (XPath: {span_xpath})")
            span_element = wait.until(EC.visibility_of_element_located((By.XPATH, span_xpath)))
            
            # 获取文本并提取数字
            result_text = span_element.text.strip()
            self.update_progress(f"从span.text-danger获取到结果: '{result_text}'")
            if result_text: 
                self.continuous_failures = 0  # 成功后重置失败计数
                return self.extract_number_from_text(result_text)
            else:
                # 如果文本为空，尝试获取父级h4的文本
                self.update_progress("span文本为空，尝试获取父级h4文本")
                try:
                    h4_element = driver.find_element(By.XPATH, "//h4[contains(., '为您找到')]")
                    h4_text = h4_element.text.strip()
                    self.update_progress(f"从h4获取到文本: '{h4_text}'")
                    self.continuous_failures = 0  # 成功后重置失败计数
                    return self.extract_number_from_text(h4_text)
                except Exception as e_h4:
                    self.update_progress(f"获取h4文本失败: {str(e_h4)}")
                    # 检查是否遇到了反爬机制
                    if not self.check_anti_crawl(f"获取h4文本失败: {str(e_h4)}"):
                        return 0
                    return 0 # 返回0表示未找到或提取失败
        except TimeoutException:
            self.update_progress("等待搜索结果数量元素超时")
            # 检查页面是否提示"未找到相关结果"
            try:
                no_result_xpath = "//div[contains(@class, 'nodata')] | //div[contains(text(), '未找到相关结果')] | //p[contains(text(), '未找到相关结果')]"
                driver.find_element(By.XPATH, no_result_xpath)
                self.update_progress("页面提示未找到相关结果，返回 0")
                self.continuous_failures = 0  # 成功识别到"无结果"也是正常情况，重置失败计数
                return 0
            except NoSuchElementException:
                self.update_progress("未找到'无结果'提示，可能页面结构有变或加载失败")
                self.save_debug_info(driver, "result_count_timeout")
                # 检查是否遇到了反爬机制
                if not self.check_anti_crawl("未找到搜索结果元素，也未找到'无结果'提示"):
                    return 0
                return 0 # 超时且无明确提示，返回0
        except Exception as e:
            self.update_progress(f"获取搜索结果数量时发生未知错误: {str(e)}")
            self.save_debug_info(driver, "result_count_error")
            # 检查是否遇到了反爬机制
            if not self.check_anti_crawl(f"获取搜索结果异常: {str(e)}"):
                return 0
            return 0

    def extract_and_match_results(self, driver, game_name, operator):
        """解析搜索结果列表，提取信息并根据规则进行匹配"""
        results = [] # 存储提取结果: [(简称, 著作权人), ...]
        matched_copyright_owner = ""
        is_game_name_match = "否"
        is_operator_match = "否"
        recommend_manual_check = "是" # 默认需要人工排查
        
        try:
            # 定位结果列表容器
            results_container_xpath = "//div[contains(@class, 'bigsearch-list')]//table[contains(@class, 'ntable')]"
            # 获取结果表格
            try:
                results_table = driver.find_element(By.XPATH, results_container_xpath)
                self.update_progress(f"成功定位结果表格: {results_container_xpath}")
            except NoSuchElementException:
                self.update_progress(f"未找到结果表格: {results_container_xpath}")
                self.save_debug_info(driver, "no_results_table")
                return matched_copyright_owner, is_game_name_match, is_operator_match, recommend_manual_check
            
            # 获取所有结果行 (tr)
            result_rows = results_table.find_elements(By.TAG_NAME, "tr")
            self.update_progress(f"找到 {len(result_rows)} 条结果行")
            
            if not result_rows:
                self.update_progress("未在表格中找到结果行(tr)")
                return matched_copyright_owner, is_game_name_match, is_operator_match, recommend_manual_check

            # 提取每行的信息
            for row in result_rows:
                short_name = "-"
                owner = ""
                try:
                    # --- 提取软件简称 --- (修改后：提取span.val的完整文本)
                    self.update_progress("\n开始提取软件简称...")
                    short_name = "-" # 默认值
                    try:
                        # 尝试主要XPath查找包含软件简称的span.val
                        # XPath解释：查找任意包含 '软件简称' 文本的 span，然后找它后面紧跟着的 class 包含 'val' 的 span
                        val_span_xpath = ".//span[contains(text(), '软件简称')]/following-sibling::span[contains(@class, 'val')]"
                        val_elements = row.find_elements(By.XPATH, val_span_xpath)

                        if val_elements:
                            # 获取第一个找到的span.val的完整文本内容
                            # .text 属性会获取元素及其所有子元素的可见文本
                            short_name = val_elements[0].text.strip()
                            self.update_progress(f"通过主要XPath找到span.val，提取到完整简称: '{short_name}'")
                        else:
                            # 如果主要XPath找不到，尝试备选方法：查找包含"软件简称"的span.f，再找其兄弟span.val
                            self.update_progress("主要XPath未找到span.val，尝试备选方法...")
                            try:
                                # 查找所有 class='f' 的 span
                                label_spans = row.find_elements(By.XPATH, ".//span[@class='f']")
                                found_backup = False
                                for label_span in label_spans:
                                    # 检查哪个标签包含"软件简称"
                                    if "软件简称" in label_span.text:
                                        # 找到标签后，尝试寻找对应的数值 span (span.val)
                                        try:
                                            # 尝试1：作为直接的兄弟节点
                                            # ./following-sibling:: 表示查找当前节点之后的兄弟节点
                                            val_span = label_span.find_element(By.XPATH, "./following-sibling::span[contains(@class, 'val')]")
                                            short_name = val_span.text.strip() # 获取完整文本
                                            self.update_progress(f"通过备选方法(直接兄弟)找到span.val，提取到完整简称: '{short_name}'")
                                            found_backup = True
                                            break # 找到就跳出循环
                                        except NoSuchElementException:
                                            # 尝试2：作为父级 div 的兄弟 span.val (处理可能的嵌套结构)
                                            try:
                                                 # ./parent::div 找到直接的父级 div 元素
                                                 parent_div = label_span.find_element(By.XPATH, "./parent::div")
                                                 # 再从父级 div 查找兄弟 span.val
                                                 val_span = parent_div.find_element(By.XPATH, "./following-sibling::span[contains(@class, 'val')]")
                                                 short_name = val_span.text.strip() # 获取完整文本
                                                 self.update_progress(f"通过备选方法(父级兄弟)找到span.val，提取到完整简称: '{short_name}'")
                                                 found_backup = True
                                                 break # 找到就跳出循环
                                            except NoSuchElementException:
                                                # 如果两种结构都没找到，记录一下信息，继续检查下一个可能的标签span
                                                self.update_progress(f"备选方法在标签 '{label_span.text[:20]}...' 处未找到对应的span.val")
                                                continue

                                if not found_backup:
                                     self.update_progress("备选方法也未能找到简称对应的span.val")

                            except Exception as e_backup:
                                self.update_progress(f"执行备选提取方法时出错: {str(e_backup)}")
                                # 出错则保持 short_name 为 "-"

                        # 最终清理，如果提取结果为空字符串，也设为"-"
                        if not short_name or short_name.strip() == "":
                            short_name = "-"
                            self.update_progress("提取到的简称为空，重置为 '-'")

                    except Exception as e_sn:
                        # 捕获整个提取简称过程中的任何异常
                        self.update_progress(f"提取软件简称时发生异常: {str(e_sn)}")
                        short_name = "-" # 确保异常时为默认值 "-"

                    self.update_progress(f"最终确定的用于匹配的简称: '{short_name}'")

                    # --- 提取著作权人 --- (逻辑不变)
                    try:
                        # 1. 找到包含"著作权人："的父级 span.f
                        # 注意：著作权人可能嵌套在 app-over-hidden-text div 中
                        label_span_owner_candidates = row.find_elements(By.XPATH, ".//span[@class='f' and starts-with(normalize-space(.), '著作权人')] | .//span[starts-with(normalize-space(.), '著作权人')]//span[@class='f']")
                        if not label_span_owner_candidates:
                             raise NoSuchElementException("未找到著作权人标签 span.f")
                        
                        label_span_owner = None
                        for candidate in label_span_owner_candidates:
                            try:
                                parent_div = candidate.find_element(By.XPATH, "./ancestor::div[contains(@class, 'rline')] | ./ancestor::div[@class='f']/ancestor::div[contains(@class, 'rline')]")
                                if parent_div:
                                    label_span_owner = candidate
                                    break
                            except NoSuchElementException:
                                continue
                        if not label_span_owner:
                             label_span_owner = label_span_owner_candidates[0]
                             
                        val_span_owner = None
                        try:
                            val_span_owner = label_span_owner.find_element(By.XPATH, "following-sibling::span[contains(@class, 'val')]")
                        except NoSuchElementException:
                             try:
                                 wrapper_div = label_span_owner.find_element(By.XPATH, "./parent::div[contains(@class,'f')]")
                                 val_span_owner = wrapper_div.find_element(By.XPATH, "following-sibling::span[contains(@class, 'val')]")
                             except NoSuchElementException:
                                  val_span_owner = row.find_element(By.XPATH, ".//span[starts-with(normalize-space(.), '著作权人')]/following-sibling::span[contains(@class, 'val')]")
                        
                        if not val_span_owner:
                             raise NoSuchElementException("未找到著作权人值 span.val")

                        owner_link = val_span_owner.find_element(By.TAG_NAME, 'a')
                        owner = owner_link.text.strip()
                        self.update_progress(f"提取到著作权人: '{owner}'")
                    except NoSuchElementException:
                        self.update_progress("未找到著作权人信息")
                        owner = ""
                    
                    # 确保提取到了有效信息再添加
                    if owner or short_name != "-": 
                        results.append((short_name if short_name else "-", owner))
                        self.update_progress(f"提取到结果: 简称='{short_name if short_name else "-"}', 著作权人='{owner}'")
                    else:
                        self.update_progress("跳过一条结果，未提取到有效简称或著作权人")

                except Exception as e_row:
                    self.update_progress(f"处理结果行时出错: {str(e_row)}")
                    self.save_debug_info(driver, f"row_processing_error")
            
            self.update_progress(f"共提取到 {len(results)} 条有效结果进行匹配")
            if not results:
                return matched_copyright_owner, is_game_name_match, is_operator_match, recommend_manual_check

            # 匹配逻辑
            operator_matched_result = None
            game_name_matches = []
            
            # 1. 优先匹配运营单位 (著作权人)
            if operator: 
                for sn, own in results:
                    if own == operator:
                        self.update_progress(f"运营单位匹配成功: 结果著作权人 '{own}' == 表格运营单位 '{operator}'")
                        operator_matched_result = (sn, own)
                        break 
            
            if operator_matched_result:
                matched_copyright_owner = operator_matched_result[1]
                is_operator_match = "是"
                if operator_matched_result[0] == game_name or (operator_matched_result[0] == '-' and not game_name) :
                     is_game_name_match = "是"
                else:
                     is_game_name_match = "否"
                     self.update_progress(f"运营单位匹配成功，但简称不匹配: 结果简称 '{operator_matched_result[0]}' != 表格游戏名 '{game_name}'")
                recommend_manual_check = "否" 
            else:
                # 2. 如果运营单位未匹配，则匹配游戏名称 (软件简称)
                self.update_progress("运营单位未匹配，开始匹配游戏简称...")
                is_operator_match = "否"
                for sn, own in results:
                    if sn == game_name:
                        game_name_matches.append((sn, own))
                        self.update_progress(f"游戏简称匹配成功: '{sn}'")
                
                if len(game_name_matches) == 1:
                    self.update_progress("游戏简称唯一匹配")
                    matched_copyright_owner = game_name_matches[0][1]
                    is_game_name_match = "是"
                    recommend_manual_check = "否"
                elif len(game_name_matches) > 1:
                    self.update_progress("游戏简称存在多个匹配，标记需人工排查")
                    matched_copyright_owner = game_name_matches[0][1] 
                    is_game_name_match = "是"
                    recommend_manual_check = "是"
                else:
                    self.update_progress("游戏简称也未匹配")
                    is_game_name_match = "否"
                    recommend_manual_check = "是" 

        except NoSuchElementException:
            self.update_progress("未找到结果列表容器或表格")
            recommend_manual_check = "是"
        except Exception as e:
            self.update_progress(f"解析和匹配过程中发生错误: {str(e)}")
            self.save_debug_info(driver, "extract_match_error")
            recommend_manual_check = "是"
            
        self.update_progress(f"匹配结果: 著作权人='{matched_copyright_owner}', 简称匹配='{is_game_name_match}', 运营单位匹配='{is_operator_match}', 人工排查='{recommend_manual_check}'")
        return matched_copyright_owner, is_game_name_match, is_operator_match, recommend_manual_check

    def save_debug_info(self, driver, prefix="debug"):
        """保存页面源码和截图用于调试"""
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        source_file = f"{prefix}_pagesource_{timestamp}.html"
        screenshot_file = f"{prefix}_screenshot_{timestamp}.png"
        # 获取网页源码以便调试
        try:
            source = driver.page_source
            with open(source_file, "w", encoding="utf-8") as f:
                f.write(source)
            print(f"已保存页面源码到 {source_file}")
        except Exception as e:
            print(f"保存页面源码失败: {str(e)}")
        
        # 截图保存
        try:
            driver.save_screenshot(screenshot_file)
            print(f"已保存页面截图到 {screenshot_file}")
        except Exception as e:
            print(f"保存截图失败: {str(e)}")

    def query_copyright(self, excel_path, progress_callback=None, start_from_index=0):
        """查询著作权人信息并填充Excel"""
        # 如果传入了回调，临时替换全局回调
        old_callback = self.progress_callback
        if progress_callback:
            self.progress_callback = progress_callback
            
        # 创建Excel副本并添加新列
        new_excel_path = self.create_excel_copy_with_new_columns(excel_path)
        if not new_excel_path:
            self.update_progress("创建Excel副本失败，无法继续")
            if old_callback:
                self.progress_callback = old_callback
            return None
            
        # 识别游戏名称列和运营单位列，获取游戏-运营单位键值对和行索引映射
        game_name_col, operator_col, game_operator_dict, row_indices = self.identify_game_name_column(excel_path)
        if not game_operator_dict:
            self.update_progress("没有找到有效的游戏名称数据")
            if old_callback:
                self.progress_callback = old_callback
            return None
        
        # 加载副本Excel准备写入
        df_result = pd.read_excel(new_excel_path)
        
        # --- 处理Pandas类型警告 --- 
        # 显式设置新添加的文本列的数据类型为 object
        text_columns = [
            "匹配著作权人", 
            "当前结果与游戏简称是否一致", 
            "当前结果与运营单位是否一致", 
            "是否建议人工排查"
        ]
        for col in text_columns:
            if col in df_result.columns:
                 # 确保列存在，并将其类型设置为object，填充NA值为空字符串以保持一致性
                 df_result[col] = df_result[col].astype(object).fillna('') 
            else:
                 # 如果列不存在（理论上不应发生），创建一个object类型的空列
                 df_result[col] = pd.Series(dtype='object')
                 
        # "搜索的结果数量" 列可以尝试设为整数，但需处理可能的NA
        count_col = "搜索的结果数量"
        if count_col in df_result.columns:
            # 尝试转换为整数，无法转换的（如空值）保持原样或设为特定值（如-1或NaN）
            # 先填充NaN为0或其他标记，再尝试转换类型
            df_result[count_col] = pd.to_numeric(df_result[count_col], errors='coerce').fillna(0).astype(int)
        else:
             df_result[count_col] = pd.Series(dtype='int')
        # --- 类型处理结束 ---

        # 创建Edge浏览器选项
        opt = webdriver.EdgeOptions()
        opt.add_argument('--disable-extensions')
        opt.add_argument('--no-sandbox')
        opt.add_argument('--disable-dev-shm-usage')
        opt.add_argument('--disable-gpu')
        # 添加禁用自动化控制特征的选项
        opt.add_experimental_option("excludeSwitches", ["enable-automation"])
        opt.add_experimental_option('useAutomationExtension', False)
        user_agent = random.choice(self.user_agents)
        opt.add_argument(f'--user-agent={user_agent}')
        
        driver = None
        try:
            # 使用WebDriverHelper创建浏览器实例
            self.update_progress("正在准备启动浏览器...", 20)
            
            # 先尝试清理可能存在的残留进程
            WebDriverHelper.kill_msedgedriver()
            self.update_progress("已清理可能存在的WebDriver进程", 21)
            
            # 创建浏览器实例
            driver = WebDriverHelper.create_driver(
                options=opt, 
                headless=False,  # 不使用无头模式，因为需要用户登录
                progress_callback=self.filter_webdriver_message
            )
            
            if not driver:
                self.update_progress("创建浏览器实例失败，请检查环境或网络连接后重试", 0)
                if old_callback:
                    self.progress_callback = old_callback
                return None
                
            self.update_progress("浏览器实例创建成功！", 30)
            
            # 注册WebDriver到任务管理器
            task_manager.register_webdriver(driver)
            
            # 设置页面加载超时
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            
            # 使用第一个游戏作为示例，访问企查查著作权查询页面
            first_game = list(game_operator_dict.keys())[0]
            encoded_name = quote(first_game)
            qcc_url = f"https://www.qcc.com/web_searchCopyright?searchKey={encoded_name}&type=2"
            
            self.update_progress(f"正在访问企查查著作权查询页面: {qcc_url}", 30)
            driver.get(qcc_url)
            
            # 等待页面加载
            self.wait_for_page_load(driver)
            
            # 提示用户登录
            self.update_progress("\n检测到需要登录企查查账号。请登录后继续...", 35)
            
            # 调用登录确认回调
            if self.login_confirm_callback:
                login_confirmed = self.login_confirm_callback()
                if not login_confirmed:
                    self.update_progress("用户取消了登录，终止查询", 0)
                    if old_callback:
                        self.progress_callback = old_callback
                    return None
            else:
                # 如果没有回调，使用旧的控制台交互方式
                self.update_progress("1. 请在浏览器中登录您的企查查账号")
                self.update_progress("2. 登录完成后，请在此终端输入数字1并按回车键继续")
                
                # 等待用户确认已登录
                while True:
                    user_input = input("请输入数字1表示已完成登录: ")
                    if user_input.strip() == "1":
                        self.update_progress("已确认登录完成，继续处理...")
                        break
                    else:
                        self.update_progress("输入无效，请输入数字1")
            
            # 获取游戏列表，转换为列表以便按索引访问
            game_list = list(game_operator_dict.items())
            total_games = len(game_list)
            
            # 从指定索引开始查询
            if start_from_index > 0:
                self.update_progress(f"从第 {start_from_index+1} 个游戏继续查询")
            
            # 查询每个游戏的著作权信息
            for i in range(start_from_index, total_games):
                # 记录当前处理索引，用于可能的恢复
                self.current_game_index = i
                
                # 获取当前游戏信息
                game_name, operator = game_list[i]
                
                try:
                    self.update_progress(f"\n[{i+1}/{total_games}] 查询游戏: {game_name}", 
                                         35 + int(60 * (i / total_games)))
                    self.update_progress(f"游戏名称: {game_name}, 运营单位: {operator if operator else '未提供'}")
                    
                    encoded_name = quote(game_name)
                    search_url = f"https://www.qcc.com/web_searchCopyright?searchKey={encoded_name}&type=2"
                    
                    # 访问搜索页面
                    self.update_progress(f"访问搜索页面: {search_url}")
                    driver.get(search_url)
                    
                    # 等待页面加载完成
                    if not self.wait_for_page_load(driver, timeout=15):
                        self.update_progress("页面加载失败，尝试刷新...")
                        driver.refresh()
                        self.random_delay(1, 2)
                        if not self.wait_for_page_load(driver, timeout=15):
                            self.update_progress("刷新后仍然无法加载页面，跳过此游戏")
                            # 记录错误或标记
                            row_idx = row_indices.get(game_name)
                            if row_idx is not None:
                                df_result.loc[row_idx, "是否建议人工排查"] = "是 (页面加载失败)"
                            continue
                    
                    # 在每次搜索后应用筛选条件
                    self.update_progress("应用筛选条件...")
                    clicked_filter1 = self.click_filter_checkbox(driver, '1年内')
                    
                    # 如果遇到反爬机制并且暂停了，需要等待用户登录后继续
                    if self.paused:
                        # 保存当前进度
                        df_result.to_excel(new_excel_path, index=False)
                        self.update_progress(f"已保存当前进度到: {new_excel_path}")
                        
                        # 重置当前页面并重新尝试
                        driver.get(search_url)
                        if not self.wait_for_page_load(driver, timeout=15):
                            continue
                            
                        # 重新尝试点击筛选条件
                        clicked_filter1 = self.click_filter_checkbox(driver, '1年内')
                        if not clicked_filter1:
                            continue
                        # 如果确认后继续，需要重置暂停标志
                        self.paused = False # <-- 确保重置暂停状态
                    
                    self.random_delay(0.5, 1) # 减少延迟
                    clicked_filter2 = self.click_filter_checkbox(driver, '1-3年')
                    
                    # 同样检查第二个筛选条件是否触发了反爬机制
                    if self.paused:
                        # 保存当前进度
                        df_result.to_excel(new_excel_path, index=False)
                        self.update_progress(f"已保存当前进度到: {new_excel_path}")
                        
                        # 重置当前页面并重新尝试
                        driver.get(search_url)
                        if not self.wait_for_page_load(driver, timeout=15):
                            continue
                            
                        # 重新尝试点击前两个筛选条件
                        clicked_filter1 = self.click_filter_checkbox(driver, '1年内')
                        self.random_delay(0.5, 1)
                        clicked_filter2 = self.click_filter_checkbox(driver, '1-3年')
                        if not clicked_filter1 or not clicked_filter2:
                            continue
                        # 如果确认后继续，需要重置暂停标志
                        self.paused = False # <-- 确保重置暂停状态

                    # --- 新增点击 '3-5年' ---
                    self.random_delay(0.5, 1) # 增加点击间隔
                    clicked_filter3 = self.click_filter_checkbox(driver, '3-5年')

                    # 检查第三个筛选条件是否触发了反爬机制
                    if self.paused:
                        # 保存当前进度
                        df_result.to_excel(new_excel_path, index=False)
                        self.update_progress(f"已保存当前进度到: {new_excel_path}")

                        # 重置当前页面并重新尝试
                        driver.get(search_url)
                        if not self.wait_for_page_load(driver, timeout=15):
                            continue

                        # 重新尝试点击所有三个筛选条件
                        clicked_filter1 = self.click_filter_checkbox(driver, '1年内')
                        self.random_delay(0.5, 1)
                        clicked_filter2 = self.click_filter_checkbox(driver, '1-3年')
                        self.random_delay(0.5, 1)
                        clicked_filter3 = self.click_filter_checkbox(driver, '3-5年')
                        if not clicked_filter1 or not clicked_filter2 or not clicked_filter3:
                            continue
                        # 如果确认后继续，需要重置暂停标志
                        self.paused = False # <-- 确保重置暂停状态
                    # --- 新增结束 ---

                    self.random_delay(1, 2) # 点击后等待页面刷新

                    # 更新成功检查逻辑，包含第三个筛选器
                    if not clicked_filter1 or not clicked_filter2 or not clicked_filter3:
                        self.update_progress("警告：未能成功点击所有筛选条件 ('1年内', '1-3年', '3-5年')，结果可能不准确。")
                    else:
                        self.update_progress("已成功应用筛选条件 ('1年内', '1-3年', '3-5年')。")

                    # 获取搜索结果数量 (应用筛选后)
                    result_count = self.get_search_result_count(driver)
                    
                    # 检查是否因为反爬机制暂停了 (获取结果数量后)
                    if self.paused:
                        # 保存当前进度
                        df_result.to_excel(new_excel_path, index=False)
                        self.update_progress(f"已保存当前进度到: {new_excel_path}")
                        
                        # 重置当前页面并重新尝试整个流程
                        driver.get(search_url)
                        if not self.wait_for_page_load(driver, timeout=15):
                            continue
                            
                        # 重新尝试整个筛选和获取结果数量过程
                        clicked_filter1 = self.click_filter_checkbox(driver, '1年内')
                        self.random_delay(0.5, 1)
                        clicked_filter2 = self.click_filter_checkbox(driver, '1-3年')
                        self.random_delay(0.5, 1)
                        clicked_filter3 = self.click_filter_checkbox(driver, '3-5年') # 别忘了这里也要重试
                        self.random_delay(1, 2)
                        result_count = self.get_search_result_count(driver)
                        
                        # 如果再次暂停，跳过当前游戏
                        if self.paused: 
                            continue
                        # 如果确认后继续，需要重置暂停标志
                        self.paused = False # <-- 确保重置暂停状态

                    self.update_progress(f"筛选后搜索结果数量: {result_count}")
                    
                    matched_copyright_owner = ""
                    is_game_name_match = "否"
                    is_operator_match = "否"
                    recommend_manual_check = "是"
                    
                    row_idx = row_indices.get(game_name)
                    
                    if row_idx is not None:
                        df_result.loc[row_idx, "搜索的结果数量"] = result_count
                        self.update_progress(f"已更新行 {row_idx} 的搜索结果数量: {result_count}")
                        
                        if result_count > 0:
                            self.update_progress("开始解析和匹配搜索结果...")
                            current_operator = game_operator_dict.get(game_name, "") 
                            matched_copyright_owner, is_game_name_match, is_operator_match, recommend_manual_check = \
                                self.extract_and_match_results(driver, game_name, current_operator)
                            
                            # 检查是否因为反爬机制暂停了
                            if self.paused:
                                # 保存当前进度
                                df_result.to_excel(new_excel_path, index=False)
                                self.update_progress(f"已保存当前进度到: {new_excel_path}")
                                
                                # 重置当前页面并重新尝试整个流程
                                continue # 跳到下一个游戏的处理
                            
                            df_result.loc[row_idx, "匹配著作权人"] = matched_copyright_owner
                            df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = is_game_name_match
                            df_result.loc[row_idx, "当前结果与运营单位是否一致"] = is_operator_match
                            df_result.loc[row_idx, "是否建议人工排查"] = recommend_manual_check
                            self.update_progress(f"已更新行 {row_idx} 的匹配结果")
                            
                            # 添加详细的匹配结果输出
                            self.update_progress(f"匹配结果: 著作权人='{matched_copyright_owner}', " +
                                               f"简称匹配='{is_game_name_match}', " + 
                                               f"运营单位匹配='{is_operator_match}', " + 
                                               f"人工排查='{recommend_manual_check}'")
                        else:
                            self.update_progress("搜索结果为0，标记建议人工排查")
                            df_result.loc[row_idx, "是否建议人工排查"] = "是"
                            df_result.loc[row_idx, "匹配著作权人"] = ""
                            df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = "否"
                            df_result.loc[row_idx, "当前结果与运营单位是否一致"] = "否"
                            
                            # 添加零结果的匹配输出
                            self.update_progress(f"匹配结果: 著作权人='', 简称匹配='否', 运营单位匹配='否', 人工排查='是'")

                    else:
                        self.update_progress(f"警告：未找到游戏 '{game_name}' 在原始Excel中的行索引")

                    # 保存中间结果
                    if (i + 1) % 5 == 0 or i == total_games - 1:
                        try:
                           df_result.to_excel(new_excel_path, index=False)
                           self.update_progress(f"已保存进度到: {new_excel_path}")
                        except Exception as save_e:
                           self.update_progress(f"保存Excel时出错: {str(save_e)}")
                
                except Exception as e:
                    self.update_progress(f"处理游戏 {game_name} 时出错: {str(e)}")
                    self.save_debug_info(driver, f"game_error_{game_name.replace(' ', '_')}")
                    row_idx = row_indices.get(game_name)
                    if row_idx is not None:
                        df_result.loc[row_idx, "是否建议人工排查"] = f"是 (处理异常: {str(e)[:50]}...)"
                    try:
                        df_result.to_excel(new_excel_path, index=False)
                    except Exception as save_e:
                           self.update_progress(f"保存Excel时出错: {str(save_e)}")
                
                # 在处理完一个游戏后增加较长的随机延迟（3-5秒）
                self.random_delay(3, 5)
                self.update_progress("增加随机延迟，减轻反爬机制作用...")
            
            # 保存最终结果
            try:
               df_result.to_excel(new_excel_path, index=False)
               self.update_progress(f"\n著作权人查询完成，结果已保存到: {new_excel_path}", 95)
            except Exception as save_e:
               self.update_progress(f"最终保存Excel时出错: {str(save_e)}")
            
        except Exception as e:
            self.update_progress(f"浏览器操作过程中出错: {str(e)}")
            # 发生异常时尝试保存当前进度
            try:
                if 'df_result' in locals() and new_excel_path:
                    df_result.to_excel(new_excel_path, index=False)
                    self.update_progress(f"已保存当前进度到: {new_excel_path}")
            except Exception as save_error:
                self.update_progress(f"保存进度时出错: {str(save_error)}")
        finally:
            # 关闭浏览器
            if driver:
                # 取消注册WebDriver
                task_manager.unregister_webdriver(driver)
                try:
                    # 使用WebDriverHelper安全关闭浏览器
                    WebDriverHelper.quit_driver(driver)
                    self.update_progress("浏览器已关闭", 100)
                except Exception as e:
                    self.update_progress(f"关闭浏览器出错: {str(e)}")
            
            # 恢复原来的回调
            if old_callback:
                self.progress_callback = old_callback
        
        return new_excel_path
    
    def process_excel(self, excel_path, start_from_index=0):
        """处理Excel文件并查询著作权人信息"""
        # 直接处理整个Excel文件，可以指定开始的索引以恢复进度
        return self.query_copyright(excel_path, start_from_index=start_from_index)

    def check_anti_crawl(self, exception_msg=None):
        """检查是否遇到了反爬机制"""
        self.continuous_failures += 1
        if self.continuous_failures >= self.max_failures:
            self.continuous_failures = 0  # 重置失败计数
            self.paused = True  # 标记为暂停状态
            self.paused_reason = exception_msg if exception_msg else "连续失败次数超过阈值"
            
            self.update_progress(f"检测到可能遇到网站反爬机制，需要重新登录或等待一段时间解除限制...")
            
            # 调用登录确认回调
            if self.login_confirm_callback:
                login_confirmed = self.login_confirm_callback()
                if login_confirmed:
                    self.update_progress("用户已确认重新登录，继续查询...")
                    self.paused = False  # 解除暂停状态
                    return True  # 继续处理
                else:
                    self.update_progress("用户取消了登录，终止查询", 0)
                    return False  # 终止处理
            else:
                # 如果没有回调，使用旧的控制台交互方式
                self.update_progress("1. 请在浏览器中重新登录")
                self.update_progress("2. 登录完成后，请在此终端输入数字1并按回车键继续")
                
                # 等待用户确认已登录
                while True:
                    user_input = input("请输入数字1表示已完成登录: ")
                    if user_input.strip() == "1":
                        self.update_progress("已确认登录完成，继续处理...")
                        self.paused = False  # 解除暂停状态
                        return True  # 继续处理
                    else:
                        self.update_progress("输入无效，请输入数字1")
        return True  # 连续失败次数未超过阈值，继续处理


# 主程序入口
if __name__ == "__main__":
    print("著作权人查询工具测试脚本")
    print("="*50)
    
    # 创建查询实例
    query = CopyrightQuery()
    
    # 手动设置Excel文件路径
    # 可以通过命令行参数提供Excel文件路径
    if len(sys.argv) > 1:
        excel_file_path = sys.argv[1]
    else:
        # 默认文件路径
        excel_file_path = r"D:\git-01\著作权人测试.xlsx"  # 替换为实际的Excel文件路径
    
    # 检查文件是否存在
    if not os.path.exists(excel_file_path):
        print(f"错误: 文件不存在 - {excel_file_path}")
        excel_file_path = input("请输入正确的Excel文件路径: ").strip('"').strip("'")
        
    # 处理Excel文件
    query.process_excel(excel_file_path) 