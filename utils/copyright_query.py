import os
import pandas as pd
import time
import re
import random
from urllib.parse import quote
import shutil  # 用于复制文件
import openpyxl
import sys
from selenium.webdriver.common.keys import Keys

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
            # encoded_name = quote(first_game) # 第一次访问时再编码
            # qcc_url = f"https://www.qcc.com/web_searchCopyright?searchKey={encoded_name}&type=2" # 移到循环内
            
            # self.update_progress(f"正在访问企查查著作权查询页面: {qcc_url}", 30)
            # driver.get(qcc_url) # 移到循环内处理首次访问
            
            # # 等待页面加载 # 移到循环内处理首次访问
            # self.wait_for_page_load(driver)
            
            # 提示用户登录
            self.update_progress("\n检测到可能需要登录企查查账号。", 35)
            self.update_progress("如果页面长时间未加载或显示登录界面，请手动登录。", 35)
            self.update_progress("脚本将在登录后或可直接访问时继续。", 35)
            
            # --- 调整：首次访问放在循环内处理，并增加登录确认 --- 
            is_first_game = True # 标记是否是处理第一个游戏
            qcc_base_url = "https://www.qcc.com/web_searchCopyright?type=2&searchKey=" # 基础URL，后面拼接游戏名
            is_state_filtered = False # <--- 新增状态变量：跟踪页面是否预期为已筛选状态

            # --- 登录确认逻辑提前 --- 
            # 访问一个基础页面（如搜索首页）来触发登录检查
            try:
                driver.get("https://www.qcc.com/") # 访问首页，通常需要登录才能搜索
                self.wait_for_page_load(driver, 10)
                # 简单检查是否跳转到了登录页 (URL可能包含passport)
                if "passport.qcc.com" in driver.current_url:
                    login_needed = True
                else:
                    # 尝试定位一个登录后才有的元素，例如用户头像或用户名
                    try:
                        # 此处需要替换为实际登录后可见元素的定位器
                        # 例如： driver.find_element(By.CSS_SELECTOR, ".user-avatar") 
                        # 暂时假设能找到某个元素表示已登录
                        driver.find_element(By.XPATH, "//a[contains(@href, '/user_center')]") # 示例：用户中心链接
                        login_needed = False
                        self.update_progress("检测到已登录状态。")
                    except NoSuchElementException:
                        login_needed = True 
            except Exception as e_nav:
                 self.update_progress(f"访问企查查首页检查登录状态时出错: {e_nav}，假设需要登录。")
                 login_needed = True

            if login_needed:
                self.update_progress("请在弹出的浏览器窗口中登录您的企查查账号。", 36)
                if self.login_confirm_callback:
                    login_confirmed = self.login_confirm_callback() # 调用回调等待用户确认
                    if not login_confirmed:
                        self.update_progress("用户取消了登录，终止查询", 0)
                        if old_callback: self.progress_callback = old_callback
                        return None
                    self.update_progress("登录已确认，继续查询。", 37)
                else:
                    # 控制台确认方式
                    input("请在浏览器中完成登录后，按回车键继续...")
                    self.update_progress("继续查询...", 37)
            # --- 登录确认结束 --- 
            
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
                    
                    # --- 步骤1: 执行搜索 (区分首次和后续) ---
                    if is_first_game:
                        # 首次访问，使用URL导航
                        encoded_name = quote(game_name)
                        search_url = qcc_base_url + encoded_name
                        self.update_progress(f"首次访问，使用URL: {search_url}")
                        try:
                             driver.get(search_url)
                             is_first_game = False # 更新标记
                        except Exception as e_first_nav:
                             self.update_progress(f"首次通过URL访问失败: {e_first_nav}, 跳过此游戏")
                             row_idx = row_indices.get(game_name)
                             if row_idx is not None: df_result.loc[row_idx, "是否建议人工排查"] = "是 (首次URL访问失败)"
                             continue # 进行下一个游戏
                    else:
                        # 后续访问，使用页面内搜索框
                        self.update_progress(f"后续访问，使用页面内搜索框查询: '{game_name}'")
                        try:
                            search_box_id = "copyrightSearchKey"
                            wait = WebDriverWait(driver, 15)
                            search_input = wait.until(EC.presence_of_element_located((By.ID, search_box_id)))
                            self.update_progress("找到搜索框，准备输入...")
                            # search_input.clear() # 改用更可靠的清空方式
                            search_input.send_keys(Keys.CONTROL + "a") # 全选
                            search_input.send_keys(Keys.DELETE)     # 删除
                            self.update_progress("已清空搜索框内容")
                            search_input.send_keys(game_name)
                            search_input.send_keys(Keys.RETURN) # 模拟回车
                            self.update_progress(f"已在搜索框输入 '{game_name}' 并回车")
                        except (TimeoutException, NoSuchElementException) as e_searchbox:
                            self.update_progress(f"错误：无法找到或操作搜索框: {e_searchbox}")
                            self.save_debug_info(driver, f"searchbox_error_{game_name[:10]}")
                            self.update_progress("尝试回退到URL导航...")
                            try:
                                encoded_name = quote(game_name)
                                search_url = qcc_base_url + encoded_name
                                driver.get(search_url)
                                self.update_progress(f"URL回退导航成功: {search_url}")
                            except Exception as e_fallback_nav:
                                 self.update_progress(f"URL回退失败: {e_fallback_nav}, 跳过此游戏")
                                 row_idx = row_indices.get(game_name)
                                 if row_idx is not None: df_result.loc[row_idx, "是否建议人工排查"] = "是 (搜索框和URL导航均失败)"
                                 continue # 进行下一个游戏
                                 
                    # 等待页面加载完成 (无论哪种搜索方式后都需要)
                    if not self.wait_for_page_load(driver, timeout=20): # 增加超时时间
                        self.update_progress("页面加载失败，尝试刷新...")
                        driver.refresh()
                        self.random_delay(2, 3)
                        if not self.wait_for_page_load(driver, timeout=20):
                            self.update_progress("刷新后仍然无法加载页面，跳过此游戏")
                            row_idx = row_indices.get(game_name)
                            if row_idx is not None: df_result.loc[row_idx, "是否建议人工排查"] = "是 (页面加载失败)"
                            continue # 进行下一个游戏
                    # --- 搜索执行结束 ---
                            
                    # --- 步骤2: 获取当前状态结果 (搜索后，显式操作筛选前) ---
                    self.update_progress(f"\n获取当前页面状态结果 (预期状态: {'已筛选' if is_state_filtered else '未筛选'})...")
                    current_state_count = 0
                    current_state_owner = ""
                    current_state_game_match = "否"
                    current_state_operator_match = "否"
                    current_state_manual_check = "是"
                    
                    try:
                        current_state_count = self.get_search_result_count(driver)
                        # 检查暂停状态 (获取当前状态结果数后)
                        if self.paused:
                             df_result.to_excel(new_excel_path, index=False)
                             self.update_progress(f"获取当前状态结果数时暂停，已保存进度: {new_excel_path}")
                             # 暂停后，无法确定状态，最好跳过当前，让下一个游戏重新开始判断
                             # 并且重置状态标记，避免影响下一个游戏
                             is_state_filtered = False # 重置状态
                             continue # 跳到下一个游戏
                             
                        self.update_progress(f"当前状态结果数量: {current_state_count}")
                        
                        if current_state_count > 0:
                            current_operator = game_operator_dict.get(game_name, "") 
                            current_state_owner, current_state_game_match, current_state_operator_match, current_state_manual_check = \
                                self.extract_and_match_results(driver, game_name, current_operator)
                            # 检查暂停状态 (提取当前状态结果后)
                            if self.paused:
                                df_result.to_excel(new_excel_path, index=False)
                                self.update_progress(f"提取当前状态结果时暂停，已保存进度: {new_excel_path}")
                                is_state_filtered = False # 重置状态
                                continue # 跳到下一个游戏
                        else:
                             self.update_progress("当前状态结果为0，无需提取详情")

                    except Exception as e_current_extract:
                        self.update_progress(f"提取当前状态结果时出错: {str(e_current_extract)}")
                        current_state_manual_check = f"是 (提取当前状态结果异常: {str(e_current_extract)[:30]}...)"
                        try: current_state_count = self.get_search_result_count(driver) # 再次尝试获取数量
                        except: current_state_count = -1 # 标记错误
                        # 即使出错，也可能需要根据预期状态进行筛选操作，所以不直接continue

                    # --- 步骤3 & 4 & 5: 根据预期状态决定操作并记录结果 ---
                    row_idx = row_indices.get(game_name)
                    if row_idx is None:
                        self.update_progress(f"警告：未找到游戏 '{game_name}' 在原始Excel中的行索引")
                        # 如果找不到行，我们仍然需要根据逻辑更新 is_state_filtered 标志
                        # 否则会影响下一个游戏的判断。这里简化处理：假设操作能完成并更新状态。
                        # （更健壮的方式是把状态更新放到 try...finally 中确保执行）
                        # 暂时先按找到行的逻辑走，最后更新 is_state_filtered
                        pass # 继续下面的逻辑尝试更新状态
                        
                    if is_state_filtered:
                        # === 情况A: 预期页面是已筛选状态 ===
                        self.update_progress("处理预期为[已筛选]状态的情况...")
                        if current_state_count > 0:
                             # A.1: 已筛选状态有结果 -> 直接使用
                             self.update_progress("当前已筛选状态结果 > 0，直接使用此结果")
                             if row_idx is not None: # 只有找到行才写入
                                 df_result.loc[row_idx, "搜索的结果数量"] = current_state_count 
                                 df_result.loc[row_idx, "匹配著作权人"] = current_state_owner
                                 df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = current_state_game_match
                                 df_result.loc[row_idx, "当前结果与运营单位是否一致"] = current_state_operator_match
                                 df_result.loc[row_idx, "是否建议人工排查"] = current_state_manual_check
                                 self.update_progress(f"已更新行 {row_idx} (使用当前已筛选结果)")
                                 # 下一个游戏开始时，状态依然是已筛选
                                 next_is_state_filtered = True 
                        else: # current_state_count == 0 or extraction failed
                             # A.2: 已筛选状态无结果 (或提取失败) -> 取消筛选并使用取消后的结果
                             self.update_progress("当前已筛选状态结果为 0 或提取失败，尝试取消筛选...")
                             # 显式点击筛选器以取消勾选
                             self.update_progress("尝试点击筛选器以取消...")
                             clicked_filter1_off = self.click_filter_checkbox(driver, '1年内')
                             self.random_delay(0.5, 1)
                             clicked_filter2_off = self.click_filter_checkbox(driver, '1-3年')
                             self.random_delay(0.5, 1)
                             clicked_filter3_off = self.click_filter_checkbox(driver, '3-5年')
                             self.random_delay(1, 2) # 等待页面刷新
                             
                             # 检查暂停(取消筛选时)
                             if self.paused:
                                 if row_idx is not None: # 保存一下之前的(0结果)状态
                                      df_result.loc[row_idx, "搜索的结果数量"] = current_state_count # 通常是0
                                      df_result.loc[row_idx, "匹配著作权人"] = current_state_owner # 通常是""
                                      df_result.loc[row_idx, "是否建议人工排查"] = current_state_manual_check + " (取消筛选时暂停)"
                                 df_result.to_excel(new_excel_path, index=False) 
                                 self.update_progress(f"取消筛选时暂停，已保存进度: {new_excel_path}")
                                 is_state_filtered = False # 重置状态
                                 continue
                             
                             if not clicked_filter1_off or not clicked_filter2_off or not clicked_filter3_off:
                                 self.update_progress("警告：取消筛选操作可能未完全成功")
                             else:
                                 self.update_progress("已尝试点击取消筛选")
                             
                             # 获取取消筛选后的结果
                             unfiltered_count = 0
                             unfiltered_owner = ""
                             unfiltered_game_match = "否"
                             unfiltered_operator_match = "否"
                             unfiltered_manual_check = "是"
                             try:
                                 unfiltered_count = self.get_search_result_count(driver)
                                 if self.paused: raise Exception("获取取消筛选后结果数时暂停") # 主动抛出以便统一处理
                                 self.update_progress(f"取消筛选后结果数量: {unfiltered_count}")
                                 if unfiltered_count > 0:
                                     current_operator = game_operator_dict.get(game_name, "") 
                                     unfiltered_owner, unfiltered_game_match, unfiltered_operator_match, unfiltered_manual_check = \
                                         self.extract_and_match_results(driver, game_name, current_operator)
                                     if self.paused: raise Exception("提取取消筛选后结果时暂停")
                             except Exception as e_unfilter:
                                 self.update_progress(f"获取或提取取消筛选后结果时出错: {e_unfilter}")
                                 unfiltered_manual_check = f"是 (取消筛选后处理异常: {str(e_unfilter)[:30]}...)"
                                 if self.paused: # 如果是暂停导致的异常
                                     if row_idx is not None: # 保存一下之前的(0结果)状态
                                          df_result.loc[row_idx, "搜索的结果数量"] = current_state_count
                                          df_result.loc[row_idx, "是否建议人工排查"] = current_state_manual_check + " (处理取消筛选结果时暂停)"
                                     df_result.to_excel(new_excel_path, index=False)
                                     self.update_progress(f"处理取消筛选结果时暂停，已保存进度: {new_excel_path}")
                                     is_state_filtered = False # 重置状态
                                     continue
                                     
                             # 记录取消筛选后的结果
                             if row_idx is not None:
                                 df_result.loc[row_idx, "搜索的结果数量"] = unfiltered_count 
                                 df_result.loc[row_idx, "匹配著作权人"] = unfiltered_owner
                                 df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = unfiltered_game_match
                                 df_result.loc[row_idx, "当前结果与运营单位是否一致"] = unfiltered_operator_match
                                 df_result.loc[row_idx, "是否建议人工排查"] = unfiltered_manual_check + " (来自取消筛选)"
                                 self.update_progress(f"已更新行 {row_idx} (使用取消筛选后的结果)")
                                 # 下一个游戏开始时，状态是未筛选
                                 next_is_state_filtered = False
                    else:
                        # === 情况B: 预期页面是未筛选状态 ===
                        self.update_progress("处理预期为[未筛选]状态的情况...")
                        # 先记录下当前的未筛选结果 (即使是0也要记，用于后续对比)
                        initial_unfiltered_count = current_state_count
                        initial_unfiltered_owner = current_state_owner
                        initial_unfiltered_game_match = current_state_game_match
                        initial_unfiltered_operator_match = current_state_operator_match
                        initial_unfiltered_manual_check = current_state_manual_check
                        self.update_progress(f"已记录初始未筛选结果: 数量={initial_unfiltered_count}")
                        
                        # 尝试应用筛选
                        self.update_progress("尝试应用筛选条件...")
                        clicked_filter1_on = self.click_filter_checkbox(driver, '1年内')
                        self.random_delay(0.5, 1)
                        clicked_filter2_on = self.click_filter_checkbox(driver, '1-3年')
                        self.random_delay(0.5, 1)
                        clicked_filter3_on = self.click_filter_checkbox(driver, '3-5年')
                        self.random_delay(1, 2) # 等待页面刷新
                        
                        # 检查暂停(应用筛选时)
                        if self.paused:
                             if row_idx is not None: # 保存初始结果
                                 df_result.loc[row_idx, "搜索的结果数量"] = initial_unfiltered_count 
                                 df_result.loc[row_idx, "匹配著作权人"] = initial_unfiltered_owner
                                 df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = initial_unfiltered_game_match
                                 df_result.loc[row_idx, "当前结果与运营单位是否一致"] = initial_unfiltered_operator_match
                                 df_result.loc[row_idx, "是否建议人工排查"] = initial_unfiltered_manual_check + " (应用筛选时暂停)"
                             df_result.to_excel(new_excel_path, index=False)
                             self.update_progress(f"应用筛选时暂停，已使用初始未筛选结果保存进度: {new_excel_path}")
                             is_state_filtered = False # 重置状态
                             continue
                        
                        if not clicked_filter1_on or not clicked_filter2_on or not clicked_filter3_on:
                             self.update_progress("警告：应用筛选操作可能未完全成功")
                        else:
                             self.update_progress("已尝试点击应用筛选")
                             
                        # 获取显式应用筛选后的结果
                        filtered_count = 0
                        filtered_owner = ""
                        filtered_game_match = "否"
                        filtered_operator_match = "否"
                        filtered_manual_check = "是"
                        try:
                            filtered_count = self.get_search_result_count(driver)
                            if self.paused: raise Exception("获取应用筛选后结果数时暂停")
                            self.update_progress(f"应用筛选后结果数量: {filtered_count}")
                            if filtered_count > 0:
                                current_operator = game_operator_dict.get(game_name, "") 
                                filtered_owner, filtered_game_match, filtered_operator_match, filtered_manual_check = \
                                    self.extract_and_match_results(driver, game_name, current_operator)
                                if self.paused: raise Exception("提取应用筛选后结果时暂停")
                        except Exception as e_filter_on:
                            self.update_progress(f"获取或提取应用筛选后结果时出错: {e_filter_on}")
                            filtered_manual_check = f"是 (应用筛选后处理异常: {str(e_filter_on)[:30]}...)"
                            if self.paused:
                                if row_idx is not None: # 保存初始结果
                                    df_result.loc[row_idx, "搜索的结果数量"] = initial_unfiltered_count 
                                    df_result.loc[row_idx, "匹配著作权人"] = initial_unfiltered_owner
                                    df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = initial_unfiltered_game_match
                                    df_result.loc[row_idx, "当前结果与运营单位是否一致"] = initial_unfiltered_operator_match
                                    df_result.loc[row_idx, "是否建议人工排查"] = initial_unfiltered_manual_check + " (处理筛选结果时暂停)"
                                df_result.to_excel(new_excel_path, index=False)
                                self.update_progress(f"处理应用筛选结果时暂停，已使用初始未筛选结果保存进度: {new_excel_path}")
                                is_state_filtered = False # 重置状态
                                continue
                                
                        # 判断使用哪个结果
                        if row_idx is not None: # 只有找到行才写入和更新状态
                            if filtered_count > 0:
                                 # B.1: 应用筛选后有结果 -> 使用筛选后的
                                 self.update_progress("应用筛选后结果 > 0，使用筛选后结果")
                                 df_result.loc[row_idx, "搜索的结果数量"] = filtered_count 
                                 df_result.loc[row_idx, "匹配著作权人"] = filtered_owner
                                 df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = filtered_game_match
                                 df_result.loc[row_idx, "当前结果与运营单位是否一致"] = filtered_operator_match
                                 df_result.loc[row_idx, "是否建议人工排查"] = filtered_manual_check
                                 self.update_progress(f"已更新行 {row_idx} (使用应用筛选后的结果)")
                                 # 下一个游戏开始时，状态是已筛选
                                 next_is_state_filtered = True
                            else: # filtered_count == 0
                                 # B.2: 应用筛选后无结果 -> 使用初始未筛选的
                                 self.update_progress("应用筛选后结果为 0，回退使用初始未筛选结果")
                                 df_result.loc[row_idx, "搜索的结果数量"] = initial_unfiltered_count 
                                 df_result.loc[row_idx, "匹配著作权人"] = initial_unfiltered_owner
                                 df_result.loc[row_idx, "当前结果与游戏简称是否一致"] = initial_unfiltered_game_match
                                 df_result.loc[row_idx, "当前结果与运营单位是否一致"] = initial_unfiltered_operator_match
                                 # 修改人工排查建议以反映情况
                                 manual_check_note = initial_unfiltered_manual_check
                                 if manual_check_note == "否": manual_check_note = "否 (筛选后无结果，使用筛选前数据)"
                                 elif manual_check_note.startswith("是"): manual_check_note += " (筛选后无结果)"
                                 else: manual_check_note += " (筛选后无结果)"
                                 df_result.loc[row_idx, "是否建议人工排查"] = manual_check_note
                                 self.update_progress(f"已更新行 {row_idx} (使用初始未筛选结果)")
                                 # 下一个游戏开始时，状态是未筛选 (因为筛选尝试失败了)
                                 next_is_state_filtered = False
                        else: # row_idx is None
                            # 如果没找到行，无法写入，但需要猜测下一个状态
                            # 保守起见，假设如果尝试应用筛选，状态会变；如果尝试取消，状态也会变
                            if filtered_count > 0: # B.1发生
                                next_is_state_filtered = True
                            else: # B.2发生
                                next_is_state_filtered = False
                                
                    # 在循环末尾更新状态标记，供下一次迭代使用
                    is_state_filtered = next_is_state_filtered
                    
                    # --- 原来的步骤4/5 (获取筛选后数量/判断写入) 被上面的逻辑覆盖，移除或注释 --- 
                    # self.update_progress("\\n获取筛选后的结果数量...")

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