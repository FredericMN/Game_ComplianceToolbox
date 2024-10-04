# 文件: project-01/utils/crawler.py

import os  # 添加 os 模块用于文件名操作
import time
import datetime
import openpyxl
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def crawl_new_games(start_date_str, end_date_str, progress_callback):
    # 配置 WebDriver
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) " 
                         "AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/113.0.0.0 Safari/537.36")
    # 如果不需要显示浏览器界面，可以取消下一行的注释
    # options.add_argument('--headless')
    
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
    
    try:
        # 准备 Excel 文件
        today_str = datetime.date.today().strftime("%Y-%m-%d")
        excel_filename = f"游戏数据_{today_str}.xlsx"
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "游戏数据"
        headers = ["日期", "状态", "名称", "厂商", "类型", "评分"]
        ws.append(headers)
        
        # 处理日期范围
        if start_date_str and end_date_str:
            try:
                start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
                end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
            except ValueError:
                progress_callback("日期格式不正确，请输入yyyy-mm-dd格式的日期")
                driver.quit()
                return
            if end_date < start_date:
                progress_callback("结束日期不能早于起始日期")
                driver.quit()
                return
            delta_days = (end_date - start_date).days + 1
            date_list = [start_date + datetime.timedelta(days=i) for i in range(delta_days)]
        else:
            start_date = datetime.date.today()
            date_list = [start_date + datetime.timedelta(days=i) for i in range(5)]
        
        # 爬取数据
        for crawl_date in date_list:
            crawl_date_str = crawl_date.strftime("%Y-%m-%d")
            progress_callback(f"正在爬取日期：{crawl_date_str}")
            base_url = f"https://www.taptap.cn/app-calendar/{crawl_date_str}"
            driver.get(base_url)
            
            try:
                games = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.daily-event-list__content > a.tap-router"))
                )
            except Exception as e:
                progress_callback(f"加载游戏列表时出错：{e}")
                continue

            progress_callback(f"找到 {len(games)} 个游戏。")
            for index, game in enumerate(games, start=1):
                try:
                    # 提取游戏名称
                    try:
                        name_element = game.find_element(By.CSS_SELECTOR, "div.daily-event-app-info__title")
                        name = name_element.get_attribute("content").strip()
                    except:
                        name = "未知名称"

                    # 提取游戏类型
                    try:
                        type_elements = game.find_elements(By.CSS_SELECTOR, "div.daily-event-app-info__tag div.tap-label-tag")
                        types = "/".join([type_elem.text.strip() for type_elem in type_elements])
                    except:
                        types = "未知类型"

                    # 提取评分
                    try:
                        rating_element = game.find_element(By.CSS_SELECTOR, "div.daily-event-app-info__rating .tap-rating__number")
                        rating = rating_element.text.strip()
                    except:
                        rating = "未知评分"

                    # 提取状态
                    try:
                        status_element = game.find_element(By.CSS_SELECTOR, "span.event-type-label__title")
                        status = status_element.text.strip()
                    except:
                        try:
                            status_element = game.find_element(By.CSS_SELECTOR, "div.event-recommend-label__title")
                            status = status_element.text.strip()
                        except:
                            status = "未知状态"

                    progress_callback(f"正在爬取第 {index} 个游戏：{name}")

                    # 提取游戏链接
                    href = game.get_attribute("href")
                    game_url = href if href.startswith("http") else f"https://www.taptap.cn{href}"

                    # 打开游戏子页面
                    driver.execute_script("window.open(arguments[0]);", game_url)
                    driver.switch_to.window(driver.window_handles[1])

                    # 等待页面加载完成
                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.single-info__content"))
                        )
                    except Exception:
                        progress_callback("找不到厂商信息")
                        manufacturer = "未知厂商"
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                        ws.append([crawl_date_str, status, name, manufacturer, types, rating])
                        time.sleep(0.2)
                        continue

                    # 提取厂商信息
                    try:
                        info_blocks = driver.find_elements(By.CSS_SELECTOR, "div.single-info__content")
                        manufacturer = "未知厂商"
                        priority = ["厂商", "发行", "发行商", "开发"]
                        info_dict = {}
                        for block in info_blocks:
                            try:
                                label_element = block.find_element(By.CSS_SELECTOR, ".caption-m12-w12")
                                label = label_element.text.strip()
                                value_element = block.find_element(By.CSS_SELECTOR, ".single-info__content__value")
                                value = value_element.text.strip()
                                info_dict[label] = value
                            except:
                                continue
                        for key in priority:
                            if key in info_dict:
                                manufacturer = info_dict[key]
                                break
                        if manufacturer == "未知厂商":
                            progress_callback("找不到厂商信息")
                    except:
                        manufacturer = "未知厂商"
                        progress_callback("找不到厂商信息")

                    progress_callback(f"名称: {name}, 厂商: {manufacturer}, 类型: {types}, 评分: {rating}, 状态: {status}")

                    # 关闭子页面并切换回主页面
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

                    # 追加数据到 Excel
                    ws.append([crawl_date_str, status, name, manufacturer, types, rating])

                    # 等待一段时间
                    time.sleep(0.2)
                except Exception as e:
                    progress_callback(f"第 {index} 个游戏爬取失败：{e}")
                    # 关闭可能打开的子页面
                    if len(driver.window_handles) > 1:
                        driver.close()
                        driver.switch_to.window(driver.window_handles[0])
                    # 追加数据为未知
                    ws.append([crawl_date_str, "未知状态", "未知名称", "未知厂商", "未知类型", "未知评分"])
                    time.sleep(0.2)
                    continue

        # 保存 Excel 文件
        wb.save(excel_filename)
        progress_callback(f"数据已成功保存到 {excel_filename}，接下来将继续匹配版号数据，请稍后……")

        # 调用版号匹配函数，并传递是否添加后缀的参数
        match_version_numbers(excel_filename, progress_callback, append_suffix=False)

    finally:
        # 关闭 WebDriver
        driver.quit()

def match_version_numbers(excel_filename, progress_callback, append_suffix=True):
    # 配置 WebDriver
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) " 
                         "AppleWebKit/537.36 (KHTML, like Gecko) "
                         "Chrome/113.0.0.0 Safari/537.36")
    # 如果不需要显示浏览器界面，可以取消下一行的注释
    # options.add_argument('--headless')
    
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
    
    try:
        # 打开 Excel 文件
        try:
            wb = openpyxl.load_workbook(excel_filename)
        except FileNotFoundError:
            progress_callback(f"文件 {excel_filename} 未找到。")
            driver.quit()
            return

        ws = wb.active

        # 检查并添加新表头（包含“游戏名称”）
        existing_headers = [cell.value for cell in ws[1]]
        new_headers = ["游戏名称", "出版单位", "运营单位", "文号", "出版物号", "时间", "游戏类型", "申报类别", "是否存在多个结果"]
        
        for header in new_headers:
            if header not in existing_headers:
                ws.cell(row=1, column=len(existing_headers) + 1, value=header)
                existing_headers.append(header)
        
        # 确定新表头的起始列索引
        new_header_start_col = len(existing_headers) - len(new_headers) + 1  # 1-based index
        
        # 自动检测游戏名称所在的列（“游戏名称”或“名称”）
        game_name_col = None
        for idx, header in enumerate(existing_headers, start=1):
            if header in ["游戏名称", "名称"]:
                game_name_col = idx
                break
        
        if game_name_col is None:
            progress_callback("未找到表头中的“游戏名称”或“名称”列，请检查Excel文件的表头。")
            driver.quit()
            return
        
        # 从现有 Excel 中读取所有游戏名称及其所在行（跳过表头）
        game_rows = []
        for row in ws.iter_rows(min_row=2, values_only=False):
            game_name = row[game_name_col - 1].value  # openpyxl的列索引从0开始
            game_rows.append((row, game_name))
        
        # 定义函数以移除游戏名称中的括号及其内容
        def clean_game_name(name):
            if not name:
                return ""
            # 移除中文括号及内容
            name = re.sub(r'（[^）]*）', '', name)
            # 移除英文括号及内容
            name = re.sub(r'\([^)]*\)', '', name)
            return name.strip()
        
        # 爬取版号信息的函数
        def get_game_info_from_nppa(game_name_original):
            game_name_clean = clean_game_name(game_name_original)
            search_url = f"https://www.nppa.gov.cn/bsfw/jggs/cxjg/index.html?mc={game_name_clean}&cbdw=&yydw=&wh=undefined&description=#"
            driver.get(search_url)
            
            try:
                # 等待查询结果加载
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#dataCenter"))
                )
                time.sleep(1)  # 确保内容完全加载

                results = driver.find_elements(By.CSS_SELECTOR, "#dataCenter tr")
                if len(results) == 0:
                    return ["", "", "", "", "", "", "", "", "否"]  # 无结果
                elif len(results) > 1:
                    # 多个结果时，先找完全匹配的，否则选择第一个结果
                    exact_match = None
                    for result in results:
                        try:
                            game_title = result.find_element(By.CSS_SELECTOR, "a").text.strip()
                            if game_title == game_name_original:
                                exact_match = result
                                break
                        except:
                            continue
                    if exact_match:
                        return extract_game_info(exact_match, "是")
                    else:
                        return extract_game_info(results[0], "是")
                else:
                    return extract_game_info(results[0], "否")
            except Exception:
                # 在匹配过程中遇到异常，返回默认值
                return ["", "", "", "", "", "", "", "", "否"]  # 出现异常时

        # 提取游戏信息的函数
        def extract_game_info(result, multiple_results):
            try:
                cells = result.find_elements(By.TAG_NAME, "td")
                if len(cells) < 7:
                    return ["", "", "", "", "", "", "", "", multiple_results]
                
                game_name = cells[1].text.strip()
                publisher = cells[2].text.strip()
                operator = cells[3].text.strip()
                approval_number = cells[4].text.strip()
                publication_number = cells[5].text.strip()
                date = cells[6].text.strip()

                # 获取详情页链接
                try:
                    detail_url = cells[1].find_element(By.TAG_NAME, "a").get_attribute("href")
                except:
                    detail_url = ""

                if detail_url:
                    # 打开详情页
                    driver.execute_script("window.open(arguments[0]);", detail_url)
                    driver.switch_to.window(driver.window_handles[1])

                    try:
                        # 等待详情页内容加载
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, ".cFrame.nFrame"))
                        )
                        time.sleep(1)  # 确保内容完全加载

                        # 提取游戏类型和申报类别
                        game_type = ""
                        apply_category = ""
                        rows = driver.find_elements(By.CSS_SELECTOR, ".cFrame.nFrame table tr")
                        for row in rows:
                            try:
                                label = row.find_element(By.XPATH, "./td[1]").text.strip()
                                value = row.find_element(By.XPATH, "./td[2]").text.strip()
                                if label == "游戏类型":
                                    game_type = value
                                elif label == "申报类别":
                                    apply_category = value
                            except:
                                continue
                    except Exception:
                        # 在提取详情页信息时遇到异常，设置默认值
                        game_type = ""
                        apply_category = ""

                    # 关闭详情页并切换回主窗口
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                else:
                    game_type = ""
                    apply_category = ""

                return [
                    game_name, publisher, operator, approval_number, 
                    publication_number, date, game_type, apply_category, multiple_results
                ]
            except Exception:
                # 在提取游戏信息时遇到异常，返回默认值
                return ["", "", "", "", "", "", "", "", multiple_results]

        # 遍历每个游戏名称并爬取版号信息
        for idx, (row_cells, game_name_original) in enumerate(game_rows, start=2):
            if game_name_original:
                game_name_display = game_name_original
                progress_callback(f"正在查找游戏: {game_name_display} (第 {idx-1} 个游戏)")
                game_info = get_game_info_from_nppa(game_name_original)
                
                # 将信息写入对应行的指定列（包括“游戏名称”）
                for i, info in enumerate(game_info):
                    ws.cell(row=idx, column=new_header_start_col + i, value=info)
                
                # 输出对应的信息到控制台
                output_info = {
                    "游戏名称": game_info[0],
                    "出版单位": game_info[1],
                    "运营单位": game_info[2],
                    "文号": game_info[3],
                    "出版物号": game_info[4],
                    "时间": game_info[5],
                    "游戏类型": game_info[6],
                    "申报类别": game_info[7],
                    "是否存在多个结果": game_info[8]
                }
                progress_callback(f"查询结果: {output_info}\n")
                
                # 等待一段时间以避免过快请求
                time.sleep(1)
            else:
                progress_callback(f"第 {idx-1} 行游戏名称为空，跳过。")
                continue

        # 确定输出文件名
        if append_suffix:
            base, ext = os.path.splitext(excel_filename)
            output_filename = f"{base}-已匹配版号数据{ext}"
        else:
            output_filename = excel_filename

        # 保存更新后的 Excel 文件
        try:
            wb.save(output_filename)
            progress_callback(f"版号信息已成功补充到 {output_filename}")
        except PermissionError:
            progress_callback(f"无法保存文件 {output_filename}。请确保文件未被打开并重试。")
    
    finally:
        # 关闭 WebDriver
        driver.quit()
