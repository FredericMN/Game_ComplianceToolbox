# utils/large_model.py

import os
from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline
from huggingface_hub import HfApi
from docx import Document
from docx.shared import RGBColor
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import requests

MODEL_NAME = "alibaba-pai/pai-bert-base-zh-llm-risk-detection"  # 更新为目标模型
MODEL_PATH = os.path.join(os.getcwd(), "models")

# 全局变量用于缓存模型
_cached_classifier = None


def check_model_configured():
    """检查模型和分词器是否已下载并配置"""
    model_dir = os.path.join(MODEL_PATH, MODEL_NAME.replace('/', '_'))
    # 根据 alibaba-pai/pai-bert-base-zh-llm-risk-detection 模型调整所需文件列表
    required_files = ["pytorch_model.bin", "config.json", "tokenizer_config.json", "vocab.txt"]
    return all(os.path.exists(os.path.join(model_dir, f)) for f in required_files)


def check_and_download_model(progress_callback):
    """
    检查并下载大模型。如果模型已配置完成，则直接返回。
    否则，下载模型和分词器，并通过 progress_callback 更新下载进度。
    """
    api = HfApi()
    model_dir = os.path.join(MODEL_PATH, MODEL_NAME.replace('/', '_'))
    if not os.path.exists(model_dir):
        os.makedirs(model_dir)
    if check_model_configured():
        progress_callback("大模型已配置完成。")
        return
    else:
        progress_callback("正在下载大模型，请稍候...")

    try:
        # 获取仓库中的所有文件
        repo_files = api.list_repo_files(repo_id=MODEL_NAME)
        total_files = len(repo_files)
        progress_callback(f"总共需要下载 {total_files} 个文件。")

        # 假设使用 'main' 版本
        revision = "main"

        for idx, file_name in enumerate(repo_files, start=1):
            progress_callback(f"正在下载文件 {idx}/{total_files}: {file_name}")
            # 构建下载 URL
            download_url = f"https://huggingface.co/{MODEL_NAME}/resolve/{revision}/{file_name}"
            # 本地文件路径
            local_file_path = os.path.join(model_dir, file_name)
            temp_file_path = local_file_path + ".part"  # 临时文件名
            os.makedirs(os.path.dirname(local_file_path), exist_ok=True)
            # 下载文件
            with requests.get(download_url, stream=True) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                downloaded_size = 0
                chunk_size = 8192
                # 为减少信号发送频率，设定一个阈值
                update_threshold = chunk_size * 100  # 每下载100个块更新一次
                next_update = update_threshold

                with open(temp_file_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=chunk_size):
                        if chunk:
                            f.write(chunk)
                            downloaded_size += len(chunk)
                            if downloaded_size >= next_update or downloaded_size == total_size:
                                if total_size > 0:
                                    percent = int(downloaded_size * 100 / total_size)
                                else:
                                    percent = 0
                                progress_callback(f"下载进度: {percent}% ({downloaded_size}/{total_size} bytes)")
                                next_update += update_threshold
            # 下载完成后重命名为目标文件名
            os.rename(temp_file_path, local_file_path)
            progress_callback(f"已下载文件 {idx}/{total_files}: {file_name}")

        progress_callback("大模型下载完成并已配置。")
    except Exception as e:
        raise Exception(f"下载大模型时发生错误：{e}")


def load_classifier(device='cpu'):
    """
    加载并缓存分类器，以避免重复加载模型。
    """
    global _cached_classifier
    if _cached_classifier is None:
        model_dir = os.path.join(MODEL_PATH, MODEL_NAME.replace('/', '_'))
        tokenizer = AutoTokenizer.from_pretrained(model_dir)
        model = AutoModelForSequenceClassification.from_pretrained(model_dir)
        _cached_classifier = pipeline(
            'text-classification',
            model=model,
            tokenizer=tokenizer,
            device=0 if device == 'cuda' else -1,
            max_length=128,
            top_k=None,  # 使用 top_k=None 代替 return_all_scores=True
            truncation=True  # 显式启用文本长度截断
        )
    return _cached_classifier


def analyze_files_with_model(file_paths, progress_callback, device='cpu', normal_threshold=0.95, other_threshold=0.5):
    """
    使用大模型分析多个文件内容，并根据风险等级进行标记。
    支持 .docx 和 .xlsx 文件。
    返回分析结果的列表，每个元素对应一个文件的统计信息。
    """
    # 检查模型是否已配置
    if not check_model_configured():
        raise Exception("大模型未配置，请先下载并配置大模型。")

    # 加载模型
    progress_callback("加载模型中...")
    try:
        classifier = load_classifier(device)
    except Exception as e:
        raise Exception(f"加载模型时发生错误：{e}")

    results = []

    # 预先计算总处理项数
    total_items = 0
    items_per_file = []  # 存储每个文件的项数
    for file_path in file_paths:
        file_type = file_path.split('.')[-1].lower()
        if file_type == 'docx':
            doc = Document(file_path)
            paragraphs = [para for para in doc.paragraphs if para.text.strip()]
            items = len(paragraphs)
        elif file_type == 'xlsx':
            wb = load_workbook(file_path, read_only=True)
            cells = [cell for sheet in wb.worksheets for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row) for cell in row if cell.value]
            items = len(cells)
        else:
            items = 0  # Unsupported file type
        items_per_file.append(items)
        total_items += items

    if total_items == 0:
        raise Exception("未找到任何可处理的段落或单元格。")

    progress_callback(f"总共有 {total_items} 个待处理的段落或单元格。")

    # 初始化全局进度跟踪
    global_progress = {'processed': 0}

    for file_idx, file_path in enumerate(file_paths, start=1):
        result = {
            'file_path': file_path,
            'total_word_count': 0,
            'normal_count': 0,
            'low_vulgar_count': 0,
            'porn_count': 0,
            'other_risk_count': 0,
            'adult_count': 0,
            'new_file_path': ''
        }
        try:
            progress_callback(f"开始分析文件 {file_idx}/{len(file_paths)}: {os.path.basename(file_path)}")
            # 分析单个文件
            file_result = analyze_single_file_with_model(
                file_path, classifier, progress_callback, global_progress, total_items,
                normal_threshold, other_threshold)
            result.update(file_result)
            results.append(result)
            progress_callback(f"完成分析文件 {file_idx}/{len(file_paths)}: {os.path.basename(file_path)}\n")
        except Exception as e:
            progress_callback(f"分析文件 {os.path.basename(file_path)} 时发生错误：{e}\n")
            results.append(result)  # 记录错误但继续

    progress_callback("所有文件分析完成！")
    return results


def analyze_single_file_with_model(file_path, classifier, progress_callback, global_progress, total_items, normal_threshold, other_threshold):
    """
    使用大模型分析单个文件内容，并根据风险等级进行标记。
    支持 .docx 和 .xlsx 文件。
    返回分析结果的统计信息。
    """
    # 初始化统计信息
    total_word_count = 0
    normal_count = 0
    low_vulgar_count = 0
    porn_count = 0
    other_risk_count = 0
    adult_count = 0

    # 分析文件
    file_type = file_path.split('.')[-1].lower()
    if file_type == 'docx':
        doc = Document(file_path)
        paragraphs = [para for para in doc.paragraphs if para.text.strip()]
        total_paragraphs = len(paragraphs)
        texts = [para.text.strip() for para in paragraphs]
        total_word_count = sum(len(text) for text in texts)

        # 批量分类
        batch_size = 32  # 根据需要调整批量大小
        labels = []
        for i in range(0, len(texts), batch_size):
            batch_texts = texts[i:i + batch_size]
            try:
                results_batch = classifier(batch_texts)
                for res in results_batch:
                    # 查找“正常”类别的分数
                    normal_score = next((item['score'] for item in res if item['label'] == "正常"), 0)
                    if normal_score >= normal_threshold:
                        labels.append("正常")
                    else:
                        # 获取其他类别中分数最高的类别
                        other_scores = [item for item in res if item['label'] != "正常"]
                        if not other_scores:
                            labels.append("其他风险")
                            continue
                        max_other = max(other_scores, key=lambda x: x['score'])
                        if max_other['score'] >= other_threshold:
                            labels.append(max_other['label'])
                        else:
                            labels.append("其他风险")
            except Exception as e:
                progress_callback(f"批量分析段落 {i + 1} - {i + len(batch_texts)} 时发生错误：{e}")
                labels.extend(["未知"] * len(batch_texts))  # 标记为未知

            # 更新全局进度
            processed = len(batch_texts)
            global_progress['processed'] += processed
            percent = int((global_progress['processed'] / total_items) * 100)
            progress_callback(f"进度: {percent}%")

        # 应用标签和标记
        for para, label in zip(paragraphs, labels):
            if label == "正常":
                normal_count += 1
                # 无标记
            elif label == "低俗":
                # 标记为绿色
                for run in para.runs:
                    run.font.color.rgb = RGBColor(0, 255, 0)  # 绿色
                low_vulgar_count += 1
            elif label == "色情":
                # 标记为黄色
                for run in para.runs:
                    run.font.color.rgb = RGBColor(255, 255, 0)  # 黄色
                porn_count += 1
            elif label == "其他风险":
                # 标记为红色
                for run in para.runs:
                    run.font.color.rgb = RGBColor(255, 0, 0)  # 红色
                other_risk_count += 1
            elif label == "成人":
                # 标记为蓝色
                for run in para.runs:
                    run.font.color.rgb = RGBColor(0, 0, 255)  # 蓝色
                adult_count += 1
            else:
                # 未知标签不做处理
                pass

        new_file_path = file_path.replace('.docx', '_analyzed.docx')
        doc.save(new_file_path)

    elif file_type == 'xlsx':
        wb = load_workbook(file_path)
        for sheet in wb.worksheets:
            cells = [cell for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row) for cell in row if cell.value]
            total_cells = len(cells)
            texts = [str(cell.value).strip() for cell in cells]
            total_word_count += sum(len(text) for text in texts)

            # 批量分类
            batch_size = 64  # 根据需要调整批量大小
            labels = []
            for i in range(0, len(texts), batch_size):
                batch_texts = texts[i:i + batch_size]
                try:
                    results_batch = classifier(batch_texts)
                    for res in results_batch:
                        # 查找“正常”类别的分数
                        normal_score = next((item['score'] for item in res if item['label'] == "正常"), 0)
                        if normal_score >= normal_threshold:
                            labels.append("正常")
                        else:
                            # 获取其他类别中分数最高的类别
                            other_scores = [item for item in res if item['label'] != "正常"]
                            if not other_scores:
                                labels.append("其他风险")
                                continue
                            max_other = max(other_scores, key=lambda x: x['score'])
                            if max_other['score'] >= other_threshold:
                                labels.append(max_other['label'])
                            else:
                                labels.append("其他风险")
                except Exception as e:
                    progress_callback(f"批量分析单元格 {i + 1} - {i + len(batch_texts)} 时发生错误：{e}")
                    labels.extend(["未知"] * len(batch_texts))  # 标记为未知

                # 更新全局进度
                processed = len(batch_texts)
                global_progress['processed'] += processed
                percent = int((global_progress['processed'] / total_items) * 100)
                progress_callback(f"进度: {percent}%")

            # 应用标签和标记
            for cell, label in zip(cells, labels):
                if label == "正常":
                    normal_count += 1
                    # 无标记
                elif label == "低俗":
                    # 标记为绿色
                    fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')  # 绿色
                    cell.fill = fill
                    low_vulgar_count += 1
                elif label == "色情":
                    # 标记为黄色
                    fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 黄色
                    cell.fill = fill
                    porn_count += 1
                elif label == "其他风险":
                    # 标记为红色
                    fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')  # 红色
                    cell.fill = fill
                    other_risk_count += 1
                elif label == "成人":
                    # 标记为蓝色
                    fill = PatternFill(start_color='0000FF', end_color='0000FF', fill_type='solid')  # 蓝色
                    cell.fill = fill
                    adult_count += 1
                else:
                    # 未知标签不做处理
                    pass

        new_file_path = file_path.replace('.xlsx', '_analyzed.xlsx')
        wb.save(new_file_path)

    else:
        raise ValueError("仅支持 .docx 和 .xlsx 文件")

    progress_callback("分析完成！")

    # 构建结果统计信息
    result = {
        'total_word_count': total_word_count,
        'normal_count': normal_count,
        'low_vulgar_count': low_vulgar_count,
        'porn_count': porn_count,
        'other_risk_count': other_risk_count,
        'adult_count': adult_count,
        'new_file_path': new_file_path
    }

    return result
