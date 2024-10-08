# utils/large_model.py

import os
from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline
from huggingface_hub import HfApi
from docx import Document
from docx.shared import RGBColor
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import requests

MODEL_NAME = "nlptown/bert-base-multilingual-uncased-sentiment"  # 示例模型，可根据需要更换
MODEL_PATH = os.path.join(os.getcwd(), "models")

def check_model_configured():
    """检查模型和分词器是否已下载并配置"""
    model_dir = os.path.join(MODEL_PATH, MODEL_NAME.replace('/', '_'))
    # 根据实际模型文件调整所需文件列表
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

def analyze_file_with_model(file_path, progress_callback):
    """
    使用大模型分析文件内容，并根据情感结果进行标记。
    支持 .docx 和 .xlsx 文件。
    返回分析结果的统计信息。
    """
    # 检查模型是否已配置
    if not check_model_configured():
        raise Exception("大模型未配置，请先下载并配置大模型。")

    # 加载模型
    progress_callback("加载模型中...")
    try:
        model_dir = os.path.join(MODEL_PATH, MODEL_NAME.replace('/', '_'))
        tokenizer = AutoTokenizer.from_pretrained(model_dir)
        model = AutoModelForSequenceClassification.from_pretrained(model_dir)
        classifier = pipeline('sentiment-analysis', model=model, tokenizer=tokenizer)
    except Exception as e:
        raise Exception(f"加载模型时发生错误：{e}")

    # 初始化统计信息
    total_word_count = 0
    positive_count = 0
    neutral_count = 0
    negative_count = 0

    # 分析文件
    file_type = file_path.split('.')[-1].lower()
    if file_type == 'docx':
        doc = Document(file_path)
        total_paragraphs = len(doc.paragraphs)
        for idx, para in enumerate(doc.paragraphs, start=1):
            text = para.text.strip()
            total_word_count += len(text)
            if text:
                try:
                    result = classifier(text[:512])[0]  # 截断到512字符以内
                    label = result['label']
                    # 根据模型结果进行标记
                    if label in ['1 star', '2 stars']:
                        # 标记为黄色
                        for run in para.runs:
                            run.font.color.rgb = RGBColor(255, 255, 0)  # 黄色
                        negative_count += 1
                    elif label in ['3 stars']:
                        neutral_count += 1
                    elif label in ['4 stars', '5 stars']:
                        positive_count += 1
                except Exception as e:
                    progress_callback(f"分析段落 {idx} 时发生错误：{e}")
            progress = int(idx / total_paragraphs * 100)
            progress_callback(f"分析进度: {progress}%")
        new_file_path = file_path.replace('.docx', '_analyzed.docx')
        doc.save(new_file_path)
    elif file_type == 'xlsx':
        wb = load_workbook(file_path)
        for sheet in wb.worksheets:
            total_rows = sheet.max_row
            for idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=total_rows), start=1):
                for cell in row:
                    if cell.value:
                        text = str(cell.value).strip()
                        total_word_count += len(text)
                        try:
                            result = classifier(text[:512])[0]
                            label = result['label']
                            if label in ['1 star', '2 stars']:
                                # 标记为黄色
                                fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                                cell.fill = fill
                                negative_count += 1
                            elif label in ['3 stars']:
                                neutral_count += 1
                            elif label in ['4 stars', '5 stars']:
                                positive_count += 1
                        except Exception as e:
                            progress_callback(f"分析单元格 {cell.coordinate} 时发生错误：{e}")
                progress = int(idx / total_rows * 100)
                progress_callback(f"分析进度: {progress}%")
        new_file_path = file_path.replace('.xlsx', '_analyzed.xlsx')
        wb.save(new_file_path)
    else:
        raise ValueError("仅支持 .docx 和 .xlsx 文件")

    progress_callback("分析完成！")

    # 构建结果统计信息
    result = {
        'total_word_count': total_word_count,
        'positive_count': positive_count,
        'neutral_count': neutral_count,
        'negative_count': negative_count,
        'new_file_path': new_file_path
    }

    return result
