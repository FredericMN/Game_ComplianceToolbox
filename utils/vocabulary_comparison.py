# project-01/utils/vocabulary_comparison.py

import os
import re
import pandas as pd
from typing import List, Set, Dict
from PySide6.QtCore import QObject, Signal
from datetime import datetime

def detect_delimiters(sample_text: str) -> List[str]:
    """
    从样本文本中检测分隔符。
    统计标点符号的出现频率，选择出现频率最高的作为分隔符。
    排除点号（.）以避免分割网址等包含点号的词汇。
    """
    delimiter_counts = {}
    # 包含标准ASCII标点和Unicode标点，排除点号（.）
    # 标准ASCII标点: !"#$%&'()*+,-/ :;<=>?@[\]^_`{|}~（不包含.）
    # Unicode标点: \u2000-\u206F（通用标点），\u3000-\u303F（CJK符号和标点），\uFF00-\uFFEF（全角标点）
    # 通过在字符类中排除.来实现
    pattern = r'[!"#$%&\'()*+,\-/:;<=>?@\[\]^_`{|}~\u2000-\u206F\u3000-\u303F\uFF00-\uFFEF]'
    delimiters = re.findall(pattern, sample_text)
    
    for delim in delimiters:
        delimiter_counts[delim] = delimiter_counts.get(delim, 0) + 1
    
    # 按出现频率排序，选择出现次数大于等于1的分隔符
    sorted_delims = sorted(delimiter_counts.items(), key=lambda x: x[1], reverse=True)
    selected_delims = [re.escape(delim) for delim, count in sorted_delims if count >= 1]
    
    # 排除点号（.）作为分隔符（防止意外包含）
    if '.' in selected_delims:
        selected_delims.remove(re.escape('.'))
    
    return selected_delims

def read_txt_file(file_path: str) -> List[str]:
    """读取TXT文件并提取词汇列表"""
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # 截取前200字符进行分隔符检测
    sample_text = content[:200]
    delimiters = detect_delimiters(sample_text)
    
    if delimiters:
        # 构建正则分隔符模式，使用字符类
        # 例如，如果 delimiters = [',', ';'], 则 split_pattern = '[,;]'
        # 这里需要去重，防止重复
        unique_delims = sorted(set(delimiters), key=delimiters.index)
        split_pattern = '[' + ''.join(unique_delims) + ']'
        words = re.split(split_pattern, content)
    else:
        # 如果未检测到分隔符，按空白字符分割
        words = re.split(r'\s+', content)
    
    # 去除空字符串和标题
    words = [word.strip() for word in words if word.strip()]
    return words

def read_word_file(file_path: str) -> List[str]:
    """读取Word文件并提取词汇列表"""
    import docx
    doc = docx.Document(file_path)
    content = '\n'.join([para.text for para in doc.paragraphs if para.text.strip()])
    
    # 截取前200字符进行分隔符检测
    sample_text = content[:200]
    delimiters = detect_delimiters(sample_text)
    
    if delimiters:
        # 构建正则分隔符模式，使用字符类
        unique_delims = sorted(set(delimiters), key=delimiters.index)
        split_pattern = '[' + ''.join(unique_delims) + ']'
        words = re.split(split_pattern, content)
    else:
        # 如果未检测到分隔符，按空白字符分割
        words = re.split(r'\s+', content)
    
    # 去除空字符串和标题
    words = [word.strip() for word in words if word.strip()]
    return words

def read_excel_file(file_path: str) -> List[str]:
    """读取Excel文件并提取词汇列表"""
    words = []
    try:
        if file_path.lower().endswith('.xls'):
            df_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str, engine='xlrd')  # 指定引擎
        else:
            df_sheets = pd.read_excel(file_path, sheet_name=None, dtype=str, engine='openpyxl')  # 指定引擎
    except ImportError:
        raise ImportError("缺少必要的依赖库。请确保已安装 'xlrd' 和 'openpyxl'。")
    except Exception as e:
        raise ValueError(f"读取Excel文件时出错：{e}")

    for sheet_name, df in df_sheets.items():
        max_unique = 0
        vocab_column = None
        for column in df.columns:
            unique_values = df[column].dropna().unique()
            num_unique = len(unique_values)
            if num_unique > max_unique:
                max_unique = num_unique
                vocab_column = column

        if vocab_column:
            column_words = df[vocab_column].dropna().astype(str).tolist()
            words.extend(column_words)

    # 去除空字符串和标题（假设标题不在词汇中）
    words = [word.strip() for word in words if word.strip()]
    return words

def extract_words(file_path: str) -> List[str]:
    """根据文件类型提取词汇列表"""
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    if ext == '.txt':
        return read_txt_file(file_path)
    elif ext == '.docx':
        return read_word_file(file_path)
    elif ext in ['.xlsx', '.xls']:
        return read_excel_file(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {ext}")

def compare_vocabularies(a_file: str, b_file: str) -> Dict[str, Set[str]]:
    """比较两个词表，返回对照结果"""
    a_words = set(extract_words(a_file))
    b_words = set(extract_words(b_file))

    merged_words = a_words.union(b_words)
    a_missing_in_b = b_words - a_words
    b_missing_in_a = a_words - b_words

    return {
        'a_words': a_words,
        'b_words': b_words,
        'merged_words': merged_words,
        'a_missing_in_b': a_missing_in_b,
        'b_missing_in_a': b_missing_in_a
    }

def write_to_excel(results: Dict[str, Set[str]], output_path: str):
    """将对照结果写入Excel文件"""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet1: 合并词表
        if results['merged_words']:
            df_merged = pd.DataFrame(sorted(results['merged_words']), columns=['合并词表'])
            df_merged.to_excel(writer, sheet_name='合并词表', index=False)
        else:
            df_empty = pd.DataFrame({'信息': ['合并词表为空。']})
            df_empty.to_excel(writer, sheet_name='合并词表', index=False)

        # Sheet2: A缺失的B
        if results['a_missing_in_b']:
            df_a_missing = pd.DataFrame(sorted(results['a_missing_in_b']), columns=['A词表缺失的B词汇'])
            df_a_missing.to_excel(writer, sheet_name='A缺失的B', index=False)
        else:
            df_empty = pd.DataFrame({'信息': ['A词表未缺失B词表中的词汇。']})
            df_empty.to_excel(writer, sheet_name='A缺失的B', index=False)

        # Sheet3: B缺失的A
        if results['b_missing_in_a']:
            df_b_missing = pd.DataFrame(sorted(results['b_missing_in_a']), columns=['B词表缺失的A词汇'])
            df_b_missing.to_excel(writer, sheet_name='B缺失的A', index=False)
        else:
            df_empty = pd.DataFrame({'信息': ['B词表未缺失A词表中的词汇。']})
            df_empty.to_excel(writer, sheet_name='B缺失的A', index=False)

class VocabularyComparisonProcessor(QObject):
    """词表对照处理器"""
    finished = Signal(str, int, int, int, int, int)  # output_path, counts...

    def __init__(self, a_file: str, b_file: str, output_dir: str):
        super().__init__()
        self.a_file = a_file
        self.b_file = b_file
        self.output_dir = output_dir

    def run(self):
        try:
            results = compare_vocabularies(self.a_file, self.b_file)
            # 获取当前日期
            today_str = datetime.today().strftime('%Y-%m-%d')
            output_filename = f"词表对照结果_{today_str}.xlsx"
            output_path = os.path.join(self.output_dir, output_filename)
            write_to_excel(results, output_path)
            # 发射信号，包含输出路径和统计数据
            self.finished.emit(
                output_path,
                len(results['a_words']),
                len(results['b_words']),
                len(results['merged_words']),
                len(results['a_missing_in_b']),
                len(results['b_missing_in_a'])
            )
        except PermissionError:
            error_message = f"错误: 无法写入文件 '{output_path}'。请关闭该文件后重试。"
            self.finished.emit(error_message, 0, 0, 0, 0, 0)
        except Exception as e:
            self.finished.emit(f"错误: {str(e)}", 0, 0, 0, 0, 0)
