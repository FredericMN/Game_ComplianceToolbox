import os
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

def process_file(file_path, filename, search_text):
    """
    处理单个Excel文件，检查是否包含指定的文字片段。

    :param file_path: 文件的完整路径
    :param filename: 文件名称
    :param search_text: 要搜索的文字片段
    :return: (filename, contains_text)
    """
    print(f"[{threading.current_thread().name}] 正在处理文件: {filename}")
    contains_text = False
    try:
        excel_file = pd.ExcelFile(file_path)
        # 遍历所有工作表
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet, dtype=str)
            df = df.fillna('')
            if df.apply(lambda row: row.astype(str).str.contains(search_text, case=False, na=False).any(), axis=1).any():
                contains_text = True
                break  # 找到匹配项后，跳过该文件的其他工作表
    except Exception as e:
        print(f"[{threading.current_thread().name}] 处理文件 {filename} 时出错: {e}")
    return (filename, contains_text)

def search_in_excel_multithread(directory, search_text, max_workers=5):
    """
    在指定目录下使用多线程搜索所有Excel文件，查找是否有单元格包含指定的文字片段。

    :param directory: 要搜索的目录路径
    :param search_text: 要搜索的文字片段
    :param max_workers: 最大线程数
    :return: 包含匹配文字的文件名列表
    """
    matching_files = []
    excel_files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # 提交所有任务
        future_to_file = {executor.submit(process_file, os.path.join(directory, file), file, search_text): file for file in excel_files}
        
        for future in as_completed(future_to_file):
            file = future_to_file[future]
            try:
                filename, contains_text = future.result()
                if contains_text:
                    print(f"[{threading.current_thread().name}] 包含指定文字片段: '{search_text}'")
                    matching_files.append(filename)
                else:
                    print(f"[{threading.current_thread().name}] 不包含指定文字片段: '{search_text}'")
            except Exception as e:
                print(f"[{threading.current_thread().name}] 处理文件 {file} 时出错: {e}")

    return matching_files

def main():
    print("=== Excel 文件内容搜索工具（多线程版） ===")
    # 获取用户输入的目录路径
    directory = input("请输入要搜索的目录路径: ").strip()
    # 检查目录是否存在
    if not os.path.isdir(directory):
        print("目录不存在，请检查路径是否正确。")
        return

    # 获取用户输入的搜索文字
    search_text = input("请输入要搜索的文字片段: ").strip()
    if not search_text:
        print("搜索文字不能为空。")
        return

    # 获取用户输入的最大线程数（可选）
    try:
        max_workers_input = input("请输入最大线程数（默认5）: ").strip()
        max_workers = int(max_workers_input) if max_workers_input else 5
        if max_workers < 1:
            raise ValueError
    except ValueError:
        print("无效的线程数输入，使用默认值5。")
        max_workers = 5

    print("\n开始搜索...\n")

    # 执行搜索
    results = search_in_excel_multithread(directory, search_text, max_workers=max_workers)

    # 输出最终结果
    print("\n=== 搜索完毕 ===")
    if results:
        print("\n以下文件包含所输入的文字片段:")
        for file in results:
            print(f"- {file}")
    else:
        print("\n没有文件包含所输入的文字片段。")

if __name__ == "__main__":
    main()
