# -*- coding: utf-8 -*-
import os
import re
import json
import time
import pandas as pd
from docx import Document
from pathlib import Path

from grading_agent import grade_programming_assignment

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_DIR = PROJECT_ROOT / "大模型自动批改"

def extract_word_files_from_folder(folder_path):
    """
    从指定路径的所有子文件夹中提取 .docx 文件
    :param folder_path: 文件夹路径
    :return: 所有 Word 文件的路径列表
    """
    word_files = []
    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".docx"):
                    word_files.append(os.path.join(root, file))
        return word_files
    except Exception as e:
        print(f"提取 Word 文件时出错：{e}")
        return []

def extract_student_info_from_filename(file_name):
    """
    从文件名中提取学号和姓名
    :param file_name: Word 文件名，格式为 姓名_学号_其他内容.docx
    :return: 学号和姓名
    """
    base_name = os.path.basename(file_name)  # 获取文件名
    student_info = base_name.split('.')[0]  # 去掉文件扩展名
    print(student_info)

    student_name = student_info.split('_')[0]
    student_id = student_info.split('_')[1]
    student_name = re.sub(r"^\d+", "", student_name) or student_name
    id_match = re.match(r"[A-Za-z0-9]+", student_id)
    if id_match:
        student_id = id_match.group(0)
    return student_id, student_name

def extract_word_content(file_path):
    """
    读取 Word 文件内容
    :param file_path: Word 文件路径
    :return: Word 文件的内容字符串
    """
    try:
        document = Document(file_path)
        return "\n".join([paragraph.text for paragraph in document.paragraphs])
    except Exception as e:
        print(f"读取 Word 文件出错：{e}")
        return None

def save_results_to_excel(results, output_path):
    """
    将结果保存为 Excel 文件
    :param results: 学生信息和评分结果列表
    :param output_path: 输出 Excel 文件路径
    """
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)
    print(f"结果已保存到 Excel 文件：{output_path}")

def generate(folder_path):
    folder_path = folder_path or os.getenv("GRADING_INPUT_DIR") or str(DEFAULT_INPUT_DIR)
    if not folder_path:
        raise ValueError("请传入作业文件夹路径，或设置 GRADING_INPUT_DIR 环境变量。")

    # 提取文件夹中的所有 .docx 文件路径
    word_files = extract_word_files_from_folder(folder_path)
    if not word_files:
        print("未找到任何 Word 文件。")
        exit()

    # Process each Word file
    student_info_list = []
    for word_file in word_files:
        print(f"\nProcessing file: {word_file}")

        # 提取学号和姓名
        student_id, student_name = extract_student_info_from_filename(word_file)

        # 读取 Word 文件内容
        word_content = extract_word_content(word_file)
        if not word_content:
            print(f"Skipping this file due to read failure: {word_file}")
            continue

        # 调用本地批改智能体处理 Word 内容
        try:
            final_response = json.dumps(
                grade_programming_assignment(word_content),
                ensure_ascii=False
            )
        except Exception as e:
            final_response = f"智能体调用失败：{e}"

        student_info = {
            "学号": student_id,
            "姓名": student_name,
            "结果": final_response
        }
        student_info_string = json.dumps(student_info, ensure_ascii=False)  # 保留中文字符

        # Append result to the list
        student_info_list.append(student_info_string)
        interval = float(os.getenv("BIGMODEL_REQUEST_INTERVAL", "6"))
        if interval > 0:
            time.sleep(interval)

    return student_info_list

