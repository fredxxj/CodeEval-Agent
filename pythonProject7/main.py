# -*- coding: utf-8 -*-
import json
import os
import re
import sys

import pandas as pd
from generate import generate
# 定义解析函数
def parse_scores_from_dict(data):
    """
    从解析后的字典中提取评分和理由
    :param data: 包含学号、姓名和结果的字典
    :return: 包含提取信息的结构化字典
    """


    # 初始化数据结构
    data_structure = {
        "学号": data.get("学号", ""),
        "姓名": data.get("姓名", ""),
        "功能性得分": 0,
        "功能性理由": "",
        "鲁棒性得分": 0,
        "鲁棒性理由": "",
        "效率性得分": 0,
        "效率性理由": "",
        "可维护性得分": 0,
        "可维护性理由": "",
        "总分": 0,
        "改进建议": "",
        "解析状态": "未解析",
        "原始结果": data.get("结果", "")
    }

    try:
        result_text = data.get("结果", "")  # 获取结果字符串
        try:
            result_json = json.loads(result_text)
        except json.JSONDecodeError:
            result_json = None

        if isinstance(result_json, dict):
            for key in [
                "功能性得分",
                "功能性理由",
                "鲁棒性得分",
                "鲁棒性理由",
                "效率性得分",
                "效率性理由",
                "可维护性得分",
                "可维护性理由",
                "总分",
                "改进建议",
            ]:
                if key in result_json:
                    data_structure[key] = result_json[key]
            data_structure["解析状态"] = "成功"
            return data_structure

        # 提取功能性得分和理由
        func_score_text = result_text.split("功能性得分：[")[1]
        func_score = re.sub(r"[^0-9]", "", func_score_text.split("]")[0])  # 移除非数字字符
        data_structure["功能性得分"] = int(func_score)
        data_structure["功能性理由"] = func_score_text.split("功能性理由：")[1].split("鲁棒性得分：")[0].strip()

        # 提取鲁棒性得分和理由
        robust_score_text = result_text.split("鲁棒性得分：[")[1]
        robust_score = re.sub(r"[^0-9]", "", robust_score_text.split("]")[0])  # 移除非数字字符
        data_structure["鲁棒性得分"] = int(robust_score)
        data_structure["鲁棒性理由"] = robust_score_text.split("鲁棒性理由：")[1].split("效率性得分：")[0].strip()

        # 提取效率性得分和理由
        efficiency_score_text = result_text.split("效率性得分：[")[1]
        efficiency_score = re.sub(r"[^0-9]", "", efficiency_score_text.split("]")[0])  # 移除非数字字符
        data_structure["效率性得分"] = int(efficiency_score)
        data_structure["效率性理由"] = efficiency_score_text.split("效率性理由：")[1].split("可维护性得分：")[0].strip()

        # 提取可维护性得分和理由
        maintain_score_text = result_text.split("可维护性得分：[")[1]
        maintain_score = re.sub(r"[^0-9]", "", maintain_score_text.split("]")[0])  # 移除非数字字符
        data_structure["可维护性得分"] = int(maintain_score)
        data_structure["可维护性理由"] = maintain_score_text.split("可维护性理由：")[1].split("总分：")[0].strip()

        # 提取总分
        total_score_text = result_text.split("总分：[")[1]
        total_score = re.sub(r"[^0-9]", "", total_score_text.split("]")[0])  # 移除非数字字符
        data_structure["总分"] = int(total_score)
        data_structure["解析状态"] = "成功"

    except (IndexError, ValueError) as e:
        print(f"解析字符串时出错: {e}")
        if data_structure["原始结果"].startswith(("接口返回错误", "智能体调用失败")):
            data_structure["解析状态"] = "接口错误"
        else:
            data_structure["解析状态"] = "格式不匹配"

    return data_structure


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else os.getenv("GRADING_INPUT_DIR")
    output_file = sys.argv[2] if len(sys.argv) > 2 else os.getenv("GRADING_OUTPUT_FILE", "评分结果.xlsx")

    # 多组输入文本
    input_texts = generate(input_dir)
    parsed_texts = [json.loads(text) for text in input_texts]

    # 使用解析函数提取评分
    parsed_data_list = [parse_scores_from_dict(data) for data in parsed_texts]

    # 转换为DataFrame
    df = pd.DataFrame(parsed_data_list)

    # 保存为Excel文件
    df.to_excel(output_file, index=False)
    print(f"Excel 文件已保存到: {output_file}")


if __name__ == "__main__":
    main()
