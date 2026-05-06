# -*- coding: utf-8 -*-
import json
import os
import sys
import time
from pathlib import Path

import pandas as pd

from generate import (
    DEFAULT_INPUT_DIR,
    extract_student_info_from_filename,
    extract_word_content,
    extract_word_files_from_folder,
)
from grading_agent import grade_programming_assignment
from main import parse_scores_from_dict


def _student_key(student_id, student_name):
    return f"{student_id}|{student_name}"


def load_existing(output_file):
    if not Path(output_file).exists():
        return pd.DataFrame()
    return pd.read_excel(output_file)


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else os.getenv("GRADING_INPUT_DIR") or str(DEFAULT_INPUT_DIR)
    output_file = sys.argv[2] if len(sys.argv) > 2 else os.getenv("GRADING_OUTPUT_FILE", "评分结果.xlsx")

    existing = load_existing(output_file)
    rows = [] if existing.empty else existing.to_dict("records")
    success_keys = set()
    if not existing.empty:
        success = existing[existing.get("解析状态") == "成功"]
        success_keys = {
            _student_key(row["学号"], row["姓名"])
            for _, row in success.iterrows()
        }

    files = extract_word_files_from_folder(input_dir)
    for index, word_file in enumerate(files, start=1):
        student_id, student_name = extract_student_info_from_filename(word_file)
        key = _student_key(student_id, student_name)
        if key in success_keys:
            print(f"[{index}/{len(files)}] 跳过已成功：{student_name} {student_id}", flush=True)
            continue

        print(f"[{index}/{len(files)}] 批改：{student_name} {student_id}", flush=True)
        word_content = extract_word_content(word_file)
        try:
            result = json.dumps(grade_programming_assignment(word_content), ensure_ascii=False)
        except Exception as exc:
            result = f"智能体调用失败：{exc}"

        parsed = parse_scores_from_dict({"学号": student_id, "姓名": student_name, "结果": result})
        rows = [
            row for row in rows
            if _student_key(row.get("学号"), row.get("姓名")) != key
        ]
        rows.append(parsed)
        pd.DataFrame(rows).to_excel(output_file, index=False)
        print(f"已保存：{output_file}（{parsed['解析状态']}）", flush=True)
        interval = float(os.getenv("BIGMODEL_REQUEST_INTERVAL", "6"))
        if interval > 0:
            time.sleep(interval)


if __name__ == "__main__":
    main()
