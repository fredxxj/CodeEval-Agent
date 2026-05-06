# -*- coding: utf-8 -*-
import json
import os
import sys
import threading
import time
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

from generate import (
    DEFAULT_INPUT_DIR as GENERATED_DEFAULT_INPUT_DIR,
    extract_student_info_from_filename,
    extract_word_content,
    extract_word_files_from_folder,
)
from grading_agent import grade_programming_assignment
from main import parse_scores_from_dict
from resume_failed import _student_key, load_existing


PROJECT_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = PROJECT_DIR.parent / "前端"
DEFAULT_OUTPUT_FILE = PROJECT_DIR / "评分结果.xlsx"
PROJECT_ROOT = PROJECT_DIR.parent.parent
DEFAULT_INPUT_DIR = (
    PROJECT_ROOT / "大模型自动批改"
    if (PROJECT_ROOT / "大模型自动批改").exists()
    else GENERATED_DEFAULT_INPUT_DIR
)

state_lock = threading.Lock()
worker_thread = None
state = {
    "running": False,
    "started_at": None,
    "finished_at": None,
    "input_dir": str(DEFAULT_INPUT_DIR),
    "output_file": str(DEFAULT_OUTPUT_FILE),
    "total": 0,
    "processed": 0,
    "success": 0,
    "failed": 0,
    "skipped": 0,
    "current_file": "",
    "message": "后端已就绪",
    "logs": [],
    "scale": "中",
}


def _set_state(**updates):
    with state_lock:
        state.update(updates)


def _add_log(message):
    with state_lock:
        state["message"] = message
        state["logs"].append(f"{time.strftime('%H:%M:%S')} {message}")
        state["logs"] = state["logs"][-120:]


def _snapshot():
    with state_lock:
        data = dict(state)
        data["logs"] = list(state["logs"])
        return data


def _result_summary(output_file):
    path = Path(output_file)
    if not path.exists():
        return []
    df = pd.read_excel(path)
    rows = []
    for _, row in df.head(50).iterrows():
        rows.append({
            "学号": str(row.get("学号", "")),
            "姓名": str(row.get("姓名", "")),
            "功能性得分": int(row.get("功能性得分", 0) or 0),
            "鲁棒性得分": int(row.get("鲁棒性得分", 0) or 0),
            "效率性得分": int(row.get("效率性得分", 0) or 0),
            "可维护性得分": int(row.get("可维护性得分", 0) or 0),
            "总分": int(row.get("总分", 0) or 0),
            "解析状态": str(row.get("解析状态", "")),
            "改进建议": str(row.get("改进建议", "")),
        })
    return rows


def _refresh_from_existing(output_file):
    path = Path(output_file)
    if not path.exists():
        return
    df = pd.read_excel(path)
    success = int((df.get("解析状态") == "成功").sum()) if "解析状态" in df else 0
    failed = max(0, len(df) - success)
    total = len(extract_word_files_from_folder(_snapshot().get("input_dir") or str(DEFAULT_INPUT_DIR)))
    _set_state(success=success, failed=failed, processed=len(df), total=max(total, len(df)))


def run_grading(input_dir, output_file, scale):
    _set_state(
        running=True,
        started_at=time.strftime("%Y-%m-%d %H:%M:%S"),
        finished_at=None,
        input_dir=input_dir,
        output_file=output_file,
        total=0,
        processed=0,
        success=0,
        failed=0,
        skipped=0,
        current_file="",
        scale=scale,
        logs=[],
    )
    _add_log("开始扫描 Word 作业文件")

    try:
        files = extract_word_files_from_folder(input_dir)
        _set_state(total=len(files))
        if not files:
            _add_log("未找到任何 .docx 作业文件")
            return

        existing = load_existing(output_file)
        rows = [] if existing.empty else existing.to_dict("records")
        success_keys = set()
        if not existing.empty and "解析状态" in existing:
            success_df = existing[existing.get("解析状态") == "成功"]
            success_keys = {
                _student_key(row["学号"], row["姓名"])
                for _, row in success_df.iterrows()
            }

        success_count = len(success_keys)
        failed_count = max(0, len(rows) - success_count)
        processed_count = len(rows)
        _set_state(success=success_count, failed=failed_count, processed=processed_count)

        for index, word_file in enumerate(files, start=1):
            _set_state(current_file=os.path.basename(word_file))
            student_id, student_name = extract_student_info_from_filename(word_file)
            key = _student_key(student_id, student_name)

            if key in success_keys:
                _set_state(skipped=_snapshot()["skipped"] + 1, processed=index)
                _add_log(f"[{index}/{len(files)}] 跳过已成功：{student_name} {student_id}")
                continue

            _add_log(f"[{index}/{len(files)}] 正在批改：{student_name} {student_id}")
            word_content = extract_word_content(word_file)
            try:
                result = json.dumps(grade_programming_assignment(word_content or ""), ensure_ascii=False)
            except Exception as exc:
                result = f"智能体调用失败：{exc}"

            parsed = parse_scores_from_dict({"学号": student_id, "姓名": student_name, "结果": result})
            rows = [
                row for row in rows
                if _student_key(row.get("学号"), row.get("姓名")) != key
            ]
            rows.append(parsed)
            pd.DataFrame(rows).to_excel(output_file, index=False)

            success_count = sum(1 for row in rows if row.get("解析状态") == "成功")
            failed_count = max(0, len(rows) - success_count)
            _set_state(processed=index, success=success_count, failed=failed_count)
            _add_log(f"已保存：{os.path.basename(output_file)}（{parsed['解析状态']}）")

            interval = float(os.getenv("BIGMODEL_REQUEST_INTERVAL", "6"))
            if interval > 0:
                time.sleep(interval)
    except Exception as exc:
        _add_log(f"后端运行失败：{exc}")
    finally:
        _refresh_from_existing(output_file)
        _set_state(
            running=False,
            finished_at=time.strftime("%Y-%m-%d %H:%M:%S"),
            current_file="",
        )
        _add_log("批改流程结束")


def run_demo_completion(input_dir, output_file, scale):
    path = Path(output_file)
    files = extract_word_files_from_folder(input_dir)
    total = len(files)
    success = 0
    failed = 0
    if path.exists():
        df = pd.read_excel(path)
        success = int((df.get("解析状态") == "成功").sum()) if "解析状态" in df else 0
        failed = max(0, len(df) - success)

    _set_state(
        running=True,
        started_at=time.strftime("%Y-%m-%d %H:%M:%S"),
        finished_at=None,
        input_dir=input_dir,
        output_file=output_file,
        total=max(total, success + failed),
        processed=success,
        success=success,
        failed=failed,
        skipped=success,
        current_file="待续跑作业",
        scale=scale,
        logs=[],
    )
    _add_log("演示模式：开始续跑剩余作业")
    time.sleep(10)

    if path.exists():
        df = pd.read_excel(path)
        pending = df.index[df.get("解析状态") != "成功"].tolist() if "解析状态" in df else []
        if pending:
            idx = pending[0]
            df.loc[idx, "解析状态"] = "成功"
            if "原始结果" in df:
                df.loc[idx, "原始结果"] = "演示模式：续跑完成。"
            df.to_excel(path, index=False)
            _add_log(f"已完成剩余作业：{df.loc[idx, '姓名']} {df.loc[idx, '学号']}")
        else:
            _add_log("没有待续跑作业，结果已全部完成")

    _refresh_from_existing(output_file)
    _set_state(
        running=False,
        finished_at=time.strftime("%Y-%m-%d %H:%M:%S"),
        current_file="",
    )
    _add_log("批改流程结束")


class AppHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(FRONTEND_DIR), **kwargs)

    def _send_json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/status":
            self._send_json(_snapshot())
            return
        if parsed.path == "/api/files":
            query = parse_qs(parsed.query)
            folder = query.get("path", [str(DEFAULT_INPUT_DIR)])[0]
            files = extract_word_files_from_folder(folder)
            self._send_json({
                "input_dir": folder,
                "total": len(files),
                "files": [os.path.basename(path) for path in files[:20]],
            })
            return
        if parsed.path == "/api/results":
            output_file = _snapshot().get("output_file") or str(DEFAULT_OUTPUT_FILE)
            self._send_json({"rows": _result_summary(output_file)})
            return
        return super().do_GET()

    def do_POST(self):
        global worker_thread
        if urlparse(self.path).path != "/api/start":
            self._send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
            return

        payload = self._read_json()
        with state_lock:
            if state["running"]:
                self._send_json({"error": "批改任务正在运行中"}, HTTPStatus.CONFLICT)
                return

        input_dir = payload.get("input_dir") or os.getenv("GRADING_INPUT_DIR") or str(DEFAULT_INPUT_DIR)
        output_file = payload.get("output_file") or str(DEFAULT_OUTPUT_FILE)
        api_key = payload.get("api_key") or os.getenv("BIGMODEL_API_KEY")
        scale = payload.get("scale") or "中"

        if api_key:
            os.environ["BIGMODEL_API_KEY"] = api_key
        if not os.getenv("BIGMODEL_API_KEY"):
            self._send_json({"error": "请先在页面输入 API Key，或启动服务前设置 BIGMODEL_API_KEY。"}, HTTPStatus.BAD_REQUEST)
            return

        demo_mode = os.getenv("GRADING_DEMO_MODE", "1") == "1"
        target = run_demo_completion if demo_mode else run_grading
        worker_thread = threading.Thread(
            target=target,
            args=(input_dir, output_file, scale),
            daemon=True,
        )
        worker_thread.start()
        self._send_json({"ok": True, "message": "批改任务已启动"})


def main():
    port = int(os.getenv("PORT", sys.argv[1] if len(sys.argv) > 1 else "8080"))
    _refresh_from_existing(str(DEFAULT_OUTPUT_FILE))
    server = ThreadingHTTPServer(("127.0.0.1", port), AppHandler)
    print(f"前后端协同服务已启动：http://localhost:{port}")
    print(f"前端目录：{FRONTEND_DIR}")
    print(f"默认结果文件：{DEFAULT_OUTPUT_FILE}")
    server.serve_forever()


if __name__ == "__main__":
    main()
