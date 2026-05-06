# -*- coding: utf-8 -*-
import json
import os
import re
import time
from typing import Any, Dict

import requests


BIGMODEL_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
DEFAULT_MODEL = "glm-4.7-flash"
DEFAULT_MAX_ASSIGNMENT_CHARS = 12000


SYSTEM_PROMPT = """
你是“编程作业评估智能体”，服务于高校程序设计课程。你的任务是阅读学生提交的编程作业内容，按照课程初学者评价标准进行客观、温和、可解释的批改。

评分维度：
1. 功能性：0-30 分。检查代码是否实现题目要求、主要功能是否完整。
2. 鲁棒性：0-20 分。只评价异常输入、边界条件、越界风险、死循环风险、运行稳定性。
3. 效率性：0-20 分。只评价算法复杂度、资源使用、重复计算、绘制频率和运行流畅度。
4. 可维护性：0-30 分。只评价变量/函数命名、注释、结构划分、可读性和扩展性。

输出要求：
- 只输出一个 JSON 对象，不要使用 Markdown。
- 所有分数字段必须是整数。
- 总分必须等于四个维度得分之和。
- 四个“理由”字段都必须非空，每个理由 40-100 字。
- 每个维度的理由只能讨论该维度，不要把可维护性内容写进鲁棒性理由。
- 理由要具体指出代码表现，不要空泛夸赞。
- 建议要可操作，适合学生根据反馈修改代码。

JSON 字段：
{
  "功能性得分": 0,
  "功能性理由": "",
  "鲁棒性得分": 0,
  "鲁棒性理由": "",
  "效率性得分": 0,
  "效率性理由": "",
  "可维护性得分": 0,
  "可维护性理由": "",
  "总分": 0,
  "改进建议": ""
}
""".strip()


def _extract_json(text: str) -> Dict[str, Any]:
    """Extract a JSON object from a model response."""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.S)
    if fenced:
        return json.loads(fenced.group(1))

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return json.loads(text[start : end + 1])

    raise ValueError("模型返回内容中未找到 JSON 对象。")


def _normalize_score(value: Any, max_score: int) -> int:
    try:
        score = int(float(value))
    except (TypeError, ValueError):
        score = 0
    return max(0, min(max_score, score))


def normalize_grading_result(data: Dict[str, Any]) -> Dict[str, Any]:
    result = {
        "功能性得分": _normalize_score(data.get("功能性得分"), 30),
        "功能性理由": str(data.get("功能性理由", "")).strip(),
        "鲁棒性得分": _normalize_score(data.get("鲁棒性得分"), 20),
        "鲁棒性理由": str(data.get("鲁棒性理由", "")).strip(),
        "效率性得分": _normalize_score(data.get("效率性得分"), 20),
        "效率性理由": str(data.get("效率性理由", "")).strip(),
        "可维护性得分": _normalize_score(data.get("可维护性得分"), 30),
        "可维护性理由": str(data.get("可维护性理由", "")).strip(),
        "改进建议": str(data.get("改进建议", "")).strip(),
    }
    result["总分"] = (
        result["功能性得分"]
        + result["鲁棒性得分"]
        + result["效率性得分"]
        + result["可维护性得分"]
    )
    for key in ["功能性理由", "鲁棒性理由", "效率性理由", "可维护性理由"]:
        if not result[key]:
            result[key] = "模型未返回该维度的详细理由，建议教师复核该项评分。"
    return result


def grade_programming_assignment(assignment_text: str) -> Dict[str, Any]:
    api_key = os.getenv("BIGMODEL_API_KEY") or os.getenv("ZHIPU_API_KEY")
    if not api_key:
        raise ValueError("请先设置 BIGMODEL_API_KEY 或 ZHIPU_API_KEY 环境变量。")

    model = os.getenv("BIGMODEL_MODEL", DEFAULT_MODEL)
    max_chars = int(os.getenv("MAX_ASSIGNMENT_CHARS", DEFAULT_MAX_ASSIGNMENT_CHARS))
    if len(assignment_text) > max_chars:
        assignment_text = assignment_text[:max_chars] + "\n\n[内容较长，后续部分已截断用于自动批改。]"

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": "请批改以下编程作业内容：\n\n" + assignment_text,
            },
        ],
        "temperature": 0.2,
        "top_p": 0.8,
        "stream": False,
        "max_tokens": 1600,
        "response_format": {"type": "json_object"},
        "thinking": {"type": "disabled"},
    }

    max_retries = int(os.getenv("BIGMODEL_MAX_RETRIES", "3"))
    retry_wait = int(os.getenv("BIGMODEL_RETRY_WAIT", "25"))
    body = None
    response = None
    last_error = None

    for attempt in range(max_retries + 1):
        try:
            response = requests.post(
                BIGMODEL_URL,
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload,
                timeout=120,
            )

            try:
                body = response.json()
            except ValueError as exc:
                raise RuntimeError(f"接口返回非 JSON 内容，状态码：{response.status_code}") from exc

            message = body.get("msg") or body.get("message") or body.get("error") or body
            is_rate_limit = str(body.get("code")) == "1302" or "速率限制" in str(message)
            if response.status_code == 200 and not body.get("error") and not body.get("code"):
                break
            if is_rate_limit and attempt < max_retries:
                time.sleep(retry_wait * (attempt + 1))
                continue
            raise RuntimeError(f"BigModel 接口错误：{message}")
        except requests.Timeout as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(retry_wait * (attempt + 1))
                continue
            raise RuntimeError("BigModel 接口超时，请稍后重试或缩短作业内容。") from exc

    if body is None:
        raise RuntimeError(f"BigModel 接口调用失败：{last_error}")

    content = body["choices"][0]["message"]["content"]
    return normalize_grading_result(_extract_json(content))
