# 编程作业评估智能体

面向程序设计课程的批量作业批改工具。系统读取学生 Word 答题文件，调用国产大模型智谱 BigModel，按功能性、鲁棒性、效率性、可维护性四个维度生成评分、理由和改进建议，并导出 Excel 结果。

## 目录结构

```text
pythonProject7/        Python 批改工作流
前端/                  前后端协同演示界面
```

## Python 批改工作流

安装依赖：

```bash
cd pythonProject7
python3 -m pip install -r requirements.txt
```

配置环境变量：

```bash
export BIGMODEL_API_KEY="你的智谱 BigModel API Key"
export GRADING_INPUT_DIR="/path/to/student/docx/folder"
export GRADING_OUTPUT_FILE="评分结果.xlsx"
```

运行批改：

```bash
python3 main.py "$GRADING_INPUT_DIR" "$GRADING_OUTPUT_FILE"
```

续跑失败项：

```bash
python3 resume_failed.py "$GRADING_INPUT_DIR" "$GRADING_OUTPUT_FILE"
```

## 前后端协同演示

```bash
cd pythonProject7
export BIGMODEL_API_KEY="你的智谱 BigModel API Key"
python3 server.py
```

浏览器访问：

```text
http://localhost:8080
```

## 数据与密钥

仓库不包含学生作业、评分结果、API Key、论文 PDF、压缩包等本地数据或敏感文件。请通过环境变量配置 API Key，并在本地指定学生答题文件夹。
