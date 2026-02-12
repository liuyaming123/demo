import os
import uuid
import json
from typing import Dict, Any, List

from flask import Flask, render_template, request, jsonify
import pandas as pd
import dashscope
from dashscope import Generation


app = Flask(__name__, template_folder="templates", static_folder="static")


# 简单的内存态存储：file_id -> 文件路径
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
FILES: Dict[str, str] = {}


def allowed_file(filename: str) -> bool:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return ext in {"xls", "xlsx"}


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/upload")
def upload():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "未找到文件字段"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"ok": False, "error": "文件名为空"}), 400

    if not allowed_file(f.filename):
        return jsonify({"ok": False, "error": "只支持 .xls / .xlsx 文件"}), 400

    file_id = str(uuid.uuid4())
    save_path = os.path.join(UPLOAD_DIR, f"{file_id}_{f.filename}")
    f.save(save_path)
    FILES[file_id] = save_path

    try:
        # 预览：每个 sheet 前 5 行，最多 50 列
        sheets_dict = pd.read_excel(save_path, sheet_name=None, header=None, nrows=5)
    except Exception as e:
        return jsonify({"ok": False, "error": f"读取 Excel 失败: {e}"}), 500

    sheets_preview: List[Dict[str, Any]] = []
    for sheet_name, df in sheets_dict.items():
        # 限制列数
        df_limited = df.iloc[:, :50]
        # 使用 Python 原生类型，方便前端 JSON 处理
        data = df_limited.where(pd.notnull(df_limited), "").values.tolist()
        col_count = df_limited.shape[1]
        row_count = df_limited.shape[0]

        sheets_preview.append(
            {
                "sheet_name": str(sheet_name),
                "rows": data,
                "row_count": int(row_count),
                "col_count": int(col_count),
            }
        )

    return jsonify({"ok": True, "file_id": file_id, "sheets": sheets_preview})


def call_ai_for_sheet_preview(sheet_name: str, rows: List[List[Any]]) -> Dict[str, Any]:
    """
    调用大模型，根据前 5 行预览推断：
    - columns: 列名列表（按顺序）
    - data_start_row: 真正数据起始行号（从 1 开始，针对整张 sheet）

    为了方便解析，这里约定模型返回一个 JSON：
    {
      "columns": ["col1", "col2", ...],
      "data_start_row": 3
    }
    """
    # 使用 DashScope 官方 SDK + qwen3-max
    dashscope.api_key = os.getenv("DASHSCOPE_API_KEY")
    if not dashscope.api_key:
        raise RuntimeError("环境变量 DASHSCOPE_API_KEY 未设置")

    # 将预览数据序列化为文本，帮助模型理解
    lines = []
    for idx, row in enumerate(rows, start=1):
        # 将空值统一成空字符串
        row_str = ["" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v) for v in row]
        lines.append(f"第 {idx} 行: {row_str}")
    preview_text = "\n".join(lines)

    system_prompt = """
你是一名擅长理解 Excel 表格结构的数据工程师，任务是根据给定的前几行内容，
推断这一张表真正的数据列名以及真正数据开始的行号。
你必须只输出一个 JSON 对象，形如：
{"columns": ["列名1", "列名2"], "data_start_row": 3}
其中：
- columns: 按列顺序给出，使用字符串数组；
- data_start_row: 整张 sheet 中真正数据（每一行代表一条记录）开始的行号，从 1 开始计数。
不要输出任何解释性文字。
    """.strip()

    user_prompt = f"""
下面是 Excel 中某个 sheet（名称：{sheet_name}）的前 5 行、最多 50 列的内容预览：
{preview_text}

请根据这些内容判断：
1. 这一张表中真正的列名（用于后续转成 JSON 的字段名）是什么？按顺序给出。
2. 真正数据开始的行号是几？（从 1 开始，针对此整张 sheet 的行号）。
如果前几行是表头或说明文字，请综合判断最合理的数据起始行。

输出格式为 JSON，形如：
{{"columns": ["列名1", "列名2"], "data_start_row": 3}}
    """.strip()

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    # 非流式调用 qwen3-max，显式使用 result_format='text'
    response = Generation.call(
        model="qwen3-max",
        messages=messages,
        temperature=0.1,
    )

    if response.status_code != 200:
        raise RuntimeError(f"DashScope 调用失败: {response.code} {response.message}")

    # 优先从 output.text 读取 JSON 字符串，兼容 result_format='message' 时从 choices[0].message.content 读取
    raw_content = response.output.get("text")
    if not raw_content:
        choices = response.output.get("choices") or []
        if choices:
            # Choice 和 Message 都是 DictMixin，可按字典方式访问
            raw_content = choices[0].get("message", {}).get("content")

    if not raw_content:
        raise RuntimeError("模型返回内容为空，无法解析列信息。")

    # 去掉可能的 Markdown 代码块包装 ```json ... ```
    raw_content = str(raw_content).strip()
    if raw_content.startswith("```"):
        # 可能是 ```json 或 ``` 包裹
        lines = raw_content.splitlines()
        # 去掉第一行和最后一行 ```xxx / ```
        if len(lines) >= 2 and lines[0].startswith("```") and lines[-1].startswith("```"):
            raw_content = "\n".join(lines[1:-1]).strip()

    parsed = json.loads(raw_content)
    columns = parsed.get("columns") or []
    data_start_row = parsed.get("data_start_row")

    # 兜底处理
    if not isinstance(columns, list):
        columns = []
    columns = [str(c) for c in columns]
    try:
        data_start_row = int(data_start_row)
    except Exception:
        data_start_row = 1

    return {"columns": columns, "data_start_row": data_start_row}


@app.post("/analyze")
def analyze():
    payload = request.get_json(silent=True) or {}
    file_id = payload.get("file_id")
    sheets_req = payload.get("sheets") or []

    if not file_id or file_id not in FILES:
        return jsonify({"ok": False, "error": "file_id 无效，请重新上传文件"}), 400

    # 为了严格满足“把前 5 行、最多 50 列内容交给模型”的要求，
    # 这里使用前端传来的 rows（已经是前 5 行、最多 50 列）
    results = {}
    errors = {}
    for sheet in sheets_req:
        name = sheet.get("sheet_name")
        rows = sheet.get("rows") or []
        try:
            result = call_ai_for_sheet_preview(name, rows)
            results[name] = result
        except Exception as e:
            errors[name] = str(e)

    return jsonify({"ok": True, "results": results, "errors": errors})


@app.post("/convert")
def convert():
    payload = request.get_json(silent=True) or {}
    file_id = payload.get("file_id")
    sheets_cfg = payload.get("sheets") or []

    if not file_id or file_id not in FILES:
        return jsonify({"ok": False, "error": "file_id 无效，请重新上传文件"}), 400

    excel_path = FILES[file_id]
    try:
        # 读取所有 sheet，后面按需取用
        all_sheets = pd.read_excel(excel_path, sheet_name=None, header=None)
    except Exception as e:
        return jsonify({"ok": False, "error": f"读取 Excel 失败: {e}"}), 500

    converted: Dict[str, Any] = {}
    errors: Dict[str, str] = {}

    for sheet in sheets_cfg:
        name = sheet.get("sheet_name")
        columns_raw = sheet.get("columns")
        data_start_row = sheet.get("data_start_row")

        if name not in all_sheets:
            errors[name] = "该 sheet 在文件中不存在"
            continue

        # columns 可能是字符串，也可能已经是数组
        if isinstance(columns_raw, str):
            # 尝试按逗号拆分
            columns = [c.strip() for c in columns_raw.strip("[]").split(",") if c.strip()]
        elif isinstance(columns_raw, list):
            columns = [str(c) for c in columns_raw]
        else:
            columns = []

        try:
            data_start_row = int(data_start_row)
        except Exception:
            data_start_row = 1

        try:
            df = all_sheets[name]
            # 真正数据开始行：从 1 开始计数，DataFrame 是 0 基
            start_idx = max(data_start_row - 1, 0)
            df_data = df.iloc[start_idx:, : len(columns)]

            # 将 NaN 转为空字符串
            df_data = df_data.where(pd.notnull(df_data), "")

            records: List[Dict[str, Any]] = []
            for _, row in df_data.iterrows():
                values = row.tolist()
                rec = {col: values[i] if i < len(values) else "" for i, col in enumerate(columns)}
                # 如果整行都是空，可以选择跳过；这里保留，让用户自己决定
                records.append(rec)

            converted[name] = records
        except Exception as e:
            errors[name] = str(e)

    return jsonify({"ok": True, "converted": converted, "errors": errors})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8001)))

