import os
import json

import dashscope
from dashscope import Generation


def clean_json_str(raw: str) -> str:
    """
    去掉模型可能包裹的 ```json ... ``` 代码块，只保留内部 JSON 字符串。
    """
    s = str(raw).strip()
    if s.startswith("```"):
        lines = s.splitlines()
        if len(lines) >= 2 and lines[0].startswith("```") and lines[-1].startswith("```"):
            s = "\n".join(lines[1:-1]).strip()
    return s


def main() -> None:
    # 从环境变量中读取 DashScope API Key
    dashscope.api_key = os.getenv("DASHSCOPE_API_KEY")
    if not dashscope.api_key:
        raise RuntimeError("请先设置环境变量 DASHSCOPE_API_KEY 再运行本脚本。")

    # 与 app.py 中含义一致的系统提示词和用户提示词（简化版本）
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

    # 这里构造一个非常简单的“虚拟预览数据”，方便离线测试模型行为
    fake_preview = """第 1 行: ['姓名', '年龄', '城市']
第 2 行: ['张三', '18', '北京']
第 3 行: ['李四', '20', '上海']
"""

    user_prompt = f"""
下面是 Excel 中某个 sheet（名称：测试Sheet）的前几行内容预览：
{fake_preview}

请根据这些内容判断：
1. 这一张表中真正的列名（用于后续转成 JSON 的字段名）是什么？按顺序给出。
2. 真正数据开始的行号是几？（从 1 开始，针对此整张 sheet 的行号）。
如果前几行是表头或说明文字，请综合判断最合理的数据起始行。

输出格式为 JSON，形如：
{{"columns": ["姓名", "年龄", "城市"], "data_start_row": 2}}
    """.strip()

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    # 注意：根据你返回的错误信息，qwen3-max 不允许 result_format 为非 "message" 的值，
    # 且推荐直接使用默认（即 message）。这里显式设置为 "message" 便于查看结构。
    print(">>> 正在调用 DashScope Generation.call(model='qwen3-max', result_format='message') ...")
    resp = Generation.call(
        model="qwen3-max",
        messages=messages,
        result_format="message",
        temperature=0.1,
    )

    print("\n=== 原始 Response 对象（简略） ===")
    print(f"status_code = {resp.status_code}")
    print(f"code        = {resp.code}")
    print(f"message     = {resp.message}")

    if resp.status_code != 200:
        print("调用失败，完整响应：")
        print(resp)
        return

    # 对于 result_format='message'，内容一般位于 output.choices[0].message.content
    choices = resp.output.get("choices") or []
    if not choices:
        print("未在 output.choices 中找到内容：", resp.output)
        return

    first_choice = choices[0]
    message_obj = first_choice.get("message") or {}
    content = message_obj.get("content")

    print("\n=== message.content 原始文本 ===")
    print(content)

    # 清理并尝试解析 JSON
    cleaned = clean_json_str(content)
    print("\n=== 清理后的待解析 JSON 文本 ===")
    print(cleaned)

    try:
        parsed = json.loads(cleaned)
    except Exception as e:
        print("\n!!! JSON 解析失败：", e)
        return

    print("\n=== 解析后的 JSON 对象 ===")
    print(parsed)

    columns = parsed.get("columns")
    data_start_row = parsed.get("data_start_row")
    print("\n>>> 提取到的值：")
    print("columns       =", columns)
    print("data_start_row =", data_start_row)


if __name__ == "__main__":
    main()

