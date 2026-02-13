# Excel 转 JSON 小工具（Flask + 原生前端）

这是一个在本地运行的轻量级网页工具，支持上传 Excel 文件，预览各个 Sheet 的前 5 行，并通过大模型智能分析列名与数据起始行，最后将 Excel 转换为 JSON 并下载。

## 功能概览

- 上传并检测 Excel 文件类型（仅支持 `.xls` / `.xlsx`）。
- 对每个 Sheet 展示前 5 行、最多 50 列的内容预览，并显示行号从 1 开始。
- 对所有 Sheet 提供全选 / 取消全选 / 单个勾选。
- 对勾选的 Sheet 调用大模型（如 `qwen3-max`）进行 **智能分析**：
  - 推断列名列表（按顺序，Python 列表形式）。
  - 推断真正数据起始行号（从 1 开始）。
- 前端可修改智能分析结果（列名与起始行）。
- 支持：
  - 批量转换选中 Sheet；
  - 单独转换某一个 Sheet；
  - 重复转换（配置修改后再次转换）。
- 转换成功后，可预览 JSON（每个 Sheet 单独预览），并以 Sheet 名称作为文件名下载。

## 目录结构

- `app.py`：Flask 后端主程序。
- `templates/index.html`：前端页面。
- `static/style.css`：页面样式（现代深色风格）。
- `static/app.js`：前端逻辑（上传、预览、智能分析、转换与下载）。
- `requirements.txt`：Python 依赖。

## 环境准备

1. 创建并激活虚拟环境（可选但推荐）：

```bash
cd /Users/liu/practice/abc_play/play/test
python -m venv .venv
source .venv/bin/activate  # Windows 使用 .venv\Scripts\activate
```

2. 安装依赖：

```bash
pip install -r requirements.txt
```

3. 配置大模型 API 环境变量（DashScope）：

后端现在通过 **DashScope 官方 Python SDK** 调用 `qwen3-max` 模型，并使用全局：

```python
dashscope.api_key = os.getenv("DASHSCOPE_API_KEY")
```

因此只需要配置一个环境变量：

- `DASHSCOPE_API_KEY`：必填，你在百炼 / DashScope 控制台获取的 API Key。

示例（在 bash 中）：

```bash
export DASHSCOPE_API_KEY="你的_DashScope_API_Key"
```
> 如果你需要切换为其他 DashScope 模型（如 `qwen-plus` 等），可以在 `app.py` 的 `call_ai_for_sheet_preview` 中调整 `model` 参数。

## 运行方式

```bash
cd test
python app.py
```

默认在 `http://127.0.0.1:5000/` 提供服务。

## 与大模型的交互设计（提示词说明）

后端在 `app.py` 的 `call_ai_for_sheet_preview` 函数中，给模型的指令为（简要说明）：

- **系统提示（system）**：
  - 你是一名擅长理解 Excel 表格结构的数据工程师。
  - 任务是根据给定的前几行内容，推断真正的数据列名以及真正数据开始行号。
  - 要求严格只输出一个 JSON 对象，结构为：
    - `{"columns": ["列名1", "列名2"], "data_start_row": 3}`。
- **用户提示（user）**：
  - 给出某个 Sheet 名称。
  - 给出该 Sheet 前 5 行（最多 50 列）的文本化预览。
  - 要求根据预览判断：
    - 真实的列名列表（用于后续 JSON 字段名）。
    - 真正数据起始行（从 1 开始，针对整张 Sheet 的行号）。

模型返回内容将被解析为 JSON，用于前端展示与后续转换。

## 注意事项

- 智能分析和转换都可以多次重复执行；前端会保留最近一次转换结果，方便对比与导出。
- 对于列名输入区域，你可以：
  - 直接粘贴 Python 列表字面量，例如：`["姓名", "年龄", "城市"]`；
  - 或者简单写成逗号分隔字符串，例如：`姓名, 年龄, 城市`。
- 如果 Excel 非常大，上传时服务器会一次性读取整个文件；但预览和模型调用只使用前 5 行、最多 50 列。

# demo
