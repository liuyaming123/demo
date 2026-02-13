// 全局状态
let currentFileId = null;
let sheetsState = {}; // sheetName -> { selected, analysis, columns, dataStartRow, json }

const $ = (id) => document.getElementById(id);

function setStatus(el, msg, type = "") {
  el.textContent = msg || "";
  el.className = "status-text" + (type ? " " + type : "");
}

function updateAnalyzeButtonState() {
  const anySelected = Object.values(sheetsState).some((s) => s.selected);
  $("analyzeBtn").disabled = !anySelected;
}

function updateConvertButtonsState() {
  const anySelected = Object.values(sheetsState).some((s) => s.selected);
  $("convertSelectedBtn").disabled = !anySelected;
}

function renderSheets() {
  const container = $("sheetsContainer");
  container.innerHTML = "";

  const names = Object.keys(sheetsState);
  if (!names.length) return;

  names.forEach((name) => {
    const sheet = sheetsState[name];
    const card = document.createElement("div");
    card.className = "sheet-card";

    const header = document.createElement("div");
    header.className = "sheet-header";
    header.innerHTML = `
      <div class="sheet-title">
        <label>
          <input type="checkbox" class="sheet-select" data-name="${name}" ${sheet.selected ? "checked" : ""} />
          <span>选择</span>
        </label>
        <span class="sheet-name-pill">${name}</span>
      </div>
      <div class="sheet-meta">
        共约 ${sheet.rowCount} 行 · ${sheet.colCount} 列（预览前 5 行，最多 50 列）
      </div>
    `;
    card.appendChild(header);

    const body = document.createElement("div");
    body.className = "sheet-body";

    // 预览表格
    const previewWrapper = document.createElement("div");
    previewWrapper.className = "sheet-preview-wrapper";
    const table = document.createElement("table");
    table.className = "sheet-preview-table";

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    // 行号列
    const thIndex = document.createElement("th");
    thIndex.textContent = "#";
    thIndex.className = "row-index-cell";
    headRow.appendChild(thIndex);
    // 数据列：这里只用索引展示
    if (sheet.rows.length) {
      const colLen = sheet.rows[0].length;
      for (let i = 0; i < colLen; i++) {
        const th = document.createElement("th");
        th.textContent = `列 ${i + 1}`;
        headRow.appendChild(th);
      }
    }
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    sheet.rows.forEach((row, idx) => {
      const tr = document.createElement("tr");
      const tdIndex = document.createElement("td");
      tdIndex.textContent = idx + 1;
      tdIndex.className = "row-index-cell";
      tr.appendChild(tdIndex);

      row.forEach((v) => {
        const td = document.createElement("td");
        td.textContent = v === null || v === undefined ? "" : v;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    previewWrapper.appendChild(table);

    // 分析与转换配置
    const analysisPanel = document.createElement("div");
    analysisPanel.className = "analysis-panel";

    const analysisHeader = document.createElement("div");
    analysisHeader.className = "analysis-header";
    const statusSpan = document.createElement("span");
    statusSpan.className = "analysis-status";
    if (!sheet.analysis) {
      statusSpan.textContent = "尚未智能分析";
    } else if (sheet.analysis.error) {
      statusSpan.textContent = "智能分析失败：" + sheet.analysis.error;
      statusSpan.classList.add("error");
    } else {
      statusSpan.textContent = "已完成智能分析，可修改下方结果";
    }
    analysisHeader.appendChild(statusSpan);

    const singleConvertBtn = document.createElement("button");
    singleConvertBtn.textContent = "仅转换本 Sheet";
    singleConvertBtn.className = "btn secondary";
    singleConvertBtn.addEventListener("click", () => convertSheets([name]));
    analysisHeader.appendChild(singleConvertBtn);
    analysisPanel.appendChild(analysisHeader);

    // 列名
    const columnsGroup = document.createElement("div");
    columnsGroup.className = "field-group";
    const columnsLabel = document.createElement("label");
    columnsLabel.textContent = "列名（按顺序，Python 列表格式或用逗号分隔）：";
    const columnsInput = document.createElement("textarea");
    columnsInput.value = Array.isArray(sheet.columns)
      ? JSON.stringify(sheet.columns, null, 0)
      : sheet.columns || "";
    columnsInput.addEventListener("input", (e) => {
      sheetsState[name].columns = e.target.value;
    });
    columnsGroup.appendChild(columnsLabel);
    columnsGroup.appendChild(columnsInput);

    // 起始行
    const rowGroup = document.createElement("div");
    rowGroup.className = "field-group";
    const rowLabel = document.createElement("label");
    rowLabel.textContent = "真正数据起始行（从 1 开始，基于整张 Sheet 的行号）：";
    const rowInput = document.createElement("input");
    rowInput.type = "number";
    rowInput.min = "1";
    rowInput.value = sheet.dataStartRow || 1;
    rowInput.addEventListener("input", (e) => {
      sheetsState[name].dataStartRow = parseInt(e.target.value || "1", 10);
    });
    rowGroup.appendChild(rowLabel);
    rowGroup.appendChild(rowInput);

    analysisPanel.appendChild(columnsGroup);
    analysisPanel.appendChild(rowGroup);

    // json 预览
    if (sheet.json) {
      const jsonWrapper = document.createElement("div");
      jsonWrapper.className = "json-preview";
      const title = document.createElement("div");
      title.className = "json-preview-title";
      title.textContent = "JSON 预览（最多展示前 50 条）：";
      const pre = document.createElement("pre");
      const previewData = sheet.json.slice(0, 50);
      pre.textContent = JSON.stringify(previewData, null, 2);

      const downloadBtn = document.createElement("button");
      downloadBtn.className = "btn secondary";
      downloadBtn.style.marginTop = "6px";
      downloadBtn.textContent = "下载 JSON 文件";
      downloadBtn.addEventListener("click", () => downloadJson(name, sheet.json));

      jsonWrapper.appendChild(title);
      jsonWrapper.appendChild(pre);
      jsonWrapper.appendChild(downloadBtn);
      analysisPanel.appendChild(jsonWrapper);
    }

    body.appendChild(previewWrapper);
    body.appendChild(analysisPanel);

    card.appendChild(body);

    // footer：占位，可展示额外信息
    const footer = document.createElement("div");
    footer.className = "sheet-footer";
    footer.innerHTML = `<span></span>`;
    card.appendChild(footer);

    container.appendChild(card);
  });

  // 绑定选择事件
  container.querySelectorAll(".sheet-select").forEach((checkbox) => {
    checkbox.addEventListener("change", (e) => {
      const name = e.target.getAttribute("data-name");
      sheetsState[name].selected = e.target.checked;
      // 同步全选状态
      const allSelected =
        Object.keys(sheetsState).length > 0 && Object.values(sheetsState).every((s) => s.selected);
      $("selectAllSheets").checked = allSelected;
      updateAnalyzeButtonState();
      updateConvertButtonsState();
    });
  });
}

async function uploadFile() {
  const fileInput = $("fileInput");
  const file = fileInput.files[0];
  const uploadStatusEl = $("uploadStatus");

  if (!file) {
    setStatus(uploadStatusEl, "请先选择要上传的 Excel 文件。", "error");
    return;
  }

  setStatus(uploadStatusEl, "正在上传并解析，请稍候...", "success");

  const form = new FormData();
  form.append("file", file);

  try {
    const resp = await fetch("/upload", {
      method: "POST",
      body: form,
    });
    const data = await resp.json();
    if (!data.ok) {
      setStatus(uploadStatusEl, data.error || "上传失败", "error");
      return;
    }

    currentFileId = data.file_id;
    sheetsState = {};
    (data.sheets || []).forEach((sheet) => {
      sheetsState[sheet.sheet_name] = {
        selected: true,
        rows: sheet.rows || [],
        rowCount: sheet.row_count || 0,
        colCount: sheet.col_count || 0,
        analysis: null,
        columns: "",
        dataStartRow: 1,
        json: null,
      };
    });

    $("sheetsSection").style.display = "block";
    $("convertSection").style.display = "block";
    $("selectAllSheets").checked = true;
    setStatus(uploadStatusEl, "上传并预览成功。", "success");

    renderSheets();
    updateAnalyzeButtonState();
    updateConvertButtonsState();
  } catch (e) {
    console.error(e);
    setStatus(uploadStatusEl, "上传或解析失败，请检查服务端日志。", "error");
  }
}

async function analyzeSelectedSheets() {
  if (!currentFileId) return;

  const selectedNames = Object.keys(sheetsState).filter((n) => sheetsState[n].selected);
  if (!selectedNames.length) return;

  $("analyzeBtn").disabled = true;
  setStatus($("convertStatus"), "正在调用智能分析，请稍候...", "success");

  const reqSheets = selectedNames.map((name) => ({
    sheet_name: name,
    rows: sheetsState[name].rows,
  }));

  try {
    const resp = await fetch("/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        file_id: currentFileId,
        sheets: reqSheets,
      }),
    });
    const data = await resp.json();
    if (!data.ok) {
      setStatus($("convertStatus"), data.error || "智能分析失败。", "error");
    } else {
      const { results = {}, errors = {} } = data;
      Object.keys(results).forEach((name) => {
        const r = results[name] || {};
        sheetsState[name].analysis = { error: null };
        sheetsState[name].columns = Array.isArray(r.columns) ? JSON.stringify(r.columns) : "";
        sheetsState[name].dataStartRow = r.data_start_row || 1;
      });
      Object.keys(errors).forEach((name) => {
        sheetsState[name].analysis = { error: errors[name] };
      });
      setStatus($("convertStatus"), "智能分析完成，可在各个 Sheet 中查看和调整结果。", "success");
      renderSheets();
    }
  } catch (e) {
    console.error(e);
    setStatus($("convertStatus"), "智能分析请求失败，请检查服务端日志。", "error");
  } finally {
    $("analyzeBtn").disabled = false;
  }
}

function normalizeColumnsInput(raw) {
  if (!raw) return [];
  const s = raw.trim();
  // 尝试 JSON / Python 列表
  if (s.startsWith("[") && s.endsWith("]")) {
    try {
      // 将单引号替换为双引号，便于 JSON 解析
      const jsonLike = s.replace(/'/g, '"');
      const arr = JSON.parse(jsonLike);
      if (Array.isArray(arr)) {
        return arr.map((x) => String(x).trim()).filter(Boolean);
      }
    } catch (e) {
      // ignore
    }
  }
  // 回退：按逗号分割
  return s
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}

async function convertSheets(names) {
  if (!currentFileId) return;
  const selected = names && names.length ? names : Object.keys(sheetsState).filter((n) => sheetsState[n].selected);
  if (!selected.length) return;

  $("convertSelectedBtn").disabled = true;
  setStatus($("convertStatus"), "正在转换为 JSON...", "success");

  const reqSheets = selected.map((name) => {
    const s = sheetsState[name];
    const cols = normalizeColumnsInput(s.columns);
    return {
      sheet_name: name,
      columns: cols.length ? cols : s.columns, // 如果能解析则用数组，否则把字符串交给后端
      data_start_row: s.dataStartRow || 1,
    };
  });

  try {
    const resp = await fetch("/convert", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        file_id: currentFileId,
        sheets: reqSheets,
      }),
    });
    const data = await resp.json();
    if (!data.ok) {
      setStatus($("convertStatus"), data.error || "转换失败。", "error");
    } else {
      const { converted = {}, errors = {} } = data;
      Object.keys(converted).forEach((name) => {
        sheetsState[name].json = converted[name];
      });
      Object.keys(errors).forEach((name) => {
        if (!sheetsState[name].analysis) sheetsState[name].analysis = {};
        sheetsState[name].analysis.error = errors[name];
      });
      setStatus($("convertStatus"), "转换完成，可在各个 Sheet 中预览及下载 JSON。", "success");
      renderSheets();
    }
  } catch (e) {
    console.error(e);
    setStatus($("convertStatus"), "转换请求失败，请检查服务端日志。", "error");
  } finally {
    $("convertSelectedBtn").disabled = false;
  }
}

function downloadJson(sheetName, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${sheetName}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function initEvents() {
  $("uploadBtn").addEventListener("click", uploadFile);
  $("analyzeBtn").addEventListener("click", analyzeSelectedSheets);
  $("convertSelectedBtn").addEventListener("click", () => convertSheets());

  $("selectAllSheets").addEventListener("change", (e) => {
    const checked = e.target.checked;
    Object.keys(sheetsState).forEach((name) => {
      sheetsState[name].selected = checked;
    });
    renderSheets();
    updateAnalyzeButtonState();
    updateConvertButtonsState();
  });
}

document.addEventListener("DOMContentLoaded", initEvents);

