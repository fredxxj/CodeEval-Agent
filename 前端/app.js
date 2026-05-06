const defaultInputDir = "/Users/harking/x/智能体批改项目/大模型自动批改";

const scaleText = {
  低: "低尺度：更鼓励完成度，适合初学阶段和形成性评价。",
  中: "中尺度：兼顾功能完成、代码质量和改进建议，是当前默认阅卷尺度。",
  高: "高尺度：更严格关注边界条件、结构封装和可维护性，适合结课评价。",
};

let activePanelName = "import";
let appState = {
  running: false,
  input_dir: defaultInputDir,
  output_file: "评分结果.xlsx",
  total: 0,
  processed: 0,
  success: 0,
  failed: 0,
  skipped: 0,
  current_file: "",
  message: "后端已就绪",
  logs: [],
  scale: "中",
};
let resultRows = [];

const activePanel = document.querySelector("#activePanel");
const panelButtons = document.querySelectorAll("[data-panel]");
const scaleSelect = document.querySelector("#scaleSelect");
const startGrading = document.querySelector("#startGrading");
const folderPicker = document.querySelector("#folderPicker");
const runState = document.querySelector("#runState");
const heroMetrics = document.querySelector("#heroMetrics");
const sidePanel = document.querySelector("#sidePanel");

function percent() {
  if (!appState.total) return appState.success ? 100 : 0;
  return Math.min(100, Math.round((appState.success / appState.total) * 1000) / 10);
}

function statusLabel() {
  if (appState.running) return "批改运行中";
  if (appState.success) return `已完成 ${appState.success} 份`;
  return "后端已就绪";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function getJson(url, options) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "请求失败");
  return data;
}

async function refreshStatus() {
  try {
    appState = await getJson("/api/status");
    const results = await getJson("/api/results");
    resultRows = results.rows || [];
    renderChrome();
    renderPanel(activePanelName);
  } catch (error) {
    appState.message = `无法连接后端：${error.message}`;
    renderChrome();
  }
}

function renderChrome() {
  const done = appState.success || 0;
  const total = appState.total || Math.max(done, resultRows.length, 0);
  const average = resultRows.length
    ? (resultRows.reduce((sum, row) => sum + Number(row["总分"] || 0), 0) / resultRows.length).toFixed(1)
    : "--";
  const highest = resultRows.length
    ? Math.max(...resultRows.map((row) => Number(row["总分"] || 0)))
    : "--";

  runState.innerHTML = `<span></span>${escapeHtml(statusLabel())}`;
  runState.classList.toggle("running", Boolean(appState.running));
  heroMetrics.innerHTML = `
    <div><span>已完成</span><strong>${done} / ${total || 0}</strong></div>
    <div><span>平均分</span><strong>${average}</strong></div>
    <div><span>最高分</span><strong>${highest}</strong></div>
  `;

  sidePanel.innerHTML = `
    <h2>本轮批改概览</h2>
    <div class="progress-block">
      <div class="progress-head">
        <span>完成进度</span>
        <strong>${percent()}%</strong>
      </div>
      <div class="progress-track"><i style="width: ${percent()}%"></i></div>
    </div>
    <dl class="quick-stats">
      <div><dt>成功批改</dt><dd>${done} 份</dd></div>
      <div><dt>失败 / 待续跑</dt><dd>${appState.failed || 0} 份</dd></div>
      <div><dt>已跳过成功项</dt><dd>${appState.skipped || 0} 份</dd></div>
      <div><dt>评价维度</dt><dd>功能 / 鲁棒 / 效率 / 维护</dd></div>
    </dl>
    <p class="note">${escapeHtml(appState.message || "等待操作")}</p>
  `;
}

function setActive(name) {
  activePanelName = name;
  panelButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.panel === name);
  });
}

function renderPanel(name) {
  setActive(name);
  if (name === "import") renderImport();
  if (name === "progress") renderProgress();
  if (name === "results") renderResults();
}

function renderImport(files = []) {
  activePanel.innerHTML = `
    <p class="eyebrow">Input Folder</p>
    <h2>学生试题位置导入</h2>
    <div class="folder-picker">
      <label>
        <span>当前目录</span>
        <input id="folderPath" value="${escapeHtml(appState.input_dir || defaultInputDir)}" />
      </label>
      <button id="scanFolder" type="button">扫描目录</button>
    </div>
    <label class="api-field">
      <span>API Key（录制时可留空，服务端已设置则自动使用）</span>
      <input id="apiKey" type="password" placeholder="BIGMODEL_API_KEY" autocomplete="off" />
    </label>
    <div class="file-preview" id="filePreview">
      ${
        files.length
          ? files.slice(0, 8).map((file) => `<span>${escapeHtml(file)}</span>`).join("")
          : `<span>${appState.total || 0} 份 Word 作业已记录在后端状态中</span>`
      }
    </div>
    <p id="folderHint">系统递归读取该目录下的 Word 文档，自动提取姓名、学号和答题内容。</p>
  `;
  document.querySelector("#scanFolder").addEventListener("click", scanFolder);
}

function renderScale() {
  activePanelName = "scale";
  panelButtons.forEach((button) => button.classList.remove("active"));
  activePanel.innerHTML = `
    <p class="eyebrow">Grading Scale</p>
    <h2>阅卷尺度：${scaleSelect.value}</h2>
    <p>${scaleText[scaleSelect.value]}</p>
    <div class="scale-preview">
      <span>功能性 30</span>
      <span>鲁棒性 20</span>
      <span>效率性 20</span>
      <span>可维护性 30</span>
    </div>
  `;
}

function renderProgress() {
  activePanel.innerHTML = `
    <p class="eyebrow">Progress</p>
    <h2>进度查看</h2>
    <div class="progress-large">
      <strong>${appState.success || 0} / ${appState.total || 0}</strong>
      <span>${escapeHtml(appState.running ? appState.current_file || "正在批改" : appState.message || "等待开始")}</span>
      <div class="progress-track"><i style="width: ${percent()}%"></i></div>
    </div>
    <ul class="status-list">
      ${(appState.logs || []).slice(-8).map((line) => `<li><span class="dot done"></span>${escapeHtml(line)}</li>`).join("")}
    </ul>
  `;
}

function renderResults() {
  const rows = resultRows.slice(0, 12);
  activePanel.innerHTML = `
    <p class="eyebrow">Results</p>
    <h2>结果查看</h2>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>学生</th>
            <th>功能</th>
            <th>鲁棒</th>
            <th>效率</th>
            <th>维护</th>
            <th>总分</th>
            <th>状态</th>
          </tr>
        </thead>
        <tbody>
          ${
            rows.length
              ? rows.map((row, index) => `
                <tr>
                  <td>学生${String(index + 1).padStart(2, "0")}</td>
                  <td>${row["功能性得分"]}</td>
                  <td>${row["鲁棒性得分"]}</td>
                  <td>${row["效率性得分"]}</td>
                  <td>${row["可维护性得分"]}</td>
                  <td><b>${row["总分"]}</b></td>
                  <td><span class="pill ${row["解析状态"] === "成功" ? "done" : "pending"}">${escapeHtml(row["解析状态"])}</span></td>
                </tr>
              `).join("")
              : `<tr><td colspan="7">暂无结果，点击“开始改卷”后自动生成。</td></tr>`
          }
        </tbody>
      </table>
    </div>
    <p>结果文件：${escapeHtml(appState.output_file || "评分结果.xlsx")}。页面仅展示匿名化摘要，完整评分理由和改进建议保存在 Excel 中。</p>
  `;
}

async function scanFolder() {
  const input = document.querySelector("#folderPath");
  const folder = input.value.trim();
  try {
    const data = await getJson(`/api/files?path=${encodeURIComponent(folder)}`);
    appState.input_dir = folder;
    appState.total = data.total;
    renderImport(data.files || []);
    document.querySelector("#folderHint").textContent = `已扫描到 ${data.total} 份 Word 答题记录。`;
    renderChrome();
  } catch (error) {
    document.querySelector("#folderHint").textContent = error.message;
  }
}

async function startBackend() {
  const folderInput = document.querySelector("#folderPath");
  const apiKeyInput = document.querySelector("#apiKey");
  const payload = {
    input_dir: folderInput?.value?.trim() || appState.input_dir || defaultInputDir,
    api_key: apiKeyInput?.value?.trim() || "",
    scale: scaleSelect.value,
  };

  try {
    await getJson("/api/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    activePanelName = "progress";
    await refreshStatus();
  } catch (error) {
    appState.message = error.message;
    renderChrome();
    renderProgress();
  }
}

panelButtons.forEach((button) => {
  button.addEventListener("click", () => renderPanel(button.dataset.panel));
});

scaleSelect.addEventListener("change", renderScale);
scaleSelect.addEventListener("focus", renderScale);
startGrading.addEventListener("click", startBackend);
folderPicker.addEventListener("change", () => {});

refreshStatus();
setInterval(refreshStatus, 2000);
