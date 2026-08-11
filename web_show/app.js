/* ============================================================
   人脸识别考勤系统 - Web 前端逻辑
   ============================================================ */

"use strict";

/* ---------- 元数据 ---------- */
const INPUT_FILES = [
  { key: "photo_ledger",        name: "照片台账表",            file: "01_照片台账表.xlsx" },
  { key: "work_plan",           name: "工作安排表",            file: "02_工作安排表.xlsx" },
  { key: "person_class",        name: "人员分类表",            file: "03_人员分类表.xlsx" },
  { key: "project_info",        name: "项目信息表",            file: "04_项目信息表.xlsx" },
  { key: "attendance_template", name: "工程与外协考勤表模板",   file: "05_工程与外协考勤表模板.xlsx" },
];

const TABLES = [
  { key: "table1",     label: "表1 相机出工信息" },
  { key: "table2",     label: "表2 工作安排" },
  { key: "table3",     label: "表3 识别结果表" },
  { key: "table4",     label: "表4 识别出工人" },
  { key: "table5",     label: "表5 该日出工人员" },
  { key: "table6",     label: "表6 全体出工情况" },
  { key: "table7",     label: "表7 出工地点时长" },
  { key: "table8_9_10", label: "表8/9/10 最终考勤" },
];

const SEQUENCE = TABLES.map(t => t.key);

const DL_FILES = [
  { file: "08_09_10_最终考勤表.xlsx", name: "最终考勤表（表8/9/10）" },
];

/* ---------- 工具函数 ---------- */
async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  const ct = res.headers.get("content-type") || "";
  if (ct.includes("application/json")) return await res.json();
  return res;
}

function toast(msg, type = "") {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast show " + type;
  clearTimeout(el._timer);
  el._timer = setTimeout(() => { el.className = "toast"; }, 2600);
}

function setMask(show, text = "处理中…") {
  const m = document.getElementById("loading-mask");
  document.getElementById("loading-text").textContent = text;
  m.hidden = !show;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

/* ---------- 状态渲染 ---------- */
let status = null;

/* 人工复核状态 */
let reviewPhotos = [];      // 待复核照片列表
let reviewIdx = 0;          // 当前复核的照片索引
let reviewHandled = false;  // 本次会话是否已完成/跳过复核
let reviewKnownNames = [];  // 人脸库姓名，用作输入建议

async function renderStatus() {
  try {
    status = await api("/api/status");
  } catch (e) {
    toast("无法连接服务器", "error");
    return;
  }
  renderImageCounts();
  renderInputList();
  renderChipsFromStatus();
  renderStepButtons();
  renderStep4Cards();
  renderReviewPanel();
  const clearBtn = document.getElementById("btn-clear");
  if (clearBtn) clearBtn.disabled = status.recognition_running;
}

function renderReviewPanel() {
  const panel = document.getElementById("review-panel");
  if (!panel) return;
  panel.hidden = !status.review_pending;
  if (!status.review_pending) return;
  const parts = [`有 <b>${status.review_photos || 0}</b> 张照片含未识别人员，需人工复核修正`];
  if (status.no_face_photos > 0) {
    parts.push(`另有 <b>${status.no_face_photos}</b> 张照片未检测到人脸（不纳入复核，请另行核实）`);
  }
  document.getElementById("review-summary").innerHTML = parts.join("；");
}

function renderImageCounts() {
  const ref = status.image_counts.ref || 0;
  const tgt = status.image_counts.target || 0;
  setCount("count-ref", ref);
  setCount("count-target", tgt);
  const btn = document.getElementById("btn-recognize");
  btn.disabled = !(ref > 0 && tgt > 0);
  if (ref > 0 && tgt > 0) btn.title = "";
}

function setCount(id, n) {
  const el = document.getElementById(id);
  el.textContent = n + " 张图片";
  el.classList.toggle("has-image", n > 0);
}

function renderInputList() {
  const list = document.getElementById("input-list");
  list.innerHTML = "";
  const uploads = status.uploads || {};
  for (const f of INPUT_FILES) {
    const ok = !!uploads[f.key];
    const row = document.createElement("div");
    row.className = "input-row";
    row.innerHTML = `
      <span class="name">${f.name}</span>
      <span class="badge ${ok ? "ok" : "missing"}">${ok ? "已上传 ✓" : "未上传"}</span>
      <span class="desc">input/${f.file}</span>
      <button class="btn small ${ok ? "ghost" : "primary"}" data-upload-key="${f.key}">${ok ? "重新上传" : "上传"}</button>
    `;
    row.querySelector("[data-upload-key]").addEventListener("click", () => pickAndUpload(f));
    list.appendChild(row);
  }
}

function renderChipsFromStatus() {
  const tables = status.tables || {};
  const wrap = document.getElementById("table-chips");
  if (!wrap) return;
  wrap.innerHTML = "";
  for (const t of TABLES) {
    const chip = document.createElement("div");
    chip.className = "table-chip " + (tables[t.key] ? "done" : "pending");
    chip.dataset.key = t.key;
    chip.innerHTML = `<span class="chip-dot"></span>${t.label}`;
    wrap.appendChild(chip);
  }
}

function renderStepButtons() {
  const tables = status.tables || {};
  const allInputs = INPUT_FILES.every(f => status.uploads && status.uploads[f.key]);

  // 第1步 → 第2步：识别已完成（且存在未识别人脸时必须先完成/跳过人工复核）
  const recognitionDone = status.recognition_done || !!tables["table3"];
  const reviewBlocking = status.review_pending && !reviewHandled;
  const next1 = document.getElementById("btn-next-1");
  if (next1) next1.disabled = !recognitionDone || reviewBlocking;

  // 第2步 → 第3步：5 个输入表齐全
  const next2 = document.getElementById("btn-next-2");
  if (next2) next2.disabled = !allInputs;

  // 第3步开始生成：输入表齐全
  const gen = document.getElementById("btn-generate");
  if (gen) gen.disabled = !allInputs;

  // 若 8/9/10 已存在，提供直接进入完成页
  const skip = document.getElementById("btn-skip-final");
  if (skip) {
    const finalDone = DL_FILES.every(d => {
      return status.outputs && status.outputs.some(o => o.name === d.file);
    });
    skip.hidden = !finalDone;
  }
}

function renderStep4Cards() {
  const outputs = (status.outputs || []).map(o => o.name);
  for (const d of DL_FILES) {
    const card = document.querySelector(`.dl-card[data-file="${d.file}"]`);
    if (!card) continue;
    const exists = outputs.includes(d.file);
    const btn = card.querySelector(".js-download");
    btn.disabled = !exists;
    btn.textContent = exists ? "下载" : "未生成";
  }
}

/* ---------- 步骤切换 ---------- */
function gotoStep(n) {
  for (let i = 1; i <= 4; i++) {
    document.getElementById("step-" + i).hidden = i !== n;
  }
  document.querySelectorAll(".stepper .step").forEach(el => {
    const s = parseInt(el.dataset.step, 10);
    el.classList.toggle("active", s === n);
    el.classList.toggle("done", s < n);
  });
  document.querySelectorAll(".stepper .step-line").forEach(el => {
    const s = parseInt(el.dataset.line, 10);
    el.classList.toggle("done", s < n);
  });
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function initStepper() {
  document.querySelectorAll(".stepper .step").forEach(el => {
    el.addEventListener("click", () => gotoStep(parseInt(el.dataset.step, 10)));
  });
  document.querySelectorAll("[data-goto]").forEach(el => {
    el.addEventListener("click", () => gotoStep(parseInt(el.dataset.goto, 10)));
  });
}

/* ---------- 上传 ---------- */
function setupDropzone(zoneId, fileInputId, apiZone) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(fileInputId);
  if (!zone) return;

  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => {
    if (input.files.length) uploadImages(apiZone, input.files);
    input.value = "";
  });

  ["dragenter", "dragover"].forEach(ev => zone.addEventListener(ev, e => {
    e.preventDefault(); zone.classList.add("dragover");
  }));
  ["dragleave", "drop"].forEach(ev => zone.addEventListener(ev, e => {
    e.preventDefault(); zone.classList.remove("dragover");
  }));
  zone.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) uploadImages(apiZone, e.dataTransfer.files);
  });
}

async function uploadImages(zone, fileList) {
  const files = Array.from(fileList).filter(f => /\.(jpe?g|png|bmp|webp)$/i.test(f.name));
  if (!files.length) { toast("请选择图片文件（jpg/png/bmp/webp）", "error"); return; }
  setMask(true, "上传图片中…");
  try {
    const fd = new FormData();
    files.forEach(f => fd.append("files", f));
    const res = await api("/api/upload_images/" + zone, { method: "POST", body: fd });
    if (res.success) {
      toast(`已上传 ${res.saved.length} 张图片`, "success");
    } else {
      toast(res.message || "上传失败", "error");
    }
  } finally {
    setMask(false);
  }
  await renderStatus();
}

function pickAndUpload(f) {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xlsx";
  input.onchange = async () => {
    if (!input.files.length) return;
    setMask(true, "上传 " + f.name + "…");
    try {
      const fd = new FormData();
      fd.append("file", input.files[0]);
      const res = await api("/api/upload_input/" + f.key, { method: "POST", body: fd });
      toast(res.success ? res.message : res.message, res.success ? "success" : "error");
    } finally {
      setMask(false);
    }
    await renderStatus();
  };
  input.click();
}

/* ---------- 人脸识别 ---------- */
async function startRecognition() {
  const btn = document.getElementById("btn-recognize");
  btn.disabled = true;
  reviewHandled = false;  // 重新识别后需重新复核
  reviewPhotos = [];
  const consoleEl = document.getElementById("recog-console");
  consoleEl.hidden = false;
  consoleEl.textContent = "🚀 人脸识别启动中…\n";
  document.getElementById("recog-result").hidden = true;
  try {
    const res = await api("/api/recognize", { method: "POST" });
    if (res.started) {
      toast("人脸识别已启动", "success");
      await pollRecognition();
    } else {
      appendLog(consoleEl, res.message || "无法启动");
      btn.disabled = false;
    }
  } catch (e) {
    appendLog(consoleEl, "请求失败: " + e);
    btn.disabled = false;
  }
}

async function pollRecognition() {
  const consoleEl = document.getElementById("recog-console");
  let lastCount = 0;
  for (;;) {
    await sleep(900);
    let p;
    try { p = await api("/api/progress"); } catch (e) { continue; }

    if (p.logs && p.logs.length > lastCount) {
      for (const line of p.logs.slice(lastCount)) appendLog(consoleEl, line);
      lastCount = p.logs.length;
    }

    if (!p.running) {
      renderRecognitionResult(p);
      await renderStatus();
      // 识别完成后自动弹出人工复核（存在未识别人脸时）
      if (p.done && p.success && status.review_pending && !reviewHandled) {
        openReview();
      }
      return;
    }
  }
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function appendLog(el, text) {
  el.textContent += text + "\n";
  el.scrollTop = el.scrollHeight;
}

function renderRecognitionResult(p) {
  const card = document.getElementById("recog-result");
  card.hidden = false;
  if (p.done && p.success) {
    card.className = "result-card";
    card.innerHTML = `<div class="rc-title">✅ 人脸识别完成</div>
      已生成 <a href="/api/download/${encodeURIComponent("00_程序人脸识别结果.xlsx")}">00_程序人脸识别结果.xlsx</a>
      ${p.message ? `<div style="margin-top:4px;color:var(--muted);font-size:13px">${esc(p.message)}</div>` : ""}`;
    document.getElementById("btn-recognize").disabled = true;
  } else {
    card.className = "result-card";
    card.style.borderColor = "rgba(255,59,48,.35)";
    card.style.background = "#FFEBEA";
    card.innerHTML = `<div class="rc-title" style="color:var(--danger)">❌ 人脸识别失败</div>
      <div style="font-size:13px;color:var(--text)">${esc(p.message || "未知错误")}</div>`;
    document.getElementById("btn-recognize").disabled = false;
  }
}

/* ---------- 人工复核（修正未识别人脸） ---------- */
function renderNameList() {
  let dl = document.getElementById("review-name-list");
  if (!dl) {
    dl = document.createElement("datalist");
    dl.id = "review-name-list";
    document.body.appendChild(dl);
  }
  dl.innerHTML = "";
  (reviewKnownNames || []).forEach(n => {
    const opt = document.createElement("option");
    opt.value = n;
    dl.appendChild(opt);
  });
}

async function openReview() {
  const startBtn = document.getElementById("btn-review-start");
  const summaryEl = document.getElementById("review-summary");
  startBtn.disabled = true;
  let data;
  try {
    const s = await api("/api/review/start", { method: "POST" });
    if (!s.success) {
      toast(s.message || "无法开始复核", "error");
      return;
    }
    // 后台构建（加载模型 + 重检照片）可能耗时数十秒：进度显示在复核面板上，
    // 不遮挡界面，准备期间仍可点「跳过本次复核」取消。
    summaryEl.innerHTML = "⏳ 正在准备复核数据…";
    for (;;) {
      await sleep(800);
      const p = await api("/api/review/progress");
      if (p.progress) summaryEl.innerHTML = "⏳ " + esc(p.progress);
      if (p.done) {
        if (!p.success) {
          toast(p.message === "已取消" ? "已取消复核" : (p.message || "复核数据准备失败"),
                p.message === "已取消" ? "" : "error");
          return;
        }
        break;
      }
    }
    data = await api("/api/review/data");
    if (!data.success) {
      toast(data.message || "复核数据加载失败", "error");
      return;
    }
  } catch (e) {
    toast("复核数据加载失败: " + e, "error");
    return;
  } finally {
    startBtn.disabled = false;
  }
  reviewKnownNames = data.known_names || [];
  renderNameList();
  reviewPhotos = (data.photos || []).slice();
  if (!reviewPhotos.length) {
    reviewHandled = true;
    renderReviewPanel();
    toast("没有需要复核的照片", "success");
    await renderStatus();
    return;
  }
  reviewIdx = 0;
  document.getElementById("modal-review").hidden = false;
  renderReviewPhoto();
}

function renderReviewPhoto() {
  if (reviewIdx >= reviewPhotos.length) {
    finishReview();
    return;
  }
  const p = reviewPhotos[reviewIdx];
  document.getElementById("rv-title").textContent =
    `人工复核未识别人员（${reviewIdx + 1}/${reviewPhotos.length}）`;
  document.getElementById("rv-sub").textContent = p.photo;

  const warn = document.getElementById("rv-warning");
  if (p.warning) {
    warn.textContent = "⚠️ " + p.warning;
    warn.hidden = false;
  } else {
    warn.hidden = true;
  }

  // 信息栏：照片人员名单（来自台账表）、拍照人、识别结果
  const plEl = document.getElementById("rv-person-list");
  if (p.person_list && p.person_list.length) {
    plEl.textContent = p.person_list.join("、");
    plEl.classList.remove("rv-list-warn");
  } else {
    plEl.textContent = p.list_warning || "（无人员名单）";
    plEl.classList.add("rv-list-warn");
  }
  document.getElementById("rv-reporter").textContent = p.reporter || "（无法解析）";
  document.getElementById("rv-recognized").textContent =
    (p.recognized_names && p.recognized_names.length) ? p.recognized_names.join("、") : "无";

  // 差异提示（拍照人无法上镜，不算异常，已在后端排除）
  const d = p.diffs || {};
  const diffEl = document.getElementById("rv-diffs");
  const chips = [];
  if (d.report_not_recognized && d.report_not_recognized.length) {
    chips.push(`<span class="rv-diff-chip warn">名单中有识别中无（不含拍照人）：${esc(d.report_not_recognized.join("、"))}</span>`);
  }
  if (d.recognized_not_report && d.recognized_not_report.length) {
    chips.push(`<span class="rv-diff-chip info">识别中有名单中无：${esc(d.recognized_not_report.join("、"))}</span>`);
  }
  if (!chips.length) {
    chips.push(`<span class="rv-diff-chip ok">名单与识别一致</span>`);
  }
  diffEl.innerHTML = chips.join("");

  // 照片 + 未识别人脸红框（按原图比例 % 定位，自适应缩放）
  const wrap = document.getElementById("rv-img-wrap");
  wrap.innerHTML = "";
  const img = document.createElement("img");
  img.className = "rv-img";
  img.alt = p.photo;
  img.onerror = () => {
    const err = document.createElement("div");
    err.className = "rv-img-error";
    err.textContent = "原图无法加载（可能不在 Target_Figure 中）";
    wrap.appendChild(err);
  };
  img.onload = () => {
    const nw = img.naturalWidth || 1;
    const nh = img.naturalHeight || 1;
    (p.unknown_faces || []).forEach((f, i) => {
      if (!f.box || f.box.length < 4) return;
      const [x1, y1, x2, y2] = f.box;
      // 不画红框（图片小时会遮脸），直接在脸部框正下方（胸口/身体处）标序号
      const num = document.createElement("span");
      num.className = "rv-box-num";
      num.textContent = i + 1;
      num.style.left = (((x1 + x2) / 2) / nw * 100) + "%";
      num.style.top = (y2 / nh * 100) + "%";
      wrap.appendChild(num);
    });
  };
  img.src = "/api/target_image/" + encodeURIComponent(p.photo);
  wrap.appendChild(img);

  // 姓名输入框：有几个人未识别就给几个框
  const inputsEl = document.getElementById("rv-inputs");
  inputsEl.innerHTML = "";
  if (!(p.unknown_faces || []).length) {
    inputsEl.innerHTML = `<div class="rv-empty">该照片没有可框选的未识别人脸${p.warning ? `（${esc(p.warning)}）` : ""}</div>`;
  } else {
    (p.unknown_faces || []).forEach((f, i) => {
      const row = document.createElement("div");
      row.className = "rv-input-row";
      const label = document.createElement("label");
      label.innerHTML = `未识别人脸 <b>${i + 1}</b>`;
      const input = document.createElement("input");
      input.className = "rv-name-input";
      input.placeholder = "请输入姓名（可留空跳过）";
      input.dataset.seq = f.seq || "";
      input.setAttribute("list", "review-name-list");
      input.addEventListener("keydown", e => {
        if (e.key === "Enter") { e.preventDefault(); confirmReviewPhoto(); }
      });
      row.appendChild(label);
      row.appendChild(input);
      inputsEl.appendChild(row);
    });
    const first = inputsEl.querySelector(".rv-name-input");
    if (first) first.focus();
  }
}

async function confirmReviewPhoto() {
  const p = reviewPhotos[reviewIdx];
  if (!p) return;
  const corrections = [];
  document.querySelectorAll("#rv-inputs .rv-name-input").forEach(inp => {
    const name = inp.value.trim();
    if (name) corrections.push({ seq: inp.dataset.seq || "", name });
  });
  setMask(true, "正在写入识别结果…");
  try {
    const res = await api("/api/review_submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ photo: p.photo, corrections }),
    });
    if (res.success) {
      if (res.updated > 0) toast(`已更新 ${res.updated} 条识别记录`, "success");
      else toast("未填写姓名，已跳过本张", "");
      reviewPhotos.splice(reviewIdx, 1);  // 本张处理完，进入下一张
      renderReviewPhoto();
    } else {
      toast(res.message || "写入失败", "error");
    }
  } catch (e) {
    toast("请求失败: " + e, "error");
  } finally {
    setMask(false);
  }
}

function skipReviewPhoto() {
  if (reviewIdx >= reviewPhotos.length) return;
  reviewPhotos.splice(reviewIdx, 1);
  renderReviewPhoto();
}

function skipAllReview() {
  api("/api/review/cancel", { method: "POST" });  // 取消后台构建，避免占用模型锁
  reviewPhotos = [];
  reviewHandled = true;
  document.getElementById("modal-review").hidden = true;
  renderReviewPanel();
  toast("已跳过本次复核", "");
  renderStatus();
}

function finishReview() {
  document.getElementById("modal-review").hidden = true;
  reviewHandled = true;
  renderReviewPanel();
  toast("复核完成", "success");
  renderStatus();
}

/* ---------- 生成表格（Step3） ---------- */
/* ---------- 生成流程（表1→…→最终表）屏栈导航 ---------- */
const GEN_SCREENS = [
  { table: "table1", type: "view", viewKey: "table1", title: "表1 今日相机出工信息提取表",
    hint: "红色高亮 = 存在开工/收工汇报异常的人员", withFilters: true },
  { table: "table5", type: "edit", label: "05_该日出工人员表", errorKey: "error1" },
  { table: "table6", type: "edit", label: "06_全体人员出工情况表", errorKey: "error2" },
  { table: "table8_9_10", type: "view", viewKey: "final", title: "最终考勤表（表8/9/10）",
    hint: "可直接框选表格内容复制到 Excel（合并单元格格式已保留）", withFilters: false },
];
// 从屏幕 i 推进到 i+1 需要运行的中间表
const TABLES_BETWEEN_SCREENS = {
  0: ["table2", "table3", "table4", "table5"],
  1: ["table6"],
  2: ["table7", "table8_9_10"],
};

let lastTableInfo = {};  // 表 key -> table_info（05/06 汇总信息）
let genRegenFrom = 1;    // 首个需要(重新)生成的屏幕索引

async function runOneTable(key) {
  const consoleEl = document.getElementById("gen-console");
  setChipState(key, "running");
  let res;
  try {
    res = await api("/api/run_table/" + key, { method: "POST" });
  } catch (e) {
    setChipState(key, "failed");
    appendLog(consoleEl, "【" + key + "】请求失败: " + e);
    toast("网络错误", "error");
    return null;
  }
  if (res.output) appendLog(consoleEl, res.output);
  setChipState(key, res.success ? "done" : "failed");
  if (res.table_info) lastTableInfo[key] = res.table_info;
  if (!res.success) {
    toast(`❌ ${res.label || key} 生成失败，请查看日志`, "error");
    return null;
  }
  return res;
}

async function startGeneration() {
  const genBtn = document.getElementById("btn-generate");
  genBtn.disabled = true;
  document.getElementById("gen-console").textContent = "";
  resetChips();
  lastTableInfo = {};
  genRegenFrom = 1;

  // 先运行表1，展示屏幕0
  if (!(await runOneTable("table1"))) { genBtn.disabled = false; return; }

  let pos = 0;
  for (;;) {
    const screen = GEN_SCREENS[pos];
    let choice;
    if (screen.type === "view") {
      choice = await waitForView(screen.viewKey, screen.title,
        pos === GEN_SCREENS.length - 1 ? "完成" : "继续",
        screen.hint, pos > 0, !!screen.withFilters);
    } else {
      choice = await waitForEdit(screen.table, screen.errorKey, screen.label,
        lastTableInfo[screen.table], pos > 0);
    }

    if (choice === "back") { pos -= 1; continue; }

    if (choice === "regen") {
      if (!(await runOneTable(screen.table))) { genBtn.disabled = false; return; }
      genRegenFrom = Math.min(genRegenFrom, pos + 1);  // 其后的表需重新生成
      continue;
    }

    // "next" / "saved"（保存并继续）
    if (choice === "saved" && screen.type === "edit") {
      genRegenFrom = Math.min(genRegenFrom, pos + 1);  // 编辑了 pos，其后的需重跑
    }

    const nextPos = pos + 1;
    if (nextPos >= GEN_SCREENS.length) break;  // 全部完成

    if (genRegenFrom <= nextPos) {
      // 重跑 [genRegenFrom-1, nextPos-1] 段，保证目标屏数据最新
      for (let seg = genRegenFrom - 1; seg <= nextPos - 1; seg++) {
        for (const k of TABLES_BETWEEN_SCREENS[seg]) {
          if (!(await runOneTable(k))) { genBtn.disabled = false; return; }
        }
      }
      genRegenFrom = nextPos + 1;
    }
    pos = nextPos;
  }

  genBtn.disabled = false;
  await renderStatus();

  const allDone = SEQUENCE.every(key => {
    const chip = document.querySelector(`.table-chip[data-key="${key}"]`);
    return chip && chip.classList.contains("done");
  });
  if (allDone) {
    toast("🎉 全部表格生成完成！", "success");
    gotoStep(4);
  }
}

function resetChips() {
  document.querySelectorAll(".table-chip").forEach(chip => {
    chip.className = "table-chip pending";
  });
}

function setChipState(key, state) {
  const chip = document.querySelector(`.table-chip[data-key="${key}"]`);
  if (!chip) return;
  chip.className = "table-chip " + state;
}

/* ---------- 表格增强：表头固定、筛选、行高列宽拖拽 ---------- */
function enhanceTable(container, withFilters) {
  container.querySelectorAll(".sheet-table").forEach(table => {
    if (withFilters) enableTableFilters(table);
    enableTableResize(table);
  });
}

function enableTableFilters(table) {
  const headerRow = table.querySelector("thead tr");
  if (!headerRow) return;
  Array.from(headerRow.cells).forEach((th, ci) => {
    if (!th.textContent.trim() || th.querySelector(".th-filter")) return;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "th-filter";
    btn.textContent = "▾";
    btn.title = "筛选";
    btn.addEventListener("click", e => {
      e.preventDefault(); e.stopPropagation();
      showFilterMenu(table, ci, btn);
    });
    th.appendChild(btn);
  });
}

let activeFilterMenu = null;

function showFilterMenu(table, ci, btn) {
  closeFilterMenu();
  const menu = document.createElement("div");
  menu.className = "filter-menu";
  const values = new Set();
  table.querySelectorAll("tbody tr").forEach(tr => {
    const cell = tr.cells[ci];
    if (cell) {
      const v = cell.textContent.trim();
      if (v) values.add(v);
    }
  });
  const current = btn.dataset.filter || "";
  const add = (label, value) => {
    const item = document.createElement("div");
    item.className = "filter-menu-item" + (value === current ? " active" : "");
    item.textContent = label;
    item.addEventListener("click", () => {
      btn.dataset.filter = value || "";
      btn.classList.toggle("active", !!value);
      applyFilter(table, ci, value || "");
      closeFilterMenu();
    });
    menu.appendChild(item);
  };
  add("（全部）", "");
  Array.from(values).sort((a, b) => a.localeCompare(b, "zh")).forEach(v => add(v, v));
  document.body.appendChild(menu);
  activeFilterMenu = menu;
  const rect = btn.getBoundingClientRect();
  menu.style.left = Math.min(rect.left, window.innerWidth - 170) + "px";
  menu.style.top = (rect.bottom + 4) + "px";
  setTimeout(() => {
    document.addEventListener("mousedown", function handler(ev) {
      if (!menu.contains(ev.target)) { closeFilterMenu(); document.removeEventListener("mousedown", handler); }
    });
  }, 0);
}

function closeFilterMenu() {
  if (activeFilterMenu) { activeFilterMenu.remove(); activeFilterMenu = null; }
}

function applyFilter(table, ci, value) {
  table.querySelectorAll("tbody tr").forEach(tr => {
    const cell = tr.cells[ci];
    const v = cell ? cell.textContent.trim() : "";
    tr.style.display = (!value || v === value) ? "" : "none";
  });
}

function getOrCreateCol(table, ci) {
  let cg = table.querySelector("colgroup");
  const n = table.querySelector("thead tr").cells.length;
  if (!cg) {
    cg = document.createElement("colgroup");
    table.insertBefore(cg, table.firstChild);
    for (let i = 0; i < n; i++) cg.appendChild(document.createElement("col"));
  } else while (cg.children.length < n) cg.appendChild(document.createElement("col"));
  return cg.children[ci];
}

function startColResize(table, ci, e) {
  e.preventDefault();
  const th = table.querySelector("thead tr").cells[ci];
  const startX = e.clientX;
  const startW = th.offsetWidth;
  const col = getOrCreateCol(table, ci);
  const onMove = ev => { col.style.width = Math.max(40, startW + (ev.clientX - startX)) + "px"; };
  const onUp = () => {
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("resizing");
  };
  document.body.classList.add("resizing");
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);
}

function startRowResize(tr, e) {
  e.preventDefault();
  const startY = e.clientY;
  const startH = tr.offsetHeight;
  const onMove = ev => { tr.style.height = Math.max(20, startH + (ev.clientY - startY)) + "px"; };
  const onUp = () => {
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("resizing");
  };
  document.body.classList.add("resizing");
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp);
}

function enableTableResize(table) {
  const headerRow = table.querySelector("thead tr");
  if (headerRow) {
    Array.from(headerRow.cells).forEach((th, ci) => {
      if (th.classList.contains("grid-op-col") || th.querySelector(".col-resize-handle")) return;
      const handle = document.createElement("div");
      handle.className = "col-resize-handle";
      handle.addEventListener("mousedown", e => startColResize(table, ci, e));
      th.appendChild(handle);
    });
  }
  table.querySelectorAll("tbody tr").forEach(tr => {
    const lastCell = tr.lastElementChild;
    if (!lastCell || lastCell.querySelector(".row-resize-handle")) return;
    lastCell.style.position = "relative";
    const handle = document.createElement("div");
    handle.className = "row-resize-handle";
    handle.addEventListener("mousedown", e => startRowResize(tr, e));
    lastCell.appendChild(handle);
  });
}

/* ---------- 表格查看 / 可编辑弹窗（表1、05、06、最终表） ---------- */
function waitForView(key, title, continueText, hint, canBack, withFilters) {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-view");
    document.getElementById("view-title").textContent = title;
    const hintEl = document.getElementById("view-hint");
    if (hint) { hintEl.textContent = hint; hintEl.hidden = false; }
    else hintEl.hidden = true;
    const body = document.getElementById("view-body");
    const tabs = document.getElementById("view-tabs");
    body.innerHTML = '<div style="padding:20px;color:var(--muted)">加载中…</div>';
    api("/api/view_sheet/" + key).then(res => {
      if (!res.success) {
        body.innerHTML = `<div style="padding:20px;color:var(--danger)">${esc(res.message || "加载失败")}</div>`;
        return;
      }
      const sheets = res.sheets || [];
      const show = html => { body.innerHTML = html; enhanceTable(body, withFilters); };
      if (sheets.length > 1) {
        tabs.hidden = false;
        tabs.innerHTML = "";
        sheets.forEach((s, i) => {
          const b = document.createElement("button");
          b.className = "sheet-tab" + (i === 0 ? " active" : "");
          b.textContent = s.name;
          b.onclick = () => {
            tabs.querySelectorAll(".sheet-tab").forEach(x => x.classList.remove("active"));
            b.classList.add("active");
            show(s.html);
          };
          tabs.appendChild(b);
        });
        show(sheets[0].html);
      } else {
        tabs.hidden = true;
        show(sheets[0] ? sheets[0].html : "");
      }
    }).catch(e => {
      body.innerHTML = `<div style="padding:20px;color:var(--danger)">加载失败: ${esc(String(e))}</div>`;
    });
    const cont = document.getElementById("btn-view-continue");
    const back = document.getElementById("btn-view-back");
    cont.textContent = continueText || "继续";
    back.hidden = !canBack;
    cont.onclick = () => { modal.hidden = true; resolve("next"); };
    back.onclick = () => { modal.hidden = true; resolve("back"); };
    modal.hidden = false;
  });
}

let editState = null;  // 当前可编辑表格的状态

function makeGridRow(rowVals, columns, editable, locked, pendingIdx) {
  const tr = document.createElement("tr");
  columns.forEach((c, ci) => {
    const td = document.createElement("td");
    const val = (rowVals && rowVals[ci] != null) ? String(rowVals[ci]) : "";
    const span = document.createElement("span");
    span.className = "cell-val" + (locked.has(ci) ? " locked" : "");
    if (editable.has(ci)) span.setAttribute("contenteditable", "true");
    span.textContent = val;
    td.appendChild(span);
    if (pendingIdx === ci && !val.trim()) {
      const badge = document.createElement("span");
      badge.className = "cell-pending";
      badge.textContent = "待确认";
      td.appendChild(badge);
    }
    tr.appendChild(td);
  });
  const op = document.createElement("td");
  op.className = "grid-op-col";
  op.style.position = "relative";
  const del = document.createElement("button");
  del.className = "btn tiny danger";
  del.textContent = "✕";
  del.title = "删除该行";
  del.onclick = () => tr.remove();
  op.appendChild(del);
  // 行高拖拽柄（新增行也自带）
  const rh = document.createElement("div");
  rh.className = "row-resize-handle";
  rh.addEventListener("mousedown", e => startRowResize(tr, e));
  op.appendChild(rh);
  tr.appendChild(op);
  return tr;
}

function renderEditGrid(data) {
  editState = data;
  const wrap = document.getElementById("edit-grid-wrap");
  const columns = data.columns || [];
  const editable = new Set(data.editable_idxs || []);
  const locked = new Set(data.locked_idxs || []);
  const pendingIdx = data.pending_idx != null ? data.pending_idx : -1;

  const tbl = document.createElement("table");
  const sticky = (columns[0] === "姓名") ? " sticky-name" : "";
  tbl.className = "sheet-table rv-grid" + sticky;
  let head = "<thead><tr>";
  columns.forEach(c => { head += `<th>${esc(c)}</th>`; });
  head += "<th class='grid-op-col'></th></tr></thead>";
  tbl.innerHTML = head;
  const tbody = document.createElement("tbody");
  (data.rows || []).forEach(rv => tbody.appendChild(makeGridRow(rv, columns, editable, locked, pendingIdx)));
  tbl.appendChild(tbody);
  wrap.innerHTML = "";
  wrap.appendChild(tbl);
  enhanceTable(wrap, true);  // 表头筛选 + 列宽拖拽（行高柄由 makeGridRow 自带）

  document.getElementById("edit-tip").textContent =
    pendingIdx >= 0
      ? "「加班时长确认-人工审核」列空值显示「待确认」，填上数字即为确认"
      : "仅「开工/收工项目名、简称」可修改，其余列为只读";
}

function saveEditGrid(key) {
  if (!editState) { toast("表格尚未加载", "error"); return Promise.resolve(null); }
  const rows = [];
  document.querySelectorAll("#edit-grid-wrap tbody tr").forEach(tr => {
    const vals = tr.querySelectorAll(".cell-val");
    rows.push(Array.from(vals).map(v => v.textContent.trim()));
  });
  setMask(true, "正在保存…");
  return api("/api/edit_save/" + key, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ columns: editState.columns, rows }),
  }).then(res => {
    if (res.success) { toast(res.message || "已保存", "success"); return true; }
    toast(res.message || "保存失败", "error");
    return null;
  }).catch(e => {
    toast("保存请求失败: " + e, "error");
    return null;
  }).finally(() => setMask(false));
}

function waitForEdit(key, errorKey, label, tableInfo, canBack) {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-edit");
    document.getElementById("edit-title").textContent = `生成 ${label} · 人工核对与修改`;

    const info = tableInfo || {};
    const normal = info.normal != null ? info.normal : "—";
    const abnormal = info.abnormal != null ? info.abnormal : "—";
    document.getElementById("edit-summary").innerHTML = `
      <span class="sum-chip normal">正常 ${normal} 人</span>
      <span class="sum-chip abnormal">异常 ${abnormal} 人</span>`;

    const links = document.getElementById("edit-links");
    links.innerHTML = "";
    if (info.error_doc) links.appendChild(downloadLink(info.error_doc, "📄 下载 核对信息错误文档"));
    if (info.table_file) links.appendChild(downloadLink(info.table_file, "📥 下载 " + label));

    const errDoc = document.getElementById("edit-error-doc");
    errDoc.innerHTML = '<div style="padding:10px;color:var(--muted)">加载核对信息错误文档…</div>';
    api("/api/view_sheet/" + errorKey).then(res => {
      errDoc.innerHTML = (res.success && res.sheets && res.sheets[0])
        ? res.sheets[0].html
        : `<div style="padding:10px;color:var(--muted)">${esc(res.message || "无错误文档")}</div>`;
    });

    const gridWrap = document.getElementById("edit-grid-wrap");
    gridWrap.innerHTML = '<div style="padding:20px;color:var(--muted)">加载表格…</div>';
    api("/api/edit_data/" + key).then(res => {
      if (!res.success) {
        gridWrap.innerHTML = `<div style="padding:20px;color:var(--danger)">${esc(res.message || "加载失败")}</div>`;
        return;
      }
      renderEditGrid(res);
    });

    const saveBtn = document.getElementById("btn-edit-save");
    const skipBtn = document.getElementById("btn-edit-skip");
    const regenBtn = document.getElementById("btn-edit-regen");
    const backBtn = document.getElementById("btn-edit-back");
    backBtn.hidden = !canBack;
    function cleanup() {
      modal.hidden = true;
      saveBtn.onclick = skipBtn.onclick = regenBtn.onclick = backBtn.onclick = null;
    }
    saveBtn.onclick = async () => {
      const ok = await saveEditGrid(key);
      if (!ok) return;  // 保存失败则不关闭
      cleanup();
      resolve("saved");
    };
    skipBtn.onclick = () => { cleanup(); resolve("next"); };
    regenBtn.onclick = () => { cleanup(); resolve("regen"); };
    backBtn.onclick = () => { cleanup(); resolve("back"); };
    modal.hidden = false;
  });
}

function downloadLink(path, text) {
  const a = document.createElement("a");
  a.href = "/api/download/" + encodeURIComponent(path);
  a.textContent = text;
  return a;
}

function addGridRow() {
  if (!editState) { toast("表格尚未加载", "error"); return; }
  const tbody = document.querySelector("#edit-grid-wrap tbody");
  if (!tbody) return;
  const editable = new Set(editState.editable_idxs || []);
  const locked = new Set(editState.locked_idxs || []);
  const row = makeGridRow(
    new Array(editState.columns.length).fill(""),
    editState.columns, editable, locked, editState.pending_idx
  );
  tbody.appendChild(row);
}

/* ---------- 下载 ---------- */
function initDownloads() {
  document.querySelectorAll(".js-download").forEach(btn => {
    btn.addEventListener("click", () => {
      const file = btn.closest(".dl-card").dataset.file;
      window.location.href = "/api/download/" + encodeURIComponent(file);
    });
  });
}

/* ---------- 清空文件夹（output/ + Target_Figure/） ---------- */
function initClearFolders() {
  const modal = document.getElementById("modal-clear");
  document.getElementById("btn-clear").addEventListener("click", () => {
    if (status && status.recognition_running) {
      toast("人脸识别运行中，请先停止再清空", "error");
      return;
    }
    modal.hidden = false;
  });
  document.getElementById("btn-clear-cancel").addEventListener("click", () => {
    modal.hidden = true;
  });
  document.getElementById("btn-clear-confirm").addEventListener("click", async () => {
    modal.hidden = true;
    setMask(true, "清空文件夹中…");
    try {
      const res = await api("/api/clear_folders", { method: "POST" });
      toast(res.message, res.success ? "success" : "error");
    } catch (e) {
      toast("清空失败: " + e, "error");
    } finally {
      setMask(false);
    }
    await renderStatus();
  });
}

/* ---------- 初始化 ---------- */
document.addEventListener("DOMContentLoaded", () => {
  initStepper();
  initDownloads();
  initClearFolders();
  setupDropzone("dz-ref", "file-ref", "ref");
  setupDropzone("dz-target", "file-target", "target");

  document.getElementById("btn-recognize").addEventListener("click", startRecognition);
  document.getElementById("btn-refresh").addEventListener("click", renderStatus);
  document.getElementById("btn-generate").addEventListener("click", startGeneration);
  document.getElementById("btn-skip-final").addEventListener("click", () => gotoStep(4));
  document.getElementById("btn-next-1").addEventListener("click", () => gotoStep(2));
  document.getElementById("btn-next-2").addEventListener("click", () => gotoStep(3));

  document.getElementById("btn-review-start").addEventListener("click", openReview);
  document.getElementById("btn-review-skip-all").addEventListener("click", skipAllReview);
  document.getElementById("btn-rv-confirm").addEventListener("click", confirmReviewPhoto);
  document.getElementById("btn-rv-skip").addEventListener("click", skipReviewPhoto);

  document.getElementById("btn-grid-add").addEventListener("click", addGridRow);

  renderStatus();
  setInterval(renderStatus, 15000); // 周期刷新
});
