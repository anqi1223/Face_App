/* ============================================================
   人脸识别考勤系统 - Web 前端逻辑
   ============================================================ */

"use strict";

/* ---------- 元数据 ---------- */
const INPUT_FILES = [
  { key: "photo_ledger",        name: "照片台账表",            file: "01_照片台账表.xlsx" },
  { key: "work_plan",           name: "工作安排表",            file: "02_工作安排表.xlsx" },
  { key: "person_class",        name: "人员分类表",            file: "03_人员分类表.xlsx" },
  { key: "attendance_template", name: "工程与外协考勤表模板",   file: "05_工程与外协考勤表模板.xlsx" },
];

// 生成 000_项目信息表 的三张输入表
const PROJINFO_FILES = [
  { key: "project_stat",   name: "04_工地项目统计",   file: "04_工地项目统计.xlsx" },
  { key: "weekly_arrange", name: "06_周工作安排表",   file: "06_周工作安排表.xlsx" },
  { key: "weekly_plan",    name: "07_周工作计划表",   file: "07_周工作计划表.xlsx" },
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
  renderProjinfoList();
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

function renderProjinfoList() {
  const list = document.getElementById("projinfo-list");
  if (!list) return;
  list.innerHTML = "";
  const pfiles = status.projinfo_files || {};
  for (const f of PROJINFO_FILES) {
    const ok = !!pfiles[f.key];
    const row = document.createElement("div");
    row.className = "input-row";
    row.innerHTML = `
      <span class="name">${f.name}</span>
      <span class="badge ${ok ? "ok" : "missing"}">${ok ? "已上传 ✓" : "未上传"}</span>
      <span class="desc">input/${f.file}</span>
      <button class="btn small ${ok ? "ghost" : "primary"}" data-pi-upload-key="${f.key}">${ok ? "重新上传" : "上传"}</button>
    `;
    row.querySelector("[data-pi-upload-key]").addEventListener("click", () => pickAndUpload(f));
    list.appendChild(row);
  }
  const genBtn = document.getElementById("btn-projinfo-gen");
  if (genBtn) genBtn.disabled = !PROJINFO_FILES.every(f => pfiles[f.key]);
  const tip = document.getElementById("projinfo-tip");
  if (tip) {
    tip.innerHTML = status.projinfo_exists
      ? "✅ <b>output/000_项目信息表</b> 已生成（本周内无需重新生成，可直接下一步；需要重跑时再点「生成项目信息表」）"
      : "⚠️ 尚未生成 000_项目信息表，请先上传三张表并点击「生成项目信息表」";
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

  // 第1步 → 第2步：识别已完成
  const recognitionDone = status.recognition_done || !!tables["table3"];
  const next1 = document.getElementById("btn-next-1");
  if (next1) next1.disabled = !recognitionDone;

  // 第2步 → 第3步：01/02/03/05 输入表齐全（人工复核需基于 01_照片台账表）
  const next2 = document.getElementById("btn-next-2");
  if (next2) next2.disabled = !allInputs;

  // 第3步 → 第4步：复核完成/跳过（存在未识别人脸时必须先处理）
  const reviewBlocking = status.review_pending && !reviewHandled;
  const next3 = document.getElementById("btn-next-3");
  if (next3) next3.disabled = reviewBlocking;

  // 第4步 → 第5步：项目信息表可跳过，下一步始终可用
  const next4 = document.getElementById("btn-next-4");
  if (next4) next4.disabled = false;

  // 第5步开始生成：输入表齐全
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
  for (let i = 1; i <= 6; i++) {
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
      // 复核改为独立第3步，由「开始人工复核」按钮触发
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

/* 让序号标签可拖动挪开（挡住脸时移开看脸） */
function makeDraggable(el) {
  el.addEventListener("mousedown", e => {
    e.preventDefault();
    const startX = e.clientX, startY = e.clientY;
    const origLeft = parseFloat(el.style.left) || 0;
    const origTop = parseFloat(el.style.top) || 0;
    el.classList.add("dragging");
    const onMove = ev => {
      el.style.left = (origLeft + ev.clientX - startX) + "px";
      el.style.top = (origTop + ev.clientY - startY) + "px";
    };
    const onUp = () => {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      el.classList.remove("dragging");
    };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
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
    // 以显示尺寸换算像素坐标（wrap 内图片宽度=100%）
    const ww = wrap.clientWidth || nw;
    const wh = wrap.clientHeight || nh;
    (p.unknown_faces || []).forEach((f, i) => {
      if (!f.box || f.box.length < 4) return;
      const [x1, y1, x2, y2] = f.box;
      // 不画红框（图片小时会遮脸），在脸部框正下方标序号；可拖动挪开看脸
      const num = document.createElement("span");
      num.className = "rv-box-num";
      num.textContent = i + 1;
      num.style.left = (((x1 + x2) / 2) / nw * ww) + "px";
      num.style.top = (y2 / nh * wh) + "px";
      makeDraggable(num);
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
    hint: "可直接框选表格内容复制到 Excel（合并单元格格式已保留）；出工情况列可修改并自动保存",
    withFilters: false, finalMode: true },
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
        screen.hint, pos > 0, !!screen.withFilters, !!screen.finalMode);
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
    gotoStep(6);
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

/* 取单元格值：优先取 .cell-val（可编辑格的实际值，排除"待确认"徽标等） */
function cellVal(cell) {
  const val = cell.querySelector(".cell-val");
  return val ? val.textContent.trim() : (cell.textContent || "").trim();
}

function showFilterMenu(table, ci, btn) {
  closeFilterMenu();
  const menu = document.createElement("div");
  menu.className = "filter-menu";
  const values = new Set();
  let hasEmpty = false;
  table.querySelectorAll("tbody tr").forEach(tr => {
    const cell = tr.cells[ci];
    if (!cell) return;
    const v = cellVal(cell);
    if (v) values.add(v);
    else hasEmpty = true;
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
  if (hasEmpty) add("（空值）", "__EMPTY__");
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
    const v = cell ? cellVal(cell) : "";
    let show;
    if (value === "") show = true;                      // 全部
    else if (value === "__EMPTY__") show = (v === "");  // 空值
    else show = (v === value);                          // 具体值
    tr.style.display = show ? "" : "none";
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
let viewSearchPos = -1;  // 最终表搜索位置
let viewCurrentSheet = "";  // 最终表当前子表名

function waitForView(key, title, continueText, hint, canBack, withFilters, finalMode) {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-view");
    document.getElementById("view-title").textContent = title;
    const hintEl = document.getElementById("view-hint");
    if (hint) { hintEl.textContent = hint; hintEl.hidden = false; }
    else hintEl.hidden = true;
    const body = document.getElementById("view-body");
    const tabs = document.getElementById("view-tabs");
    const searchBar = document.getElementById("view-search-bar");
    searchBar.hidden = !finalMode;  // 仅最终考勤表显示搜索/编辑栏
    body.innerHTML = '<div style="padding:20px;color:var(--muted)">加载中…</div>';
    api("/api/view_sheet/" + key).then(res => {
      if (!res.success) {
        body.innerHTML = `<div style="padding:20px;color:var(--danger)">${esc(res.message || "加载失败")}</div>`;
        return;
      }
      const sheets = res.sheets || [];
      viewSearchPos = -1;
      const show = (html, sheetName) => {
        body.innerHTML = html;
        enhanceTable(body, withFilters);
        viewCurrentSheet = sheetName || "";
        if (finalMode) enableFinalEdit(body);
      };
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
            show(s.html, s.name);
          };
          tabs.appendChild(b);
        });
        show(sheets[0].html, sheets[0].name);
      } else {
        tabs.hidden = true;
        show(sheets[0] ? sheets[0].html : "", sheets[0] ? sheets[0].name : "");
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

/* 最终考勤表：出工情况列可编辑，失焦自动保存 */
function enableFinalEdit(container) {
  container.querySelectorAll(".sheet-table").forEach(table => {
    const headerRow = table.querySelector("thead tr") || table.querySelector("tr:first-child");
    if (!headerRow) return;
    let colC = null;
    Array.from(headerRow.cells).forEach(cell => {
      if (cell.textContent.includes("出工情况")) {
        colC = parseInt(cell.getAttribute("data-c"), 10);
      }
    });
    if (!colC) return;
    table.querySelectorAll("td[data-c]").forEach(cell => {
      const c = parseInt(cell.getAttribute("data-c"), 10);
      const r = parseInt(cell.getAttribute("data-r"), 10);
      if (c !== colC || r <= 1) return;
      cell.setAttribute("contenteditable", "true");
      cell.classList.add("final-editable");
      cell.addEventListener("blur", () => {
        const val = cell.textContent.trim();
        api("/api/edit_final_cell", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sheet: viewCurrentSheet, changes: [{ r, c, value: val }] }),
        }).then(res => {
          if (res && res.success) toast(res.message || "已保存", "success");
          else toast((res && res.message) || "保存失败", "error");
        }).catch(() => toast("保存失败", "error"));
      });
    });
  });
}

function viewSearchNext() {
  const find = document.getElementById("vs-find").value;
  if (!find) { toast("请输入搜索内容", "error"); return; }
  const cells = Array.from(document.querySelectorAll("#view-body td[data-r], #view-body th[data-r]"));
  document.querySelectorAll("#view-body .vs-hit").forEach(c => c.classList.remove("vs-hit"));
  const start = viewSearchPos;
  for (let i = start + 1; i < cells.length; i++) {
    if (cells[i].textContent.includes(find)) {
      viewSearchPos = i;
      cells[i].classList.add("vs-hit");
      cells[i].scrollIntoView({ block: "center", behavior: "smooth" });
      return;
    }
  }
  for (let i = 0; i <= start; i++) {
    if (cells[i].textContent.includes(find)) {
      viewSearchPos = i;
      cells[i].classList.add("vs-hit");
      cells[i].scrollIntoView({ block: "center", behavior: "smooth" });
      toast("已循环到开头", "");
      return;
    }
  }
  viewSearchPos = -1;
  toast("未找到匹配", "");
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

/* 修改开工/收工项目名后，按 000_项目信息表 自动匹配对应的项目简称 */
async function autoMatchShort(project, targetCell) {
  if (!project) { targetCell.textContent = ""; return; }
  try {
    const res = await api("/api/match_short", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ project }),
    });
    if (res && res.success && res.short) targetCell.textContent = res.short;
  } catch (e) { /* 网络失败静默，保持原值 */ }
}

function wireAutoShort(tbody, nameIdx, shortIdx) {
  tbody.querySelectorAll("tr").forEach(tr => {
    const nameCell = tr.cells[nameIdx] ? tr.cells[nameIdx].querySelector(".cell-val") : null;
    const shortCell = tr.cells[shortIdx] ? tr.cells[shortIdx].querySelector(".cell-val") : null;
    if (nameCell && shortCell) {
      nameCell.addEventListener("blur", () => autoMatchShort(nameCell.textContent.trim(), shortCell));
    }
  });
}

function renderEditGrid(data) {
  editState = data;
  frCurrent = null;  // 表格重建后重置查找位置
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

  // 修改开工/收工项目名后，自动按 000_项目信息表 匹配对应的项目简称
  const kgIdx = columns.indexOf("开工项目名");
  const kgShortIdx = columns.indexOf("开工项目简称");
  const sgIdx = columns.indexOf("收工项目名");
  const sgShortIdx = columns.indexOf("收工项目简称");
  if (kgIdx >= 0 && kgShortIdx >= 0) wireAutoShort(tbody, kgIdx, kgShortIdx);
  if (sgIdx >= 0 && sgShortIdx >= 0) wireAutoShort(tbody, sgIdx, sgShortIdx);

  // 出现在核对信息错误文档中的姓名行 → 高亮底色提示（含1字之差的相近姓名）
  const hList = data.highlight_names || [];
  const hSet = new Set(hList);
  const nameMatches = name => {
    if (!name || hSet.has(name)) return !!name;
    return hList.some(h =>
      h.length === name.length && h.length >= 2 &&
      [...h].filter((ch, i) => ch !== name[i]).length === 1
    );
  };
  tbody.querySelectorAll("tr").forEach(tr => {
    const first = tr.cells[0] ? tr.cells[0].textContent.trim() : "";
    if (nameMatches(first)) tr.classList.add("row-abnormal");
  });

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

/* ---------- 查找 / 替换（05/06 编辑表） ---------- */
let frCurrent = null;  // {cell, pos} 当前查找位置

function _frCells() {
  return Array.from(document.querySelectorAll("#edit-grid-wrap .cell-val"));
}

function _frHighlight(cell) {
  _frCells().forEach(c => c.classList.remove("fr-hit"));
  cell.classList.add("fr-hit");
  cell.scrollIntoView({ block: "center", behavior: "smooth" });
  cell.focus();
}

function frFindNext() {
  const find = document.getElementById("fr-find").value;
  if (!find) { toast("请输入查找内容", "error"); return; }
  const cells = _frCells();
  // 1) 同格内继续找
  if (frCurrent && frCurrent.cell && cells.includes(frCurrent.cell)) {
    const idx = frCurrent.cell.textContent.indexOf(find, frCurrent.pos + find.length);
    if (idx >= 0) { frCurrent = { cell: frCurrent.cell, pos: idx }; _frHighlight(frCurrent.cell); return; }
  }
  // 2) 后续格
  const start = frCurrent ? cells.indexOf(frCurrent.cell) + 1 : 0;
  for (let i = start; i < cells.length; i++) {
    const idx = cells[i].textContent.indexOf(find);
    if (idx >= 0) { frCurrent = { cell: cells[i], pos: idx }; _frHighlight(cells[i]); return; }
  }
  // 3) 循环到开头
  for (let i = 0; i < start; i++) {
    const idx = cells[i].textContent.indexOf(find);
    if (idx >= 0) { frCurrent = { cell: cells[i], pos: idx }; _frHighlight(cells[i]); toast("已循环到开头", ""); return; }
  }
  frCurrent = null;
  toast("未找到匹配", "");
}

function frReplaceOne() {
  const find = document.getElementById("fr-find").value;
  const repl = document.getElementById("fr-replace").value;
  if (!find) { toast("请输入查找内容", "error"); return; }
  if (!frCurrent || !frCurrent.cell) { toast("请先点「查找下一个」", "error"); return; }
  const cell = frCurrent.cell;
  const text = cell.textContent;
  if (frCurrent.pos >= 0 && text.startsWith(find, frCurrent.pos)) {
    cell.textContent = text.slice(0, frCurrent.pos) + repl + text.slice(frCurrent.pos + find.length);
    frCurrent = { cell, pos: frCurrent.pos + repl.length };
    _frHighlight(cell);
    toast("已替换", "success");
  } else {
    toast("当前内容已变化，请重新查找", "error");
    frCurrent = null;
  }
}

function frReplaceAll() {
  const find = document.getElementById("fr-find").value;
  const repl = document.getElementById("fr-replace").value;
  if (!find) { toast("请输入查找内容", "error"); return; }
  let count = 0;
  _frCells().forEach(cell => {
    const parts = cell.textContent.split(find);
    if (parts.length > 1) {
      cell.textContent = parts.join(repl);
      count += parts.length - 1;
    }
  });
  frCurrent = null;
  _frCells().forEach(c => c.classList.remove("fr-hit"));
  toast(count > 0 ? `已替换 ${count} 处` : "未找到可替换内容", count > 0 ? "success" : "");
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

/* ---------- 生成 000_项目信息表 ---------- */
async function startProjinfoGen() {
  const genBtn = document.getElementById("btn-projinfo-gen");
  genBtn.disabled = true;
  setMask(true, "正在分析周安排表…");
  let plan;
  try {
    const res = await api("/api/projinfo/plan", { method: "POST" });
    if (!res.success) { toast(res.message || "生成失败", "error"); return; }
    plan = res;
  } catch (e) {
    toast("请求失败: " + e, "error");
    return;
  } finally {
    setMask(false);
    genBtn.disabled = false;
  }
  let selections = {};
  if (plan.review && plan.review.length) {
    selections = await showProjinfoModal(plan);
    if (!selections) return;  // 用户取消
  }
  await confirmProjinfo(selections);
}

async function confirmProjinfo(selections) {
  setMask(true, "正在生成 000_项目信息表…");
  try {
    const res = await api("/api/projinfo/confirm", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ selections: selections || {} }),
    });
    if (res.success) toast(res.message || "已生成", "success");
    else toast(res.message || "生成失败", "error");
  } catch (e) {
    toast("请求失败: " + e, "error");
  } finally {
    setMask(false);
    await renderStatus();
  }
}

function showProjinfoModal(plan) {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-projinfo");
    document.getElementById("pi-period").textContent =
      `日期范围：${plan.period} ｜ 自动 ${(plan.auto || []).length} 项，待确认 ${(plan.review || []).length} 项`;
    document.getElementById("pi-hint").textContent =
      "请逐条确认「出勤统计简称」与「工作安排简称」（可直接选或手动改）；「新增行」表示在 04 中无匹配，将按建议新增到 000。";
    const wrap = document.getElementById("pi-grid-wrap");
    const reviews = plan.review || [];

    const tbl = document.createElement("table");
    tbl.className = "sheet-table";
    tbl.innerHTML = `<thead><tr>
      <th>06项目名称</th><th>工作安排简称</th><th>出勤统计简称</th><th>输出项目名称</th><th>说明</th>
    </tr></thead>`;
    const tbody = document.createElement("tbody");
    const modeLabel = { fuzzy: "相近匹配", none: "无直接匹配", new: "新增行" };

    reviews.forEach(r => {
      const tr = document.createElement("tr");
      const mkTd = txt => { const td = document.createElement("td"); td.textContent = txt || ""; return td; };
      tr.appendChild(mkTd(r.项目名称));
      const wsTd = document.createElement("td");
      const wsIn = document.createElement("input");
      wsIn.className = "pi-input";
      wsIn.value = r.work_short || "";
      wsIn.dataset.rid = r.rid;
      wsIn.dataset.field = "work_short";
      wsIn.setAttribute("list", "pi-work-short-list");
      wsTd.appendChild(wsIn);
      tr.appendChild(wsTd);
      const ssTd = document.createElement("td");
      const ssIn = document.createElement("input");
      ssIn.className = "pi-input";
      ssIn.value = r.attendance_short || "";
      ssIn.dataset.rid = r.rid;
      ssIn.dataset.field = "attendance_short";
      ssIn.setAttribute("list", "pi-short-list");
      ssTd.appendChild(ssIn);
      tr.appendChild(ssTd);
      const pnTd = document.createElement("td");
      const pnIn = document.createElement("input");
      pnIn.className = "pi-input";
      // 输出项目名称：相近/直接匹配 → 04 原有项目名；无匹配/新增 → 建议的新名
      pnIn.value = r.suggest_name || r.项目名称 || "";
      pnIn.dataset.rid = r.rid;
      pnIn.dataset.field = "project_name";
      pnTd.appendChild(pnIn);
      tr.appendChild(pnTd);
      // 说明列：模式标签 + 悬停备注（原因）
      const tipTd = document.createElement("td");
      tipTd.textContent = modeLabel[r.mode] || r.mode;
      if (r.reason) {
        const tip = document.createElement("span");
        tip.className = "tip-icon";
        tip.setAttribute("data-tip", r.reason);
        tip.textContent = "  ⓘ";
        tipTd.appendChild(tip);
      }
      tr.appendChild(tipTd);
      tbody.appendChild(tr);
    });

    // 自动匹配项追加在底部（只读，供一起查看）
    const autos = plan.auto || [];
    if (autos.length) {
      const sep = document.createElement("tr");
      sep.className = "pi-auto-sep";
      const sepTd = document.createElement("td");
      sepTd.colSpan = 5;
      sepTd.textContent = `✅ 自动匹配项（无需确认，共 ${autos.length} 项）`;
      sep.appendChild(sepTd);
      tbody.appendChild(sep);
      autos.forEach(a => {
        const tr = document.createElement("tr");
        tr.className = "pi-auto-row";
        const mkTd = txt => { const td = document.createElement("td"); td.textContent = txt || ""; return td; };
        tr.appendChild(mkTd(a.项目名称));
        tr.appendChild(mkTd(a.工作安排简称));
        tr.appendChild(mkTd(a.出勤统计简称));
        tr.appendChild(mkTd(a.项目名称));
        tr.appendChild(mkTd("自动匹配"));
        tbody.appendChild(tr);
      });
    }

    tbl.appendChild(tbody);
    wrap.innerHTML = "";
    wrap.appendChild(tbl);

    // datalist 候选
    const buildList = (id, values) => {
      let dl = document.getElementById(id);
      if (!dl) { dl = document.createElement("datalist"); dl.id = id; document.body.appendChild(dl); }
      dl.innerHTML = "";
      const seen = new Set();
      (values || []).forEach(v => { if (!seen.has(v)) { seen.add(v); const o = document.createElement("option"); o.value = v; dl.appendChild(o); } });
    };
    const wsVals = [], ssVals = [];
    reviews.forEach(r => {
      (r.work_short_options || []).forEach(v => wsVals.push(v));
      (r.short_options || []).forEach(v => ssVals.push(v));
    });
    buildList("pi-work-short-list", wsVals);
    buildList("pi-short-list", ssVals);

    const confirmBtn = document.getElementById("btn-pi-confirm");
    const cancelBtn = document.getElementById("btn-pi-cancel");
    function cleanup() { modal.hidden = true; confirmBtn.onclick = cancelBtn.onclick = null; }
    confirmBtn.onclick = () => {
      const selections = {};
      wrap.querySelectorAll(".pi-input").forEach(inp => {
        const rid = inp.dataset.rid, field = inp.dataset.field;
        if (!selections[rid]) selections[rid] = {};
        selections[rid][field] = inp.value.trim();
      });
      cleanup();
      resolve(selections);
    };
    cancelBtn.onclick = () => { cleanup(); resolve(null); };
    modal.hidden = false;
  });
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
  document.getElementById("btn-skip-final").addEventListener("click", () => gotoStep(6));
  document.getElementById("btn-next-1").addEventListener("click", () => gotoStep(2));
  document.getElementById("btn-next-2").addEventListener("click", () => gotoStep(3));
  document.getElementById("btn-next-3").addEventListener("click", () => gotoStep(4));
  document.getElementById("btn-next-4").addEventListener("click", () => gotoStep(5));
  document.getElementById("btn-projinfo-gen").addEventListener("click", startProjinfoGen);
  document.getElementById("btn-projinfo-status").addEventListener("click", renderStatus);

  document.getElementById("btn-review-start").addEventListener("click", openReview);
  document.getElementById("btn-review-skip-all").addEventListener("click", skipAllReview);
  document.getElementById("btn-rv-confirm").addEventListener("click", confirmReviewPhoto);
  document.getElementById("btn-rv-skip").addEventListener("click", skipReviewPhoto);

  document.getElementById("btn-grid-add").addEventListener("click", addGridRow);

  document.getElementById("fr-next").addEventListener("click", frFindNext);
  document.getElementById("fr-replace-one").addEventListener("click", frReplaceOne);
  document.getElementById("fr-replace-all").addEventListener("click", frReplaceAll);
  document.getElementById("fr-find").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); frFindNext(); } });
  document.getElementById("fr-replace").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); frReplaceOne(); } });

  document.getElementById("vs-next").addEventListener("click", viewSearchNext);
  document.getElementById("vs-find").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); viewSearchNext(); } });

  renderStatus();
  setInterval(renderStatus, 15000); // 周期刷新
});
