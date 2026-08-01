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
  { file: "08_表8工程考勤表.xlsx",    name: "表8 工程考勤表" },
  { file: "09_表9外协考勤表1.xlsx",   name: "表9 外协考勤表1" },
  { file: "10_表10外协考勤表2.xlsx",  name: "表10 外协考勤表2" },
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
  const clearBtn = document.getElementById("btn-clear");
  if (clearBtn) clearBtn.disabled = status.recognition_running;
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

  // 第1步 → 第2步：识别已完成，或 00_ 结果已存在
  const recognitionDone = status.recognition_done || !!tables["table3"];
  const next1 = document.getElementById("btn-next-1");
  if (next1) next1.disabled = !recognitionDone;

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

/* ---------- 生成表格（Step3） ---------- */
async function startGeneration() {
  const genBtn = document.getElementById("btn-generate");
  const consoleEl = document.getElementById("gen-console");
  genBtn.disabled = true;
  consoleEl.textContent = "";

  resetChips();
  for (const key of SEQUENCE) {
    setChipState(key, "running");
    let res;
    try {
      res = await api("/api/run_table/" + key, { method: "POST" });
    } catch (e) {
      setChipState(key, "failed");
      appendLog(consoleEl, "【" + key + "】请求失败: " + e);
      toast("网络错误", "error");
      break;
    }

    if (res.output) appendLog(consoleEl, res.output);
    setChipState(key, res.success ? "done" : "failed");

    if (res.confirm_required) {
      const target = key === "table5" ? "05" : "06";
      let choice;
      let reRunFailed = false;
      do {
        choice = await waitForConfirm(key, res.table_info, target);
        if (!choice.confirmed) {
          setChipState(key, "running");
          res = await api("/api/run_table/" + key, { method: "POST" });
          if (res.output) appendLog(consoleEl, res.output);
          setChipState(key, res.success ? "done" : "failed");
          if (!res.success) { reRunFailed = true; break; }  // 重新生成失败，退出确认循环
        }
      } while (!choice.confirmed);
      if (reRunFailed) {
        toast(`❌ ${res.label || key} 重新生成失败，请查看日志`, "error");
        break;
      }
      await api("/api/confirm/" + target, {
        method: "POST",
        body: choice.file ? (() => { const fd = new FormData(); fd.append("file", choice.file); return fd; })() : undefined,
      });
    }

    if (!res.success) {
      toast(`❌ ${res.label || key} 生成失败，请查看日志`, "error");
      break;
    }
  }

  genBtn.disabled = false;
  await renderStatus();

  // 全部成功 → 进入完成页
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

/* ---------- 确认弹窗（05/06） ---------- */
function waitForConfirm(key, tableInfo, target) {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-confirm");
    const title = document.getElementById("modal-title");
    const sub = document.getElementById("modal-sub");
    const summary = document.getElementById("modal-summary");
    const links = document.getElementById("modal-links");
    const fileInput = document.getElementById("modal-file");
    fileInput.value = "";

    const is5 = target === "05";
    const tableName = is5 ? "05_该日出工人员表" : "06_全体人员出工情况表";
    title.textContent = `生成 ${tableName} · 人工核对`;
    sub.textContent = "请核对下方信息，可下载/回传修正表后继续";

    summary.innerHTML = "";
    const info = tableInfo || {};
    const normal = info.normal != null ? info.normal : "—";
    const abnormal = info.abnormal != null ? info.abnormal : "—";
    summary.innerHTML = `
      <span class="sum-chip normal">正常 ${normal} 人</span>
      <span class="sum-chip abnormal">异常 ${abnormal} 人</span>
    `;

    links.innerHTML = "";
    if (info.error_doc) {
      links.appendChild(downloadLink(info.error_doc, "📄 下载 核对信息错误文档"));
    }
    if (info.table_file) {
      links.appendChild(downloadLink(info.table_file, "📥 下载 " + tableName));
    }

    const confirmBtn = document.getElementById("btn-confirm");
    const regenBtn = document.getElementById("btn-regen");
    function cleanup() {
      modal.hidden = true;
      confirmBtn.onclick = null;
      regenBtn.onclick = null;
    }
    confirmBtn.onclick = () => {
      cleanup();
      resolve({ confirmed: true, file: fileInput.files[0] || null });
    };
    regenBtn.onclick = () => {
      cleanup();
      resolve({ confirmed: false });
    };
    modal.hidden = false;
  });
}

function downloadLink(path, text) {
  const a = document.createElement("a");
  a.href = "/api/download/" + encodeURIComponent(path);
  a.textContent = text;
  return a;
}

/* ---------- 下载 ---------- */
function initDownloads() {
  document.querySelectorAll(".js-download").forEach(btn => {
    btn.addEventListener("click", () => {
      const file = btn.closest(".dl-card").dataset.file;
      window.location.href = "/api/download/" + encodeURIComponent(file);
    });
  });
  document.getElementById("btn-zip").addEventListener("click", () => {
    window.location.href = "/api/download_zip";
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

  renderStatus();
  setInterval(renderStatus, 15000); // 周期刷新
});
