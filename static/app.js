// ===== 상태 =====
let OPTIONS = { purpose: [], payment: [], limits: {} };
let CONFIG = { expense_template: "비용청구양식.xlsm", overtime_template: "연장근무(수당)신청서_양식.xlsx" };
let rcptFiles = [];           // File[]
let analyzedNames = new Set();
let rcptRows = [];
let urlCache = {};
let title = "비용";
let attFile = null;
let otRows = [];
let activeTab = "expense";
let expenseTpl = null;        // 커스텀 양식 File (없으면 기본)
let overtimeTpl = null;
// AI 연결: 기본은 '사내'(서버 .env). 사용자가 고급 설정에서 OpenAI/직접 로컬로 덮어쓸 수 있음.
const OPENAI_MODELS = ["gpt-5.3", "gpt-5.3-mini", "gpt-4.1", "gpt-4.1-mini", "gpt-4o", "gpt-4o-mini"];
let AI = { mode: "default", oaModel: "gpt-4o", cuBase: "", cuModel: "", key: "" };
let lastBulk = { purpose: "", payment: "" };  // 일괄 채우기로 마지막에 적용한 값 (표시용)

// ===== 공통 =====
const $ = (id) => document.getElementById(id);
const esc = (s) => (s == null ? "" : String(s).replace(/"/g, "&quot;"));
const toInt = (v) => { if (v == null || v === "") return null; const n = parseInt(String(v).replace(/,/g, "").trim(), 10); return isNaN(n) ? null : n; };
const fmt = (n) => (n == null || n === "" ? "" : Number(n).toLocaleString());
const isWelfare = () => title === "복지비";
function dateKey(s) { const m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(esc(s)); return m ? `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` : "9999-99-99"; }
function note(el, type, msg) { el.innerHTML = msg ? `<div class="note ${type}">${msg}</div>` : ""; }

function switchTab(which) {
  activeTab = which;
  $("tab-expense").classList.toggle("active", which === "expense");
  $("tab-overtime").classList.toggle("active", which === "overtime");
  $("panel-expense").classList.toggle("hidden", which !== "expense");
  $("panel-overtime").classList.toggle("hidden", which !== "overtime");
  updateTplName();
}

async function download(resp, fallbackName) {
  const cd = resp.headers.get("Content-Disposition") || "";
  let name = fallbackName;
  const m = /filename\*=UTF-8''([^;]+)/.exec(cd);
  if (m) name = decodeURIComponent(m[1]);
  const count = resp.headers.get("X-Record-Count");
  const blob = await resp.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; document.body.appendChild(a); a.click();
  a.remove(); URL.revokeObjectURL(url);
  return { name, count };
}
async function errText(resp) { try { const j = await resp.json(); return j.detail || JSON.stringify(j); } catch { return `오류 ${resp.status}`; } }

// ===== 사이드바: 양식 / AI 설정 =====
function updateTplName() {
  const custom = activeTab === "expense" ? expenseTpl : overtimeTpl;
  const def = activeTab === "expense" ? CONFIG.expense_template : CONFIG.overtime_template;
  const st = $("sbTplName").parentElement;
  if (custom) { st.innerHTML = `<span class="ok">📎 변경된 양식</span><br/><b id="sbTplName">${esc(custom.name)}</b>`; }
  else { st.innerHTML = `<span class="ok">✓ 기본 양식 사용 중</span><br/><b id="sbTplName">${esc(def)}</b>`; }
}
async function onTplChange(file) {
  if (!file) return;
  if (activeTab === "expense") { expenseTpl = file; await fetchExpenseOptions(); }
  else { overtimeTpl = file; }
  updateTplName();
}

// ===== AI 연결 (고급) =====
function setAiMode(btn) {
  document.querySelectorAll("#aiMode button").forEach((b) => b.classList.remove("on"));
  btn.classList.add("on");
  AI.mode = btn.dataset.v;
  applyAiMode(); saveAi();
}
// 상단 상태 문구를 현재 선택된 모델에 맞춰 갱신한다.
function updateAiStatus() {
  const st = $("aiStatus"), sub = $("aiStatusSub"), reset = $("aiResetTop");
  if (AI.mode === "openai") {
    st.textContent = `✓ OpenAI · ${AI.oaModel} 사용 중`;
    sub.textContent = "이 브라우저에서만 적용됩니다.";
    reset.classList.remove("hidden");
  } else if (AI.mode === "custom") {
    st.textContent = `✓ 직접 연결 · ${AI.cuModel || "모델명 미입력"} 사용 중`;
    sub.textContent = "이 브라우저에서만 적용됩니다.";
    reset.classList.remove("hidden");
  } else {
    st.textContent = "✓ 사내 AI 모델 연결됨";
    sub.textContent = "";
    reset.classList.add("hidden");
  }
}
// 사내 기본 모델로 되돌린다. (override 해제 → 서버 .env 설정 사용)
function resetAiDefault() {
  AI.mode = "default";
  document.querySelectorAll("#aiMode button").forEach((b) => b.classList.remove("on"));
  applyAiMode(); saveAi(); updateAiStatus();
}
function applyAiMode() {
  $("openaiFields").classList.toggle("hidden", AI.mode !== "openai");
  $("customFields").classList.toggle("hidden", AI.mode !== "custom");
  $("keyField").classList.toggle("hidden", AI.mode === "default");
  const ar = $("aiReset"); if (ar) ar.classList.toggle("hidden", AI.mode === "default");
}
function saveAi() {
  AI = { mode: AI.mode, oaModel: $("oaModel").value, cuBase: $("cuBase").value.trim(), cuModel: $("cuModel").value.trim(), key: $("aiKey").value.trim() };
  localStorage.setItem("aiConn", JSON.stringify(AI));
  updateAiStatus();
  const s = $("aiSaved"); s.textContent = "저장됨 ✓"; setTimeout(() => { s.textContent = ""; }, 1500);
}
function loadAi() {
  $("oaModel").innerHTML = OPENAI_MODELS.map((m) => `<option>${esc(m)}</option>`).join("");
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem("aiConn") || "{}"); } catch {}
  AI = { mode: saved.mode || "default", oaModel: saved.oaModel || "gpt-4o", cuBase: saved.cuBase || "", cuModel: saved.cuModel || "", key: saved.key || "" };
  document.querySelectorAll("#aiMode button").forEach((b) => b.classList.toggle("on", b.dataset.v === AI.mode));
  $("oaModel").value = AI.oaModel; $("cuBase").value = AI.cuBase; $("cuModel").value = AI.cuModel; $("aiKey").value = AI.key;
  applyAiMode(); updateAiStatus();
}
function aiOverride(fd) {
  if (AI.mode === "openai") { fd.append("provider", "OpenAI"); fd.append("model", AI.oaModel); if (AI.key) fd.append("api_key", AI.key); }
  else if (AI.mode === "custom") { fd.append("provider", "로컬 서버"); fd.append("base_url", AI.cuBase); fd.append("model", AI.cuModel); if (AI.key) fd.append("api_key", AI.key); }
  // default: 아무것도 안 보냄 → 서버 .env 사용
}

// ===== 기초정보 =====
function setTitle(btn) {
  document.querySelectorAll("#titleSeg button").forEach((b) => b.classList.remove("on"));
  btn.classList.add("on"); title = btn.dataset.v;
  $("welfareWrap").classList.toggle("hidden", !isWelfare()); renderReview();
}

// ===== 드롭존 =====
function setupDrop(zone, input, onFiles) {
  zone.onclick = () => input.click();
  input.onchange = () => { onFiles([...input.files]); input.value = ""; };
  ["dragover", "dragenter"].forEach((e) => zone.addEventListener(e, (ev) => { ev.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((e) => zone.addEventListener(e, (ev) => { ev.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", (ev) => onFiles([...ev.dataTransfer.files]));
}
function addRcptFiles(files) {
  const have = new Set(rcptFiles.map((f) => f.name));
  files.forEach((f) => { if (!have.has(f.name)) { rcptFiles.push(f); have.add(f.name); } });
  renderRcptFiles(); renderThumbs(); updateAnalyzeBtn();
}
function renderRcptFiles() {
  $("rcptFiles").innerHTML = rcptFiles.map((f, i) =>
    `<div class="file"><span class="n">${i + 1}</span><span class="nm">${esc(f.name)}</span><button class="x" onclick="removeRcpt(${i})">✕</button></div>`).join("");
  $("previewToggleWrap").classList.toggle("hidden", rcptFiles.length === 0);
}
function removeRcpt(i) {
  const f = rcptFiles[i];
  if (urlCache[f.name]) { URL.revokeObjectURL(urlCache[f.name]); delete urlCache[f.name]; }
  rcptFiles.splice(i, 1); renderRcptFiles(); renderThumbs(); updateAnalyzeBtn();
}
function pendingFiles() { return rcptFiles.filter((f) => !analyzedNames.has(f.name)); }
function updateAnalyzeBtn() {
  const p = pendingFiles().length;
  const btn = $("analyzeBtn"); btn.disabled = p === 0; btn.innerHTML = `영수증 분석 (${p}장)`;
  $("resetListBtn").disabled = rcptRows.length === 0;
}

// ===== 미리보기 =====
function renderThumbs() {
  const on = $("previewToggle").checked && rcptFiles.length > 0;
  const box = $("thumbs");
  if (!on) { box.innerHTML = ""; return; }
  box.innerHTML = rcptFiles.map((f) => {
    if (!urlCache[f.name]) urlCache[f.name] = URL.createObjectURL(f);
    return `<div class="thumb" onclick="openLightbox('${esc(f.name)}')"><img src="${urlCache[f.name]}" alt="${esc(f.name)}"/><span class="cap">${esc(f.name)}</span></div>`;
  }).join("");
}
function openLightbox(name) { $("lbImg").src = urlCache[name] || ""; $("lightbox").classList.remove("hidden"); }
function closeLightbox() { $("lightbox").classList.add("hidden"); $("lbImg").src = ""; }

// ===== 분석 =====
async function analyze() {
  const pending = pendingFiles();
  if (!pending.length) return;
  // 연결 방식이 OpenAI/직접인데 필수값이 없으면 사내 모델로 잘못 빠지지 않게 막는다.
  if (AI.mode === "openai" && !AI.key) {
    note($("analyzeNote"), "warn", "OpenAI로 분석하려면 API Key가 필요해요. 사내 모델을 쓰려면 사이드바에서 ‘↩ 사내 기본 모델로 되돌리기’를 누르세요.");
    return;
  }
  if (AI.mode === "custom" && !AI.cuBase) {
    note($("analyzeNote"), "warn", "‘직접 입력’은 서버 주소(base_url)가 필요해요. 사내 모델을 쓰려면 ‘↩ 사내 기본 모델로 되돌리기’를 누르세요.");
    return;
  }
  const btn = $("analyzeBtn"); btn.disabled = true; btn.innerHTML = `<span class="spin"></span> 분석 중… (${pending.length}장)`;
  note($("analyzeNote"), "info", "");
  const fd = new FormData();
  pending.forEach((f) => fd.append("images", f));
  aiOverride(fd);
  try {
    const resp = await fetch("/api/expense/analyze", { method: "POST", body: fd });
    if (!resp.ok) throw new Error(await errText(resp));
    const data = await resp.json();
    rcptRows = rcptRows.filter((r) => r.store || r.date);  // 비어 있는 시드/수동 행 제거 후 결과 추가
    data.rows.forEach((r) => rcptRows.push(r));
    pending.forEach((f) => analyzedNames.add(f.name));
    rcptRows.sort((a, b) => dateKey(a.date).localeCompare(dateKey(b.date)));
    renderReview();
    const errs = data.rows.filter((r) => r.error);
    let msg = `새 영수증 ${data.rows.length}건 분석 완료 (누적 ${rcptRows.length}건).`;
    if (errs.length) {
      msg += ` ${errs.length}건은 인식 실패 — 표에서 직접 고치거나 다시 분석하세요.`;
      msg += `<br><span class="muted">사유: ${esc(errs[0].error)}</span>`;
      msg += `<br><button type="button" class="btn ghost mini" onclick="retryFailed()">↻ 실패 ${errs.length}건 다시 분석</button>`;
    }
    note($("analyzeNote"), errs.length ? "warn" : "info", msg);
  } catch (e) { note($("analyzeNote"), "err", "분석 실패: " + e.message); }
  finally { updateAnalyzeBtn(); }
}

// 인식 실패한 행만 골라 다시 분석한다. (원본 사진은 rcptFiles에 그대로 남아 있음)
async function retryFailed() {
  const failed = rcptRows.filter((r) => r.error);
  if (!failed.length) return;
  const names = new Set(failed.map((r) => r.filename).filter(Boolean));
  rcptRows = rcptRows.filter((r) => !r.error);   // 실패 행 제거
  names.forEach((n) => analyzedNames.delete(n));  // 다시 '미분석' 상태로
  renderReview(); updateAnalyzeBtn();
  await analyze();                                 // 미분석 사진(=방금 실패분)만 재분석
}

// ===== 검토 표 =====
function optTags(list, val) { return ['<option value=""></option>'].concat(list.map((o) => `<option ${o === val ? "selected" : ""}>${esc(o)}</option>`)).join(""); }
function welfareAlloc(rows, budgetWon) {
  const order = rows.map((_, i) => i).sort((a, b) => dateKey(rows[a].date).localeCompare(dateKey(rows[b].date)));
  let remaining = budgetWon || 0; const give = {};
  order.forEach((i) => { const a = toInt(rows[i].amount) || 0; const g = Math.min(a, remaining); remaining -= g; give[i] = g; });
  return rows.map((_, i) => give[i]);
}
function renderReview() {
  if (!rcptRows.length) rcptRows.push(blankRow());  // 분석 전에도 빈 행 1개를 보여줘 직접 입력 가능
  const welfare = isWelfare();
  $("reviewDesc").textContent = welfare
    ? "복지비는 한도까지 영수일자 순으로 청구금액이 자동 배분돼요. 목적/거래처 등은 수정 가능."
    : "분석으로 추출된 내역으로 양식을 작성합니다. '+ 행 추가'로 직접 입력할 수도 있습니다.";
  $("bulkPays").innerHTML = OPTIONS.payment.map((p) => `<button class="${lastBulk.payment === p ? "on" : ""}" onclick="bulkApply('payment','${esc(p)}')">${esc(p)}</button>`).join("");
  const alloc = welfare ? welfareAlloc(rcptRows, (toInt($("welfare").value) || 0) * 10000) : null;
  const head = `<tr><th>#</th><th>영수일자</th><th>거래처명</th><th>목적</th><th>영수금액</th><th>${welfare ? "청구금액(자동)" : "청구금액"}</th><th>결제방식</th><th>영수시간</th><th>지역</th><th>비고</th><th></th></tr>`;
  let capped = 0;
  const rowsHtml = rcptRows.map((r, i) => {
    let claimCell, over = "";
    if (welfare) { claimCell = `<span class="muted">${fmt(alloc[i])}</span>`; }
    else {
      const lim = OPTIONS.limits[r.purpose]; const base = toInt(r.claim) ?? toInt(r.amount);
      if (lim != null && base != null && base > lim) { over = `<button type="button" class="over" title="청구금액을 한도로 맞춤" onclick="applyLimit(${i})">한도 ${fmt(lim)} 적용 ↓</button>`; capped++; }
      claimCell = `<input type="text" value="${esc(r.claim)}" oninput="upd(${i},'claim',this.value)"/>${over}`;
    }
    return `<tr class="drow"><td>${i + 1}</td>
      <td><input type="text" value="${esc(r.date)}" title="${esc(r.date)}" oninput="upd(${i},'date',this.value)"/></td>
      <td><input type="text" value="${esc(r.store)}" title="${esc(r.store)}" oninput="upd(${i},'store',this.value)"/></td>
      <td><select title="${esc(r.purpose)}" onchange="upd(${i},'purpose',this.value)">${optTags(OPTIONS.purpose, r.purpose)}</select></td>
      <td class="num"><input type="text" value="${esc(r.amount)}" oninput="upd(${i},'amount',this.value)"/></td>
      <td class="num">${claimCell}</td>
      <td><select onchange="upd(${i},'payment',this.value)">${optTags(OPTIONS.payment, r.payment)}</select></td>
      <td><input type="text" value="${esc(r.time)}" oninput="upd(${i},'time',this.value)"/></td>
      <td><input type="text" value="${esc(r.region)}" title="${esc(r.region)}" oninput="upd(${i},'region',this.value)"/></td>
      <td><input type="text" value="${esc(r.note)}" title="${esc(r.note)}" oninput="upd(${i},'note',this.value)"/></td>
      <td class="rowact"><button class="x" title="이 행 삭제" onclick="delRow(${i})">✕</button></td></tr>`;
  });
  // 행과 행 사이(및 맨 위)에 '여기에 추가' 파란 + 어포던스 — 마우스를 올리면 나타남
  const insAfford = (i) => `<tr class="insrow"><td colspan="11"><button type="button" class="insbtn" title="여기에 행 추가" onclick="insertRow(${i})">+</button></td></tr>`;
  const body = insAfford(-1) + rowsHtml.map((h, i) => h + insAfford(i)).join("");
  $("rcptTable").innerHTML = head + body;
  if (welfare) {
    const budget = (toInt($("welfare").value) || 0) * 10000;
    const totalR = rcptRows.reduce((s, r) => s + (toInt(r.amount) || 0), 0);
    const totalC = alloc.reduce((s, v) => s + v, 0);
    note($("welfarePreview"), "info", `복지비 한도 ${fmt(budget)}원 · 영수 합계 ${fmt(totalR)}원 · 자동 청구 ${fmt(totalC)}원` + (totalR > budget ? " — 한도 초과분은 청구에서 제외됩니다." : ""));
  } else if (capped) { note($("welfarePreview"), "warn", `한도 초과 ${capped}건은 '생성' 시 목적별 한도로 자동 보정됩니다.`); }
  else { $("welfarePreview").innerHTML = ""; }
}
function upd(i, key, val) { rcptRows[i][key] = val; if (["purpose", "claim", "amount", "date"].includes(key)) renderReview(); }
function blankRow() { return { date: "", store: "", purpose: "", amount: "", claim: "", payment: "", time: "", region: "", note: "", filename: "", error: null }; }
function applyLimit(i) { const lim = OPTIONS.limits[rcptRows[i].purpose]; if (lim == null) return; rcptRows[i].claim = String(lim); renderReview(); }
function insertRow(i) { rcptRows.splice(i + 1, 0, blankRow()); renderReview(); updateAnalyzeBtn(); }
function delRow(i) { rcptRows.splice(i, 1); renderReview(); updateAnalyzeBtn(); }
function addRow() { rcptRows.push(blankRow()); renderReview(); }
function bulkApply(key, val) {
  if (!val) return;
  lastBulk[key] = val;                              // 무엇을 적용했는지 기억 (버튼/드롭다운 표시용)
  rcptRows.forEach((r) => { r[key] = val; });
  renderReview();
  if (key === "purpose") $("bulkPurpose").value = val;  // 고른 목적을 드롭다운에 그대로 표시
}

// ===== 비용청구서 생성 =====
async function generateExpense() {
  const dept = $("dept").value.trim(), name = $("name").value.trim();
  const missing = []; if (!dept) missing.push("소속부서명"); if (!name) missing.push("성명");
  if (missing.length) { note($("genNote"), "warn", `먼저 1단계 '기초정보 입력'에서 ${missing.join(" · ")}을(를) 입력해 주세요.`); return; }
  const rows = rcptRows.filter((r) => r.store || r.date);
  if (!rows.length) { note($("genNote"), "warn", "채울 데이터가 없어요. 먼저 영수증을 분석하세요."); return; }
  const welfare = isWelfare() ? (toInt($("welfare").value) || 0) * 10000 : 0;
  const fd = new FormData();
  fd.append("payload", JSON.stringify({ rows, basic: { dept, name, title }, welfare_budget: welfare }));
  if (expenseTpl) fd.append("template", expenseTpl);
  const btn = $("genBtn"); btn.disabled = true; btn.innerHTML = `<span class="spin"></span> 생성 중…`;
  note($("genNote"), "info", "");
  try {
    const resp = await fetch("/api/expense/generate", { method: "POST", body: fd });
    if (!resp.ok) throw new Error(await errText(resp));
    const r = await saveToClient(resp, "비용청구양식_작성완료.xlsm");
    if (r.aborted) { note($("genNote"), "warn", "저장이 취소됐어요."); }
    else note($("genNote"), "info",
      `<b>✅ 비용청구서가 다운로드됐어요</b><br/>${esc(r.name)} 가 내 PC에 저장됐어요.`);
    $("newWrap").classList.remove("hidden");
  } catch (e) { note($("genNote"), "err", "생성 실패: " + e.message); }
  finally { btn.disabled = false; btn.innerHTML = "비용청구서 생성 & 다운로드"; }
}

// ===== 초기화 =====
function resetList() { rcptRows = []; analyzedNames = new Set(); renderReview(); updateAnalyzeBtn(); note($("analyzeNote"), "info", "목록을 비웠어요. 다시 분석할 수 있어요."); }
function resetAll() {
  Object.values(urlCache).forEach((u) => URL.revokeObjectURL(u));
  urlCache = {}; rcptFiles = []; rcptRows = []; analyzedNames = new Set();
  $("dept").value = ""; $("name").value = ""; $("welfare").value = ""; $("previewToggle").checked = false;
  renderRcptFiles(); renderThumbs(); renderReview(); updateAnalyzeBtn();
  note($("analyzeNote"), "", ""); note($("genNote"), "", ""); $("newWrap").classList.add("hidden");
}

// ===== 클라이언트(접속한 내 PC) 다운로드 =====
// 응답을 접속한 사용자의 PC에 저장. 브라우저가 지원하면(보안 컨텍스트) 저장 위치를
// 직접 고르는 네이티브 창을 띄우고, 아니면 일반 다운로드(브라우저 다운로드 폴더)로 받는다.
async function saveToClient(resp, fallbackName) {
  const cd = resp.headers.get("Content-Disposition") || "";
  const m = /filename\*=UTF-8''([^;]+)/.exec(cd);
  const name = m ? decodeURIComponent(m[1]) : fallbackName;
  const blob = await resp.blob();
  if (window.isSecureContext && window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({ suggestedName: name });
      const w = await handle.createWritable();
      await w.write(blob); await w.close();
      return { name, picked: true };
    } catch (e) { if (e && e.name === "AbortError") return { aborted: true }; }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; document.body.appendChild(a); a.click();
  a.remove(); URL.revokeObjectURL(url);
  return { name, picked: false };
}

// ===== 연장근무 =====
async function parseAttendance(files) {
  attFile = files[0]; if (!attFile) return;
  note($("attNote"), "info", "읽는 중…");
  const fd = new FormData(); fd.append("attendance", attFile);
  try {
    const resp = await fetch("/api/overtime/parse", { method: "POST", body: fd });
    if (!resp.ok) throw new Error(await errText(resp));
    const d = await resp.json(); otRows = d.rows;
    note($("attNote"), "info", `📎 ${esc(attFile.name)}`);
    $("otMetrics").innerHTML =
      `<div class="metric"><div class="k">이름</div><div class="v">${esc(d.name) || "-"}</div></div>` +
      `<div class="metric"><div class="k">기간</div><div class="v">${d.year || "-"}.${d.month || "-"}</div></div>` +
      `<div class="metric"><div class="k">연장근무 대상일</div><div class="v">${otRows.length}일</div></div>`;
    renderAtt();
    $("otReviewCard").classList.remove("hidden"); $("otGenCard").classList.remove("hidden");
    if (!otRows.length) note($("attNote"), "warn", "연장근무 대상일(승인 초과>0)이 없어요.");
  } catch (e) { note($("attNote"), "err", e.message); }
}
function renderAtt() {
  const head = `<tr><th>일자</th><th>출근</th><th>퇴근</th><th>근무시작<br/>(출근+9h)</th><th>근무종료</th><th>승인초과</th><th>대체휴무 지급</th><th>대체휴무 시간</th><th>비고</th></tr>`;
  const body = otRows.map((r, i) => `<tr>
    <td>${r.day}일</td><td>${esc(r.clock_in)}</td><td>${esc(r.clock_out)}</td><td>${esc(r.work_start)}</td><td>${esc(r.work_end)}</td><td>${esc(r.approved_ot)}</td>
    <td><select onchange="updOt(${i},'payoff',this.value)">${optTags(["X", "O"], r.payoff)}</select></td>
    <td><input type="text" value="${esc(r.hours)}" placeholder="예: 1:00" oninput="updOt(${i},'hours',this.value)"/></td>
    <td><input type="text" value="${esc(r.note)}" oninput="updOt(${i},'note',this.value)"/></td></tr>`).join("");
  $("attTable").innerHTML = head + body;
}
function updOt(i, key, val) { otRows[i][key] = val; }
function otBulk(v) { otRows.forEach((r) => { r.payoff = v; }); renderAtt(); }
async function generateOvertime() {
  if (!attFile) { note($("otGenNote"), "warn", "먼저 근태현황 파일을 올려주세요."); return; }
  const extras = {};
  otRows.forEach((r) => { extras[r.day] = { payoff: r.payoff, hours: r.hours, note: r.note }; });
  const parts = [$("otDept").value.trim(), $("otPos").value.trim()].filter(Boolean);
  const fd = new FormData();
  fd.append("attendance", attFile);
  fd.append("payload", JSON.stringify({ extras, dept_position: parts.join(" / ") || null }));
  if (overtimeTpl) fd.append("template", overtimeTpl);
  const btn = $("otGenBtn"); btn.disabled = true; btn.innerHTML = `<span class="spin"></span> 생성 중…`;
  note($("otGenNote"), "info", "");
  try {
    const resp = await fetch("/api/overtime/generate", { method: "POST", body: fd });
    if (!resp.ok) throw new Error(await errText(resp));
    const r = await saveToClient(resp, "연장근무신청서_작성완료.xlsx");
    if (r.aborted) note($("otGenNote"), "warn", "저장이 취소됐어요.");
    else note($("otGenNote"), "info", `<b>✅ 신청서가 다운로드됐어요</b><br/>${esc(r.name)} 가 내 PC에 저장됐어요.`);
  } catch (e) { note($("otGenNote"), "err", "생성 실패: " + e.message); }
  finally { btn.disabled = false; btn.innerHTML = "신청서 생성 & 다운로드"; }
}

// ===== 옵션(드롭다운) =====
async function fetchExpenseOptions() {
  const fd = new FormData();
  if (expenseTpl) fd.append("template", expenseTpl);
  try {
    const r = await fetch("/api/expense/options", { method: "POST", body: fd });
    if (r.ok) {
      OPTIONS = await r.json();
      $("bulkPurpose").innerHTML = '<option value="">목적 선택</option>' + OPTIONS.purpose.map((p) => `<option>${esc(p)}</option>`).join("");
      renderReview();
    }
  } catch {}
}

// ===== 초기화 =====
async function init() {
  setupDrop($("rcptDrop"), $("rcptInput"), addRcptFiles);
  setupDrop($("attDrop"), $("attInput"), parseAttendance);
  $("tplInput").onchange = (e) => onTplChange(e.target.files[0]);
  ["oaModel", "cuBase", "cuModel", "aiKey"].forEach((id) => $(id).addEventListener("change", saveAi));
  loadAi();
  try { const r = await fetch("/api/config"); if (r.ok) CONFIG = { ...CONFIG, ...(await r.json()) }; } catch {}
  updateTplName();
  await fetchExpenseOptions();
}
init();
