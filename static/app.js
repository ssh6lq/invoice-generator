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
let otPeriod = { year: null, month: null };  // 근태현황 파일에서 읽은 연·월 (일자 표기에 사용)
let activeTab = "expense";
let expenseTpl = null;        // 커스텀 양식 File (없으면 기본)
let overtimeTpl = null;
// AI 연결: 기본은 '사내'(서버 .env). 사용자가 고급 설정에서 OpenAI/직접 로컬로 덮어쓸 수 있음.
// label = UI 표기, value = 실제 API에 보내는 모델명
const OPENAI_MODELS = [
  { label: "gpt-5.3", value: "gpt-5.3-chat-latest" },
  { label: "gpt-5.4-mini", value: "gpt-5.4-mini" },
  { label: "gpt-4o", value: "gpt-4o" },
  { label: "gpt-4o-mini", value: "gpt-4o-mini" },
];
const oaLabel = (v) => (OPENAI_MODELS.find((m) => m.value === v) || {}).label || v;
let AI = { mode: "default", oaModel: "gpt-4o", cuBase: "", cuModel: "", key: "" };
let lastBulk = { purpose: "", payment: "" };  // 일괄 채우기로 마지막에 적용한 값 (표시용)

// ===== 공통 =====
const $ = (id) => document.getElementById(id);
const esc = (s) => (s == null ? "" : String(s).replace(/"/g, "&quot;"));
const toInt = (v) => { if (v == null || v === "") return null; const n = parseInt(String(v).replace(/,/g, "").trim(), 10); return isNaN(n) ? null : n; };
const fmt = (n) => (n == null || n === "" ? "" : Number(n).toLocaleString());
const isWelfare = () => title === "복지비";
function dateKey(s) { const m = /(\d{4})\D(\d{1,2})\D(\d{1,2})/.exec(esc(s)); return m ? `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` : "9999-99-99"; }
// 숫자만 입력해도 자동으로 구분기호를 넣어준다. 날짜 → YYYY.MM.DD, 시간 → HH:MM
function fmtDateInput(v) {
  const d = String(v ?? "").replace(/\D/g, "").slice(0, 8);
  let out = d.slice(0, 4);
  if (d.length > 4) out += "." + d.slice(4, 6);
  if (d.length > 6) out += "." + d.slice(6, 8);
  return out;
}
function fmtTimeInput(v) {
  const d = String(v ?? "").replace(/\D/g, "").slice(0, 4);
  return d.length > 2 ? d.slice(0, 2) + ":" + d.slice(2, 4) : d;
}
function note(el, type, msg) { el.innerHTML = msg ? `<div class="note ${type}">${msg}</div>` : ""; }

// 매크로 차단 해제 안내: 사용자가 신뢰 영역에 등록할 '이 서버 주소'(접속한 IP/포트 그대로)
function serverAddr() { return location.origin; }
function copyServerAddr(btn) {
  const v = serverAddr();
  const flash = (msg) => { if (!btn) return; const t = btn.dataset.label || btn.textContent; btn.dataset.label = t; btn.textContent = msg; setTimeout(() => { btn.textContent = t; }, 1200); };
  // 사내망은 보통 HTTP+IP라 navigator.clipboard가 없음 → 실패 시 execCommand 폴백을 쓴다.
  const fallback = () => {
    try {
      const ta = document.createElement("textarea");
      ta.value = v; ta.style.position = "fixed"; ta.style.top = "-1000px"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.focus(); ta.select();
      const ok = document.execCommand("copy");
      ta.remove();
      flash(ok ? "복사됨 ✓" : "복사 실패");
    } catch (e) { flash("복사 실패"); }
  };
  if (navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(v).then(() => flash("복사됨 ✓")).catch(fallback);
  } else {
    fallback();
  }
}

function switchTab(which) {
  activeTab = which;
  $("tab-expense").classList.toggle("active", which === "expense");
  $("tab-overtime").classList.toggle("active", which === "overtime");
  $("panel-expense").classList.toggle("hidden", which !== "expense");
  $("panel-overtime").classList.toggle("hidden", which !== "overtime");
  $("aiCard").classList.toggle("hidden", which !== "expense");
  $("fileGuide").classList.toggle("hidden", which !== "overtime");
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
    st.textContent = `✓ OpenAI · ${oaLabel(AI.oaModel)} 사용 중`;
    sub.textContent = "해당 모델로 영수증을 분석합니다. \nAPI Key를 입력하세요.";
    reset.classList.remove("hidden");
  } else if (AI.mode === "custom") {
    st.textContent = `✓ 직접 연결 · ${AI.cuModel || "모델명 미입력"} 사용 중`;
    sub.textContent = "필요한 정보를 입력하세요.";
    reset.classList.remove("hidden");
  } else {
    st.textContent = "✓ 사내 AI 모델 연결됨";
    sub.textContent = "해당 모델로 영수증을 분석합니다.";
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
  $("oaModel").innerHTML = OPENAI_MODELS.map((m) => `<option value="${esc(m.value)}">${esc(m.label)}</option>`).join("");
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
  $("welfareWrap").classList.toggle("hidden", !isWelfare());
  recomputeClaims();   // 모드 전환 시 청구금액 자동값을 새 모드 기준으로 다시 채움
  renderReview();
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
    rcptRows.forEach((r) => { r.date = fmtDateInput(r.date); r.time = fmtTimeInput(r.time); });  // 분석값도 YYYY.MM.DD·HH:MM로 표기 통일
    rcptRows.sort((a, b) => dateKey(a.date).localeCompare(dateKey(b.date)));
    recomputeClaims();   // 분석 직후에도 현재 모드 기준으로 청구금액 자동값을 채움
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
// 비용 모드: 한 행의 청구금액을 그 목적의 한도로 재계산한다.
// 영수금액 기준으로 매번 다시 계산 → 한도보다 크면 한도로, 한도가 더 크면 영수금액 전액으로 복원.
function autoClaimExpense(i) {
  const r = rcptRows[i];
  const amt = toInt(r.amount);
  if (amt == null) return;                       // 영수금액이 없으면 손대지 않음
  const lim = OPTIONS.limits[r.purpose];
  r.claim = String(lim != null && amt > lim ? lim : amt);
}
// 복지비 모드: 전체 행에 영수일자 순 자동 배분 결과를 청구금액으로 채운다.
function recomputeWelfare() {
  const alloc = welfareAlloc(rcptRows, (toInt($("welfare").value) || 0) * 10000);
  rcptRows.forEach((r, i) => { r.claim = String(alloc[i]); });
}
// 현재 모드에 맞게 모든 행의 청구금액 자동값을 다시 채운다.
function recomputeClaims() {
  if (isWelfare()) recomputeWelfare();
  else rcptRows.forEach((_, i) => autoClaimExpense(i));
}
// 복지비 한도(예산) 입력이 바뀔 때: 배분을 다시 계산하고 표를 갱신.
function welfareBudgetChanged() { recomputeWelfare(); renderReview(); }
function renderReview() {
  // 입력 중 표를 다시 그리면 포커스가 사라지므로, 현재 포커스 셀과 커서 위치를 기억했다가 복원한다.
  const act = document.activeElement;
  const focusCell = (act && act.dataset && act.dataset.cell) ? act.dataset.cell : null;
  const caret = focusCell ? act.selectionStart : null;
  if (!rcptRows.length) rcptRows.push(blankRow());  // 분석 전에도 빈 행 1개를 보여줘 직접 입력 가능
  const welfare = isWelfare();
  $("reviewDesc").textContent = welfare
    ? "복지비는 한도까지 영수일자 순으로 청구금액이 자동 배분돼요. 목적/거래처 등은 수정 가능."
    : "추출된 내역으로 양식을 작성합니다. 목적 선택에 따라 청구금액이 목적별 한도로 자동 조정돼요.";
  $("bulkPays").innerHTML = OPTIONS.payment.map((p) => `<button class="${lastBulk.payment === p ? "on" : ""}" onclick="bulkApply('payment','${esc(p)}')">${esc(p)}</button>`).join("");
  const head = `<tr><th>#</th><th>영수일자</th><th>거래처명</th><th>목적</th><th>영수금액</th><th>청구금액(자동)</th><th>결제방식</th><th>영수시간</th><th>지역</th><th>비고</th><th></th></tr>`;
  let capped = 0;  // 비용 모드에서 목적 한도로 깎인 건수(안내용)
  const rowsHtml = rcptRows.map((r, i) => {
    // 청구금액은 모드와 무관하게 동일한 입력박스로 통일 — 자동값(r.claim)이 채워지지만 직접 수정도 가능.
    const lim = OPTIONS.limits[r.purpose], amt = toInt(r.amount);
    if (!welfare && lim != null && amt != null && amt > lim) capped++;
    const claimCell = `<input type="text" data-cell="${i}:claim" value="${esc(r.claim)}" oninput="upd(${i},'claim',this.value)"/>`;
    return `<tr class="drow"><td>${i + 1}</td>
      <td><input type="text" data-cell="${i}:date" value="${esc(r.date)}" title="${esc(r.date)}" placeholder="YYYY.MM.DD" oninput="updDate(${i},this)"/></td>
      <td><input type="text" data-cell="${i}:store" value="${esc(r.store)}" title="${esc(r.store)}" oninput="upd(${i},'store',this.value)"/></td>
      <td><select title="${esc(r.purpose)}" onchange="upd(${i},'purpose',this.value)">${optTags(OPTIONS.purpose, r.purpose)}</select></td>
      <td class="num"><input type="text" data-cell="${i}:amount" value="${esc(r.amount)}" oninput="upd(${i},'amount',this.value)"/></td>
      <td class="num">${claimCell}</td>
      <td><select onchange="upd(${i},'payment',this.value)">${optTags(OPTIONS.payment, r.payment)}</select></td>
      <td><input type="text" data-cell="${i}:time" value="${esc(r.time)}" placeholder="HH:MM" oninput="updTime(${i},this)"/></td>
      <td><input type="text" data-cell="${i}:region" value="${esc(r.region)}" title="${esc(r.region)}" oninput="upd(${i},'region',this.value)"/></td>
      <td><input type="text" data-cell="${i}:note" value="${esc(r.note)}" title="${esc(r.note)}" oninput="upd(${i},'note',this.value)"/></td>
      <td class="rowact"><button class="x" title="이 행 삭제" onclick="delRow(${i})">✕</button></td></tr>`;
  });
  // 행 사이(및 맨 위)에 '여기에 추가' 어포던스 — 평소엔 빈 줄, hover 시 연한 파란 선 + 왼쪽 끝 '+' 노출. 줄 전체 클릭 가능.
  const insAfford = (i) => `<tr class="insrow" onclick="insertRow(${i})" title="여기에 행 추가"><td colspan="11"><span class="insbtn"><svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg></span></td></tr>`;
  const body = insAfford(-1) + rowsHtml.map((h, i) => h + insAfford(i)).join("");
  $("rcptTable").innerHTML = head + body;
  // 재렌더 전에 포커스가 있던 셀로 다시 포커스를 옮기고 커서 위치도 복원한다.
  if (focusCell) {
    const el = $("rcptTable").querySelector(`[data-cell="${focusCell}"]`);
    if (el) { el.focus(); try { el.setSelectionRange(caret, caret); } catch (e) {} }
  }
  if (welfare) {
    const budget = (toInt($("welfare").value) || 0) * 10000;
    const totalR = rcptRows.reduce((s, r) => s + (toInt(r.amount) || 0), 0);
    const totalC = rcptRows.reduce((s, r) => s + (toInt(r.claim) || 0), 0);
    note($("welfarePreview"), "info", `복지비 한도 ${fmt(budget)}원 · 영수 합계 ${fmt(totalR)}원 · 자동 청구 ${fmt(totalC)}원` + (totalR > budget ? " — 한도 초과분은 청구에서 제외됩니다." : ""));
  } else if (capped) { note($("welfarePreview"), "info", `${capped}건이 목적별 한도로 자동 적용됐어요. 필요하면 청구금액을 직접 수정할 수 있어요.`); }
  else { $("welfarePreview").innerHTML = ""; }
}
function upd(i, key, val) {
  rcptRows[i][key] = val;
  if (isWelfare()) {
    // 복지비: 금액·날짜가 바뀌면 영수일자 순 자동 배분을 다시 계산한다.
    if (key === "amount" || key === "date") recomputeWelfare();
  } else {
    // 비용: 목적·영수금액이 바뀌면 그 목적의 한도로 청구금액을 다시 계산한다.
    if (key === "purpose" || key === "amount") autoClaimExpense(i);
  }
  if (["purpose", "claim", "amount", "date"].includes(key)) renderReview();
}
// 영수일자: 숫자만 쳐도 YYYY.MM.DD로 자동 정리. 복지비는 일자순 배분이라 재계산이 필요.
function updDate(i, el) {
  const f = fmtDateInput(el.value);
  el.value = f; rcptRows[i].date = f;
  if (isWelfare()) { recomputeWelfare(); renderReview(); }
}
// 영수시간: 숫자만 쳐도 HH:MM으로 자동 정리. 표 계산엔 영향 없어 재렌더하지 않음(커서 유지).
function updTime(i, el) {
  const f = fmtTimeInput(el.value);
  el.value = f; rcptRows[i].time = f;
}
function blankRow() { return { date: "", store: "", purpose: "", amount: "", claim: "", payment: "", time: "", region: "", note: "", filename: "", error: null }; }
function insertRow(i) { rcptRows.splice(i + 1, 0, blankRow()); renderReview(); updateAnalyzeBtn(); }
function delRow(i) { rcptRows.splice(i, 1); renderReview(); updateAnalyzeBtn(); }
function addRow() { rcptRows.push(blankRow()); renderReview(); }
function bulkApply(key, val) {
  if (!val) return;
  lastBulk[key] = val;                              // 무엇을 적용했는지 기억 (버튼/드롭다운 표시용)
  rcptRows.forEach((r) => { r[key] = val; });
  if (key === "purpose" && !isWelfare()) rcptRows.forEach((_, i) => autoClaimExpense(i));  // 목적 일괄 적용 시 청구금액도 한도로 재계산
  renderReview();
  // 드롭다운을 '목적 선택'으로 되돌려, 같은 목적을 다시 골라도 onchange가 발생해 재적용되게 한다.
  if (key === "purpose") $("bulkPurpose").value = "";
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
// 일괄 채우기 상태(마지막 적용 표시 + 목적 드롭다운)를 비운다.
function resetBulk() { lastBulk = { purpose: "", payment: "" }; const bp = $("bulkPurpose"); if (bp) bp.value = ""; }
function resetList() { rcptRows = []; analyzedNames = new Set(); resetBulk(); renderReview(); updateAnalyzeBtn(); note($("analyzeNote"), "info", "목록을 비웠어요. 다시 분석할 수 있어요."); }
function resetAll() {
  Object.values(urlCache).forEach((u) => URL.revokeObjectURL(u));
  urlCache = {}; rcptFiles = []; rcptRows = []; analyzedNames = new Set();
  $("dept").value = ""; $("name").value = ""; $("welfare").value = ""; $("previewToggle").checked = false;
  resetBulk();
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
    const d = await resp.json(); otRows = d.rows; otPeriod = { year: d.year, month: d.month };
    note($("attNote"), "info", `📎 ${esc(attFile.name)}`);
    $("otMetrics").innerHTML =
      `<div class="metric"><div class="k">이름</div><div class="v">${esc(d.name) || "-"}</div></div>` +
      `<div class="metric"><div class="k">기간</div><div class="v">${d.year || "-"}.${d.month || "-"}</div></div>` +
      `<div class="metric"><div class="k">연장근무 대상일</div><div class="v">${otRows.length}일</div></div>` +
      `<div class="metric"><div class="k">총 신청시간</div><div class="v"><span id="otTotalReq">${fmtHalf(otTotalReq())}</span>시간</div></div>`;
    renderAtt();
    $("otReviewCard").classList.remove("hidden"); $("otGenCard").classList.remove("hidden");
    if (!otRows.length) note($("attNote"), "warn", "연장근무 대상일(승인 초과>0)이 없어요.");
  } catch (e) { note($("attNote"), "err", e.message); }
}
// 일자 표기: 연·월을 알면 YYYY.MM.DD, 모르면 'N일'로 폴백.
function otDateLabel(day) {
  const p2 = (n) => String(n).padStart(2, "0");
  return otPeriod.year && otPeriod.month ? `${otPeriod.year}.${p2(otPeriod.month)}.${p2(day)}` : `${day}일`;
}
// HH:MM → 소수 시간. "1:30" → 1.5
function hmToDec(s) { const m = /(\d+):(\d{1,2})/.exec(String(s || "")); return m ? +m[1] + (+m[2]) / 60 : null; }
// 대체휴무 시간 입력 → 소수 시간. "1" / "1.5" / "1:30" 모두 허용.
function offHoursDec(s) { s = String(s || "").trim(); if (!s) return 0; if (s.includes(":")) return hmToDec(s) || 0; const n = parseFloat(s); return isNaN(n) ? 0 : n; }
const fmtHalf = (v) => (Number.isInteger(v) ? String(v) : v.toFixed(1));
// 승인초과 → 신청 시간(30분 단위 내림, 0.5 step). 대체휴무 지급(O)이면 대체휴무 시간만큼 차감.
// 양식의 ROUNDDOWN(MAX(0,(근무시간)-Q)*2)/2 수식과 동일.
function otReqHoursNum(r) {
  const base = hmToDec(r.approved_ot);
  if (base == null) return null;
  const sub = String(r.payoff).toUpperCase() === "O" ? offHoursDec(r.hours) : 0;
  return Math.floor(Math.max(0, base - sub) * 2) / 2;
}
function otReqHours(r) { const v = otReqHoursNum(r); return v == null ? "" : fmtHalf(v); }
function otTotalReq() { return otRows.reduce((s, r) => s + (otReqHoursNum(r) || 0), 0); }
function refreshReq(i) { const el = $(`reqh-${i}`); if (el) el.textContent = otReqHours(otRows[i]); }
function refreshOtTotal() { const el = $("otTotalReq"); if (el) el.textContent = fmtHalf(otTotalReq()); }
function renderAtt() {
  const head = `<tr><th>일자</th><th>출근</th><th>퇴근</th><th>근무 시작</th><th>근무 종료</th><th>승인초과</th><th>대체휴무</th><th>대체휴무 시간</th><th>비고</th><th>신청 시간</th></tr>`;
  const body = otRows.map((r, i) => `<tr>
    <td>${otDateLabel(r.day)}</td><td>${esc(r.clock_in)}</td><td>${esc(r.clock_out)}</td><td>${esc(r.work_start)}</td><td>${esc(r.work_end)}</td><td>${esc(r.approved_ot)}</td>
    <td><select onchange="updOt(${i},'payoff',this.value)">${optTags(["X", "O"], r.payoff)}</select></td>
    <td><input type="text" value="${esc(r.hours)}" placeholder="예: 1:00" oninput="updOt(${i},'hours',this.value)"/></td>
    <td><input type="text" value="${esc(r.note)}" oninput="updOt(${i},'note',this.value)"/></td>
    <td class="reqh" id="reqh-${i}">${otReqHours(r)}</td></tr>`).join("");
  $("attTable").innerHTML = head + body;
  refreshOtTotal();
}
// 연장근무 전체 초기화: 파일·표·지표·입력값을 비우고 첫 화면으로 되돌린다.
function resetOvertime() {
  attFile = null; otRows = []; otPeriod = { year: null, month: null };
  $("attTable").innerHTML = ""; $("otMetrics").innerHTML = "";
  $("otDept").value = ""; $("otPos").value = "";
  note($("attNote"), "", ""); note($("otGenNote"), "", "");
  $("otReviewCard").classList.add("hidden"); $("otGenCard").classList.add("hidden");
}
function updOt(i, key, val) {
  otRows[i][key] = val;
  // 대체휴무 지급/시간이 바뀌면 신청 시간과 총합을 즉시 다시 계산(표 전체 재렌더 없이).
  if (key === "payoff" || key === "hours") { refreshReq(i); refreshOtTotal(); }
}
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
  // 사용 방법: 처음 방문할 때만 펼쳐 보여주고, 이후엔 접힌 상태 유지
  if (!localStorage.getItem("howtoSeen")) {
    $("howtoExpense").open = true;
    $("howtoOvertime").open = true;
    localStorage.setItem("howtoSeen", "1");
  }
  ["oaModel", "cuBase", "cuModel", "aiKey"].forEach((id) => $(id).addEventListener("change", saveAi));
  loadAi();
  const sa = $("serverAddr"); if (sa) sa.textContent = serverAddr();  // 매크로 안내에 실제 서버 주소 표시
  try { const r = await fetch("/api/config"); if (r.ok) CONFIG = { ...CONFIG, ...(await r.json()) }; } catch {}
  updateTplName();
  await fetchExpenseOptions();
}
init();
