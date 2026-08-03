const $ = (id) => document.getElementById(id);

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function fmtDate(s) { return String(s || "").replace("T", " "); }
function fmtWon(n) { return (n || 0).toLocaleString("ko-KR") + "원"; }

const STATUS_LABEL = {
  received: "접수", reviewing: "검토중", approved: "승인",
  rejected: "반려", paid: "지급완료",
};

async function saveSubmission(id) {
  const status = $(`st-${id}`).value;
  const note = $(`note-${id}`).value;
  const fd = new FormData();
  fd.append("status", status);
  fd.append("note", note);
  await fetch(`/api/submissions/${id}/status`, { method: "POST", body: fd });
  load();
}

function toggleMenu(e, id) {
  e.stopPropagation();
  document.querySelectorAll(".kebab-menu").forEach((m) => { if (m.id !== `menu-${id}`) m.classList.add("hidden"); });
  $(`menu-${id}`).classList.toggle("hidden");
}
document.addEventListener("click", () => document.querySelectorAll(".kebab-menu").forEach((m) => m.classList.add("hidden")));

async function deleteSubmission(id) {
  if (!confirm("이 제출 내역을 삭제할까요? 되돌릴 수 없어요.")) return;
  await fetch(`/api/submissions/${id}`, { method: "DELETE" });
  load();
}

const rowsCache = {};

function fmtNum(v) { return (v === null || v === undefined || v === "") ? "" : Number(v).toLocaleString(); }

function renderRowsTable(rows) {
  if (!rows || !rows.length) return '<div class="fb-empty">저장된 검토표가 없어요.</div>';
  const body = rows.map((r, i) => `<tr>
    <td>${i + 1}</td><td>${esc(r.date)}</td><td>${esc(r.store)}</td><td>${esc(r.purpose)}</td>
    <td>${esc(r.payment)}</td>
    <td class="num">${fmtNum(r.amount)}</td><td class="num">${fmtNum(r.claim)}</td>
    <td>${esc(r.time)}</td><td>${esc(r.region)}</td>
    <td>${esc(r.participants)}</td><td>${esc(r.note)}</td></tr>`).join("");
  return `<div class="tbl-wrap"><table class="grid preview">
    <tr><th>#</th><th>영수일자</th><th>거래처명</th><th>목적</th><th>결제방식</th><th>영수금액</th><th>청구금액</th><th>시간</th><th>지역</th><th>참여자</th><th>비고</th></tr>
    ${body}</table></div>`;
}

async function toggleRows(id) {
  const box = $(`rows-${id}`);
  if (!box) return;
  const opening = box.classList.contains("hidden");
  if (!opening) { box.classList.add("hidden"); return; }
  box.classList.remove("hidden");
  if (!rowsCache[id]) {
    box.innerHTML = '<div class="fb-empty">불러오는 중…</div>';
    try {
      const resp = await fetch(`/api/submissions/${id}/rows`);
      if (!resp.ok) throw new Error();
      const { rows } = await resp.json();
      rowsCache[id] = rows;
    } catch (e) {
      box.innerHTML = '<div class="fb-empty">검토표를 불러오지 못했어요.</div>';
      return;
    }
  }
  box.innerHTML = renderRowsTable(rowsCache[id]);
}

async function load() {
  const resp = await fetch("/api/submissions");
  const { items } = await resp.json();
  const el = $("subList");
  if (!items.length) { el.innerHTML = '<div class="fb-empty">아직 제출된 청구서가 없어요.</div>'; return; }
  el.innerHTML = items.map((it) => `
    <div class="fb-item">
      <div class="fb-top">
        <div class="fb-title">${esc(it.dept)} · ${esc(it.name)}</div>
        <span class="fb-badge st-${esc(it.status)}">${esc(STATUS_LABEL[it.status] || it.status)}</span>
        <button class="kebab-btn" onclick="toggleMenu(event, ${it.id})">⋮</button>
        <div class="kebab-menu hidden" id="menu-${it.id}">
          <button onclick="deleteSubmission(${it.id})">삭제</button>
        </div>
      </div>
      <div class="fb-meta">#${it.id} · ${esc(it.title)} · ${it.count}건 · 총 ${fmtWon(it.total_claim)} · ${esc(fmtDate(it.created_at))}</div>
      <div class="sub-row">
        <select id="st-${it.id}">
          ${Object.entries(STATUS_LABEL).map(([v, l]) =>
            `<option value="${v}" ${v === it.status ? "selected" : ""}>${l}</option>`).join("")}
        </select>
        <input type="text" id="note-${it.id}" value="${esc(it.note)}" placeholder="처리 메모(선택, 반려 사유 등)"/>
        <button onclick="saveSubmission(${it.id})">저장</button>
        ${it.has_rows ? `<button class="fb-toggle" onclick="toggleRows(${it.id})">검토표 보기</button>` : ""}
      </div>
      ${it.has_rows ? `<div id="rows-${it.id}" class="rows-box hidden"></div>` : ""}
    </div>`).join("");
}

load();
