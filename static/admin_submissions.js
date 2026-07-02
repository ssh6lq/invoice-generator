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
      </div>
    </div>`).join("");
}

load();
