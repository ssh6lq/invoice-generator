const $ = (id) => document.getElementById(id);

// 사용자가 자유 입력한 제목/내용을 그대로 렌더링하므로, 태그가 실행되지 않도록 전체 이스케이프한다.
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function fmtDate(s) {
  return String(s || "").replace("T", " ");
}

async function toggleStatus(id, next) {
  const fd = new FormData();
  fd.append("status", next);
  await fetch(`/api/feedback/${id}/status`, { method: "POST", body: fd });
  load();
}

function openShot(id) {
  $("lbImg").src = `/api/feedback/${id}/screenshot`;
  $("lightbox").classList.remove("hidden");
}

function toggleMenu(e, id) {
  e.stopPropagation();
  document.querySelectorAll(".kebab-menu").forEach((m) => { if (m.id !== `menu-${id}`) m.classList.add("hidden"); });
  $(`menu-${id}`).classList.toggle("hidden");
}
document.addEventListener("click", () => document.querySelectorAll(".kebab-menu").forEach((m) => m.classList.add("hidden")));

async function deleteFeedback(id) {
  if (!confirm("이 문의를 삭제할까요? 되돌릴 수 없어요.")) return;
  await fetch(`/api/feedback/${id}`, { method: "DELETE" });
  load();
}

async function load() {
  const resp = await fetch("/api/feedback");
  const { items } = await resp.json();
  const el = $("fbList");
  if (!items.length) { el.innerHTML = '<div class="fb-empty">아직 접수된 문의가 없어요.</div>'; return; }
  el.innerHTML = items.map((it) => {
    const isDone = it.status === "done";
    const shot = it.has_screenshot
      ? `<img class="fb-shot" src="/api/feedback/${it.id}/screenshot" alt="스크린샷" onclick="openShot(${it.id})"/>`
      : "";
    return `
      <div class="fb-item">
        <div class="fb-top">
          <div class="fb-title">${esc(it.title)}</div>
          <span class="fb-badge ${isDone ? "done" : "open"}">${isDone ? "완료" : "확인중"}</span>
          <button class="kebab-btn" onclick="toggleMenu(event, ${it.id})">⋮</button>
          <div class="kebab-menu hidden" id="menu-${it.id}">
            <button onclick="deleteFeedback(${it.id})">삭제</button>
          </div>
        </div>
        <div class="fb-meta">#${it.id} · ${esc(fmtDate(it.created_at))}</div>
        <p class="fb-content">${esc(it.content)}</p>
        ${shot}
        <button class="fb-toggle" onclick="toggleStatus(${it.id}, '${isDone ? "open" : "done"}')">
          ${isDone ? "확인중으로 되돌리기" : "완료로 표시"}
        </button>
      </div>`;
  }).join("");
}

load();
