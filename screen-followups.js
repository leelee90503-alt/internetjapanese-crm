import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

function todayStr() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d.toISOString().slice(0, 10);
}

function addDays(dateStr, days) {
  const d = new Date((dateStr || todayStr()) + "T00:00:00");
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

// Overdue / Today / This Week (next 7 days) / Later -- a simple fixed
// grouping rather than a real calendar-week boundary, since staff mainly
// care about "how soon do I need to act on this."
function bucketFor(dueDate) {
  const due = new Date(dueDate + "T00:00:00");
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.round((due - today) / 86400000);
  if (diffDays < 0) return "overdue";
  if (diffDays === 0) return "today";
  if (diffDays <= 7) return "week";
  return "later";
}

const BUCKET_ORDER = ["overdue", "today", "week", "later"];
const BUCKET_LABEL = { overdue: "Overdue", today: "Today", week: "This Week", later: "Later" };
const BUCKET_BADGE = { overdue: "badge error", today: "badge warn", week: "badge neutral", later: "badge neutral" };

export async function renderFollowups(container, ctx) {
  let followUps = [];
  let showCompleted = false;
  let postponeId = null;

  async function load() {
    const { data, error } = await supabase
      .from("follow_ups")
      .select(
        `id, due_date, reason, status, completed_at, created_at,
         customers(id, name, phone),
         orders(id, order_date)`
      )
      .order("due_date", { ascending: true });
    if (error) {
      container.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
      followUps = [];
      return;
    }
    followUps = data || [];
  }

  function rowHtml(f) {
    const bucket = bucketFor(f.due_date);
    return `
      <tr data-fid="${f.id}">
        <td>${
          f.status === "done"
            ? `<span class="badge ok">Done</span>`
            : `<span class="${BUCKET_BADGE[bucket]}">${escapeHtml(f.due_date)}</span>`
        }</td>
        <td>${escapeHtml(f.customers?.name || "-")}</td>
        <td>${escapeHtml(f.customers?.phone || "-")}</td>
        <td>${f.orders?.order_date ? escapeHtml(f.orders.order_date) : "-"}</td>
        <td>${escapeHtml(f.reason || "-")}</td>
        <td>
          ${
            f.status === "pending"
              ? `<button class="btn small" data-complete="${f.id}">Complete</button>
                 <button class="btn small" data-postpone="${f.id}">Postpone</button>`
              : `<span class="muted">Completed${
                  f.completed_at ? " " + escapeHtml(new Date(f.completed_at).toLocaleDateString()) : ""
                }</span>`
          }
        </td>
      </tr>
      ${
        postponeId === f.id
          ? `<tr class="postpone-row"><td colspan="6">
              <div class="inline-form">
                <span class="muted">Postpone to:</span>
                <button class="btn small" data-postpone-quick="${f.id}:7">+1 Week</button>
                <button class="btn small" data-postpone-quick="${f.id}:14">+2 Weeks</button>
                <input type="date" class="postpone-date-input" data-postpone-date="${f.id}" value="${f.due_date}" />
                <button class="btn small primary" data-postpone-apply="${f.id}">Apply</button>
                <button class="btn small" data-postpone-cancel="${f.id}">Cancel</button>
              </div>
            </td></tr>`
          : ""
      }
    `;
  }

  function tableHtml(title, rows) {
    return `
      <h3>${title} (${rows.length})</h3>
      <table class="data-table">
        <thead><tr><th>Due Date</th><th>Customer</th><th>Phone</th><th>Order Date</th><th>Reason</th><th></th></tr></thead>
        <tbody>${rows.map(rowHtml).join("") || `<tr><td colspan="6" class="muted">None.</td></tr>`}</tbody>
      </table>`;
  }

  function draw() {
    const pending = followUps.filter((f) => f.status === "pending");
    const completed = followUps.filter((f) => f.status === "done");

    const groups = {};
    for (const key of BUCKET_ORDER) groups[key] = [];
    for (const f of pending) groups[bucketFor(f.due_date)].push(f);

    container.innerHTML = `
      <div class="screen">
        <h2>Follow-ups</h2>
        <p class="muted">Customers who need a follow-up after their order (e.g. confirming a credit was applied, or a callback) -- grouped by how soon it's due.</p>
        <label class="checkbox-inline"><input type="checkbox" id="show-completed" ${showCompleted ? "checked" : ""} /> Show completed</label>
        ${
          pending.length === 0
            ? `<div class="card"><p class="muted">No pending follow-ups.</p></div>`
            : BUCKET_ORDER.filter((k) => groups[k].length > 0)
                .map((k) => tableHtml(BUCKET_LABEL[k], groups[k]))
                .join("")
        }
        ${showCompleted ? tableHtml("Completed", completed.slice().sort((a, b) => (b.completed_at || "").localeCompare(a.completed_at || ""))) : ""}
      </div>
    `;

    container.querySelector("#show-completed").addEventListener("change", (e) => {
      showCompleted = e.target.checked;
      draw();
    });

    container.querySelectorAll("[data-complete]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.complete;
        const { error } = await supabase
          .from("follow_ups")
          .update({ status: "done", completed_at: new Date().toISOString(), completed_by: ctx.profile?.id || null })
          .eq("id", id);
        if (error) {
          alert("Failed to complete: " + error.message);
          return;
        }
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-postpone]").forEach((btn) => {
      btn.addEventListener("click", () => {
        postponeId = postponeId === btn.dataset.postpone ? null : btn.dataset.postpone;
        draw();
      });
    });
    container.querySelectorAll("[data-postpone-cancel]").forEach((btn) => {
      btn.addEventListener("click", () => {
        postponeId = null;
        draw();
      });
    });
    container.querySelectorAll("[data-postpone-quick]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const [id, days] = btn.dataset.postponeQuick.split(":");
        const input = container.querySelector(`[data-postpone-date="${id}"]`);
        const base = followUps.find((f) => f.id === id)?.due_date || todayStr();
        input.value = addDays(base, Number(days));
      });
    });
    container.querySelectorAll("[data-postpone-apply]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.postponeApply;
        const input = container.querySelector(`[data-postpone-date="${id}"]`);
        const newDate = input.value;
        if (!newDate) return;
        const { error } = await supabase.from("follow_ups").update({ due_date: newDate }).eq("id", id);
        if (error) {
          alert("Failed to postpone: " + error.message);
          return;
        }
        postponeId = null;
        await load();
        draw();
      });
    });
  }

  await load();
  draw();
}
