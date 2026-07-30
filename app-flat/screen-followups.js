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

  // Customer detail modal state -- clicking a customer name here opens the
  // same kind of full detail view as the Customer/Order screen: contact
  // info, their orders, and every follow-up (pending + completed) for them,
  // so staff can see exactly what still needs to be done without hunting
  // across screens.
  let detailCustomerId = null;
  let detailCustomer = null;
  let detailOrders = [];
  let detailFollowUps = [];
  let followUpBusy = false;
  let newFollowUpDate = "";
  let newFollowUpReason = "";

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

  async function loadDetail(customerId) {
    if (!customerId) {
      detailCustomer = null;
      detailOrders = [];
      detailFollowUps = [];
      return;
    }
    const [{ data: custData, error: custErr }, { data: orderData, error: orderErr }, { data: fuData, error: fuErr }] = await Promise.all([
      supabase.from("customers").select("id, name, phone, email, address").eq("id", customerId).single(),
      supabase
        .from("orders")
        .select(
          `id, order_date, pipeline_stage,
           salespeople(id, name),
           order_service_lines(id, account_number, plan_name, services(name), providers(name))`
        )
        .eq("customer_id", customerId)
        .order("order_date", { ascending: false }),
      supabase
        .from("follow_ups")
        .select("id, due_date, reason, status, completed_at")
        .eq("customer_id", customerId)
        .order("due_date", { ascending: true }),
    ]);
    detailCustomer = custErr ? null : custData;
    detailOrders = orderErr ? [] : orderData || [];
    detailFollowUps = fuErr ? [] : fuData || [];
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
        <td><button type="button" class="btn-link" data-customer="${f.customers?.id || ""}">${escapeHtml(f.customers?.name || "-")}</button></td>
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

  function detailModalHtml() {
    if (!detailCustomerId) return "";
    if (!detailCustomer) {
      return `
        <div class="modal-overlay" id="detail-overlay">
          <div class="modal-card">
            <div class="modal-head"><h3>Loading...</h3><button class="btn small" id="detail-close-btn">Close</button></div>
          </div>
        </div>`;
    }
    const c = detailCustomer;
    const pendingFu = detailFollowUps.filter((f) => f.status === "pending");
    const doneFu = detailFollowUps.filter((f) => f.status === "done");
    return `
      <div class="modal-overlay" id="detail-overlay">
        <div class="modal-card">
          <div class="modal-head">
            <h3>${escapeHtml(c.name || "-")}</h3>
            <button class="btn small" id="detail-close-btn">Close</button>
          </div>
          <div class="grid3">
            <div><span class="muted">Phone</span><div>${escapeHtml(c.phone || "-")}</div></div>
            <div><span class="muted">Email</span><div>${escapeHtml(c.email || "-")}</div></div>
            <div><span class="muted">Address</span><div>${escapeHtml(c.address || "-")}</div></div>
          </div>

          <h4>Orders</h4>
          <table class="data-table small">
            <thead><tr><th>Order Date</th><th>Salesperson</th><th>Status</th><th>Services</th></tr></thead>
            <tbody>
              ${
                detailOrders
                  .map(
                    (o) => `
                <tr>
                  <td>${escapeHtml(o.order_date || "-")}</td>
                  <td>${escapeHtml(o.salespeople?.name || "-")}</td>
                  <td>${escapeHtml(o.pipeline_stage || "-")}</td>
                  <td>${
                    (o.order_service_lines || [])
                      .map(
                        (l) =>
                          `${escapeHtml(l.services?.name || "?")}/${escapeHtml(l.providers?.name || "?")}${
                            l.account_number ? " (" + escapeHtml(l.account_number) + ")" : ""
                          }`
                      )
                      .join(", ") || "-"
                  }</td>
                </tr>`
                  )
                  .join("") || `<tr><td colspan="4" class="muted">No orders.</td></tr>`
              }
            </tbody>
          </table>

          <h4>Follow-ups needed</h4>
          <div id="followups-list">
            ${
              pendingFu.length === 0
                ? `<p class="muted">No pending follow-ups for this customer.</p>`
                : pendingFu
                    .map(
                      (f) => `
              <div class="note-item">
                <div class="note-meta"><span class="${BUCKET_BADGE[bucketFor(f.due_date)]}">Due ${escapeHtml(f.due_date)}</span></div>
                <div class="note-text">${escapeHtml(f.reason || "(no reason given)")}</div>
              </div>`
                    )
                    .join("")
            }
            ${
              doneFu.length > 0
                ? `<p class="muted" style="margin-top:10px">Completed (${doneFu.length}):</p>` +
                  doneFu
                    .map(
                      (f) => `
              <div class="note-item">
                <div class="note-meta"><span class="badge ok">Done${
                  f.completed_at ? " " + escapeHtml(new Date(f.completed_at).toLocaleDateString()) : ""
                }</span></div>
                <div class="note-text">${escapeHtml(f.reason || "-")}</div>
              </div>`
                    )
                    .join("")
                : ""
            }
          </div>

          <h4>Add Follow-up</h4>
          <div class="inline-form">
            <label>Due Date<input type="date" id="followup-date-input" value="${escapeHtml(newFollowUpDate || "")}" /></label>
            <button type="button" class="btn small" id="followup-quick-1w">+1 Week</button>
            <button type="button" class="btn small" id="followup-quick-2w">+2 Weeks</button>
          </div>
          <div class="inline-form">
            <textarea id="followup-reason-input" placeholder="Why is a follow-up needed? (e.g. confirm credit was applied, call back about...)" rows="2" style="flex:1">${escapeHtml(newFollowUpReason || "")}</textarea>
            <button class="btn primary" id="followup-add-btn" ${followUpBusy ? "disabled" : ""}>${followUpBusy ? "Saving..." : "Add Follow-up"}</button>
          </div>
        </div>
      </div>
    `;
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
        <p class="muted">Customers who need a follow-up after their order (e.g. confirming a credit was applied, or a callback) -- grouped by how soon it's due. Click a customer's name to see their full details and everything pending for them.</p>
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
      ${detailModalHtml()}
    `;

    container.querySelector("#show-completed").addEventListener("change", (e) => {
      showCompleted = e.target.checked;
      draw();
    });

    container.querySelectorAll("[data-customer]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const customerId = btn.dataset.customer;
        if (!customerId) return;
        detailCustomerId = customerId;
        detailCustomer = null;
        detailOrders = [];
        detailFollowUps = [];
        newFollowUpDate = "";
        newFollowUpReason = "";
        draw();
        await loadDetail(customerId);
        draw();
      });
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

    wireDetailModal();
  }

  function wireDetailModal() {
    const overlay = container.querySelector("#detail-overlay");
    if (!overlay) return;
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        detailCustomerId = null;
        draw();
      }
    });
    container.querySelector("#detail-close-btn")?.addEventListener("click", () => {
      detailCustomerId = null;
      draw();
    });
    container.querySelector(".modal-card")?.addEventListener("click", (e) => e.stopPropagation());

    container.querySelector("#followup-quick-1w")?.addEventListener("click", () => {
      const input = container.querySelector("#followup-date-input");
      newFollowUpDate = addDays(input.value || todayStr(), 7);
      draw();
    });
    container.querySelector("#followup-quick-2w")?.addEventListener("click", () => {
      const input = container.querySelector("#followup-date-input");
      newFollowUpDate = addDays(input.value || todayStr(), 14);
      draw();
    });
    container.querySelector("#followup-date-input")?.addEventListener("change", (e) => {
      newFollowUpDate = e.target.value;
    });
    container.querySelector("#followup-reason-input")?.addEventListener("input", (e) => {
      newFollowUpReason = e.target.value;
    });
    container.querySelector("#followup-add-btn")?.addEventListener("click", async () => {
      const dateInput = container.querySelector("#followup-date-input");
      const reasonInput = container.querySelector("#followup-reason-input");
      const dueDate = dateInput.value;
      const reason = reasonInput.value.trim();
      if (!dueDate) {
        alert("Pick a due date first.");
        return;
      }
      if (!detailCustomerId) return;
      followUpBusy = true;
      draw();
      const { error } = await supabase.from("follow_ups").insert({
        customer_id: detailCustomerId,
        due_date: dueDate,
        reason: reason || null,
        created_by: ctx.profile?.id || null,
      });
      if (error) {
        alert("Failed to save follow-up: " + error.message);
      } else {
        newFollowUpDate = "";
        newFollowUpReason = "";
      }
      await Promise.all([load(), loadDetail(detailCustomerId)]);
      followUpBusy = false;
      draw();
    });
  }

  await load();
  draw();
}
