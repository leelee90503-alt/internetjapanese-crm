import { supabase } from "./supabaseClient.js";
import { escapeHtml, fmtMoney } from "./normalize.js";

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

function commissionBadge(status) {
  if (status === "received") return `<span class="badge ok">Received</span>`;
  if (status === "mismatch") return `<span class="badge error">Commission Mismatch</span>`;
  if (status === "no_report") return `<span class="badge neutral">No Report</span>`;
  if (status === "missing") return `<span class="badge error">Missing</span>`;
  return `<span class="badge warn">Pending</span>`;
}

export async function renderOrders(container, ctx) {
  const isAdmin = ctx.profile?.role === "admin";
  let search = "";
  let searchDraft = ""; // text typed into the search box, not applied until "Search" is clicked
  let customers = [];
  let selected = new Set();
  let detailKey = null;
  let detailNotes = [];
  let detailFollowUps = [];
  let detailBusy = false;
  let noteBusy = false;
  let followUpBusy = false;
  let newFollowUpDate = "";
  let newFollowUpReason = "";

  // The list is built from `customers` (not `orders`) so that a customer
  // saved without a successful order/line save (e.g. a dropped connection
  // mid-upload) still shows up and can be reviewed or deleted, instead of
  // being invisible on this screen while still counted on the Dashboard.
  async function load() {
    const { data, error } = await supabase
      .from("customers")
      .select(
        `id, name, phone, address, email, created_at,
         orders(id, order_date, pipeline_stage, memo, source_order_id, order_number, category,
           salespeople(id, name),
           order_service_lines(id, account_number, expected_commission, actual_commission_amount, status,
             units, units_installed, mobile_lines_ordered, mobile_lines_installed, plan_name,
             services(id, name), providers(id, name)))`
      )
      .order("created_at", { ascending: false });
    if (error) {
      container.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
      customers = [];
      return;
    }
    customers = data || [];
  }

  // One row per order; customers with zero orders get a single placeholder
  // row (order: null) so they're still visible/selectable/deletable.
  function buildRows() {
    const out = [];
    for (const c of customers) {
      const orders = (c.orders || []).slice().sort((a, b) => (b.order_date || "").localeCompare(a.order_date || ""));
      if (orders.length === 0) {
        out.push({ rowKey: `c:${c.id}`, customer: c, order: null });
      } else {
        for (const o of orders) {
          out.push({ rowKey: `o:${o.id}`, customer: c, order: o });
        }
      }
    }
    return out;
  }

  function filtered() {
    const rows = buildRows();
    if (!search.trim()) return rows;
    const q = search.trim().toLowerCase();
    return rows.filter(
      (r) =>
        (r.customer.name || "").toLowerCase().includes(q) ||
        (r.customer.phone || "").toLowerCase().includes(q) ||
        (r.order?.salespeople?.name || "").toLowerCase().includes(q) ||
        r.order?.order_service_lines?.some((l) => (l.account_number || "").toLowerCase().includes(q))
    );
  }

  async function loadNotesFor(customerId) {
    if (!customerId) {
      detailNotes = [];
      return;
    }
    const { data, error } = await supabase
      .from("customer_notes")
      .select("id, note, author_name, created_at")
      .eq("customer_id", customerId)
      .order("created_at", { ascending: true });
    detailNotes = error ? [] : data || [];
  }

  async function loadFollowUpsFor(customerId) {
    if (!customerId) {
      detailFollowUps = [];
      return;
    }
    const { data, error } = await supabase
      .from("follow_ups")
      .select("id, due_date, reason, status, completed_at")
      .eq("customer_id", customerId)
      .order("due_date", { ascending: true });
    detailFollowUps = error ? [] : data || [];
  }

  function detailModalHtml(rows) {
    if (!detailKey) return "";
    const row = rows.find((r) => r.rowKey === detailKey);
    if (!row) return "";
    const c = row.customer;
    const o = row.order;
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
            <div><span class="muted">Order Date</span><div>${escapeHtml(o?.order_date || "-")}</div></div>
            <div><span class="muted">Salesperson</span><div>${escapeHtml(o?.salespeople?.name || "-")}</div></div>
            <div><span class="muted">Status</span><div>${o ? escapeHtml(o.pipeline_stage || "-") : `<span class="badge neutral">No Orders</span>`}</div></div>
            <div><span class="muted">Order Id (source)</span><div>${escapeHtml(o?.source_order_id || "-")}</div></div>
            <div><span class="muted">Order Number / Work Order</span><div>${escapeHtml(o?.order_number || "-")}</div></div>
            <div><span class="muted">Category</span><div>${escapeHtml(o?.category || "-")}</div></div>
          </div>
          ${
            !o
              ? `<div class="alert error" style="margin-top:10px">This customer has no saved order. This usually means the order/service-line save failed after the customer record was created (e.g. a dropped connection during upload). Delete this customer and re-upload, or contact an admin to investigate.</div>`
              : ""
          }

          <h4>Service Lines</h4>
          <table class="data-table small">
            <thead>
              <tr><th>Service</th><th>Provider</th><th>Account #</th><th>Units</th><th>Units Installed</th><th>Mobile Lines</th><th>Plan</th><th>Commission</th></tr>
            </thead>
            <tbody>
              ${
                (o?.order_service_lines || [])
                  .map(
                    (l) => `
                <tr>
                  <td>${escapeHtml(l.services?.name || "?")}</td>
                  <td>${escapeHtml(l.providers?.name || "?")}</td>
                  <td>${escapeHtml(l.account_number || "-")}</td>
                  <td>${l.units ?? "-"}</td>
                  <td>${l.units_installed ?? "-"}</td>
                  <td>${
                    l.mobile_lines_ordered != null || l.mobile_lines_installed != null
                      ? `${l.mobile_lines_ordered ?? "-"} / ${l.mobile_lines_installed ?? "-"}`
                      : "-"
                  }</td>
                  <td>${escapeHtml(l.plan_name || "-")}</td>
                  <td>${commissionBadge(l.status)} ${fmtMoney(l.actual_commission_amount)} / exp. ${fmtMoney(l.expected_commission)}</td>
                </tr>`
                  )
                  .join("") || `<tr><td colspan="8" class="muted">No service lines.</td></tr>`
              }
            </tbody>
          </table>

          <h4>Notes</h4>
          <div id="notes-list">
            ${
              detailNotes
                .map(
                  (n) => `
              <div class="note-item">
                <div class="note-meta">${escapeHtml(n.author_name || "Unknown")} &middot; ${escapeHtml(
                    new Date(n.created_at).toLocaleString()
                  )}</div>
                <div class="note-text">${escapeHtml(n.note)}</div>
              </div>`
                )
                .join("") || `<p class="muted">No notes yet.</p>`
            }
          </div>
          <div class="inline-form">
            <textarea id="note-input" placeholder="Add a note about this customer..." rows="2" style="flex:1"></textarea>
            <button class="btn primary" id="note-add-btn" ${noteBusy ? "disabled" : ""}>${noteBusy ? "Saving..." : "Add Note"}</button>
          </div>

          <h4>Follow-ups</h4>
          <div id="followups-list">
            ${
              detailFollowUps
                .map(
                  (f) => `
              <div class="note-item">
                <div class="note-meta">
                  ${
                    f.status === "done"
                      ? `<span class="badge ok">Done</span>`
                      : `<span class="badge warn">Due ${escapeHtml(f.due_date)}</span>`
                  }
                </div>
                <div class="note-text">${escapeHtml(f.reason || "-")}</div>
              </div>`
                )
                .join("") || `<p class="muted">No follow-ups yet.</p>`
            }
          </div>
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
    const rows = filtered();
    const allSelected = rows.length > 0 && rows.every((r) => selected.has(r.rowKey));
    const orphanCount = rows.filter((r) => !r.order).length;
    container.innerHTML = `
      <div class="screen">
        <h2>Customers / Orders</h2>
        <p class="muted">All staff can view all customer/order data. Click a row to see full details and notes. (Commission-related screens are admin-only)</p>
        ${
          orphanCount > 0
            ? `<div class="alert error">${orphanCount} customer(s) have no saved order (shown with a "No Orders" badge below) -- their order/service-line save likely failed after the customer record was created. Review and delete or re-upload as needed.</div>`
            : ""
        }
        <div class="inline-form">
          <input type="text" id="search" placeholder="Search by name/phone/salesperson/account number" value="${escapeHtml(searchDraft)}" style="flex:1" />
          <button class="btn" id="search-btn">Search</button>
          ${search.trim() ? `<button class="btn" id="search-clear-btn">Clear</button>` : ""}
        </div>
        ${
          isAdmin
            ? `<div class="btn-row">
                <button class="btn danger" id="delete-selected-btn" ${selected.size === 0 ? "disabled" : ""}>Delete Selected (${selected.size})</button>
                <button class="btn danger" id="delete-all-btn" ${rows.length === 0 ? "disabled" : ""}>Delete All (${rows.length})</button>
              </div>`
            : ""
        }
        <table class="data-table">
          <thead>
            <tr>
              ${isAdmin ? `<th><input type="checkbox" id="select-all" ${allSelected ? "checked" : ""} /></th>` : ""}
              <th>Order Date</th><th>Customer Name</th><th>Phone</th><th>Salesperson</th><th>Service Items</th><th>Status</th>${
                isAdmin ? "<th>Commission</th>" : ""
              }</tr>
          </thead>
          <tbody>
            ${
              rows
                .map(
                  (r) => `
              <tr data-row-key="${r.rowKey}" class="clickable-row ${!r.order ? "needs-review-row" : ""}">
                ${isAdmin ? `<td><input type="checkbox" class="row-select" data-id="${r.rowKey}" ${selected.has(r.rowKey) ? "checked" : ""} /></td>` : ""}
                <td>${r.order?.order_date || "-"}</td>
                <td>${escapeHtml(r.customer.name || "-")}</td>
                <td>${escapeHtml(r.customer.phone || "-")}</td>
                <td>${escapeHtml(r.order?.salespeople?.name || "-")}</td>
                <td>${
                  (r.order?.order_service_lines || [])
                    .map(
                      (l) =>
                        `${escapeHtml(l.services?.name || "?")}/${escapeHtml(l.providers?.name || "?")}${
                          l.account_number ? " (" + escapeHtml(l.account_number) + ")" : ""
                        }`
                    )
                    .join(", ") || "-"
                }</td>
                <td>${r.order ? escapeHtml(r.order.pipeline_stage) : `<span class="badge warn">No Orders</span>`}</td>
                ${
                  isAdmin
                    ? `<td>${
                        (r.order?.order_service_lines || [])
                          .map(
                            (l) =>
                              `${commissionBadge(l.status)} ${fmtMoney(l.actual_commission_amount)} / exp. ${fmtMoney(l.expected_commission)}`
                          )
                          .join("<br/>") || "-"
                      }</td>`
                    : ""
                }
              </tr>`
                )
                .join("") || `<tr><td colspan="${isAdmin ? 8 : 6}" class="muted">No data.</td></tr>`
            }
          </tbody>
        </table>
      </div>
      ${detailModalHtml(rows)}
    `;

    container.querySelector("#search").addEventListener("input", (e) => {
      searchDraft = e.target.value;
    });
    container.querySelector("#search").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        search = searchDraft;
        draw();
      }
    });
    container.querySelector("#search-btn")?.addEventListener("click", () => {
      search = searchDraft;
      draw();
    });
    container.querySelector("#search-clear-btn")?.addEventListener("click", () => {
      search = "";
      searchDraft = "";
      draw();
    });

    if (isAdmin) {
      container.querySelector("#select-all")?.addEventListener("change", (e) => {
        if (e.target.checked) rows.forEach((r) => selected.add(r.rowKey));
        else rows.forEach((r) => selected.delete(r.rowKey));
        draw();
      });
      container.querySelectorAll(".row-select").forEach((cb) => {
        cb.addEventListener("click", (e) => e.stopPropagation());
        cb.addEventListener("change", (e) => {
          const id = e.target.dataset.id;
          if (e.target.checked) selected.add(id);
          else selected.delete(id);
          draw();
        });
      });
      container.querySelector("#delete-selected-btn")?.addEventListener("click", async () => {
        if (selected.size === 0) return;
        const ok = confirm(`Delete ${selected.size} selected row(s)? This cannot be undone.`);
        if (!ok) return;
        await deleteRows(Array.from(selected));
      });
      container.querySelector("#delete-all-btn")?.addEventListener("click", async () => {
        if (rows.length === 0) return;
        const ok = confirm(`Delete ALL ${rows.length} row(s) currently shown? This cannot be undone.`);
        if (!ok) return;
        await deleteRows(rows.map((r) => r.rowKey));
      });
    }

    container.querySelectorAll("tr[data-row-key]").forEach((tr) => {
      tr.addEventListener("click", async () => {
        detailKey = tr.dataset.rowKey;
        detailBusy = true;
        newFollowUpDate = "";
        newFollowUpReason = "";
        draw();
        const row = rows.find((r) => r.rowKey === detailKey);
        await Promise.all([loadNotesFor(row?.customer?.id), loadFollowUpsFor(row?.customer?.id)]);
        detailBusy = false;
        draw();
      });
    });

    wireDetailModal(rows);
  }

  function wireDetailModal(rows) {
    const overlay = container.querySelector("#detail-overlay");
    if (!overlay) return;
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) {
        detailKey = null;
        draw();
      }
    });
    container.querySelector("#detail-close-btn")?.addEventListener("click", () => {
      detailKey = null;
      draw();
    });
    container.querySelector(".modal-card")?.addEventListener("click", (e) => e.stopPropagation());
    container.querySelector("#note-add-btn")?.addEventListener("click", async () => {
      const input = container.querySelector("#note-input");
      const text = input.value.trim();
      if (!text) return;
      const row = rows.find((r) => r.rowKey === detailKey);
      const customerId = row?.customer?.id;
      if (!customerId) return;
      noteBusy = true;
      draw();
      const authorName = ctx.profile?.full_name || ctx.profile?.email || "Unknown";
      const { error } = await supabase.from("customer_notes").insert({
        customer_id: customerId,
        author_id: ctx.profile?.id || null,
        author_name: authorName,
        note: text,
      });
      if (error) {
        alert("Failed to save note: " + error.message);
      }
      await loadNotesFor(customerId);
      noteBusy = false;
      draw();
    });

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
      const row = rows.find((r) => r.rowKey === detailKey);
      const customerId = row?.customer?.id;
      if (!customerId) return;
      followUpBusy = true;
      draw();
      const { error } = await supabase.from("follow_ups").insert({
        customer_id: customerId,
        order_id: row?.order?.id || null,
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
      await loadFollowUpsFor(customerId);
      followUpBusy = false;
      draw();
    });
  }

  // rowKeys look like "o:<orderId>" or "c:<customerId>" (a customer with no
  // order). Order rows delete via the orders table (cascades its lines);
  // customer-only rows delete via the customers table directly.
  async function deleteRows(rowKeys) {
    const orderIds = rowKeys.filter((k) => k.startsWith("o:")).map((k) => k.slice(2));
    const customerIds = rowKeys.filter((k) => k.startsWith("c:")).map((k) => k.slice(2));

    if (orderIds.length > 0) {
      const { error } = await supabase.from("orders").delete().in("id", orderIds);
      if (error) {
        alert("Failed to delete: " + error.message);
        return;
      }
    }
    if (customerIds.length > 0) {
      const { error } = await supabase.from("customers").delete().in("id", customerIds);
      if (error) {
        alert("Failed to delete: " + error.message);
        return;
      }
    }

    selected.clear();
    if (detailKey && rowKeys.includes(detailKey)) detailKey = null;
    await load();
    draw();
  }

  await load();
  draw();
}
