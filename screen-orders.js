import { supabase } from "./supabaseClient.js";
import { escapeHtml, fmtMoney } from "./normalize.js";

function commissionBadge(status) {
  if (status === "received") return `<span class="badge ok">Received</span>`;
  if (status === "mismatch") return `<span class="badge error">Mismatch</span>`;
  if (status === "no_report") return `<span class="badge neutral">No Report</span>`;
  return `<span class="badge warn">Pending</span>`;
}

export async function renderOrders(container, ctx) {
  const isAdmin = ctx.profile?.role === "admin";
  let search = "";
  let orders = [];
  let selected = new Set();
  let detailOrderId = null;
  let detailNotes = [];
  let detailBusy = false;
  let noteBusy = false;

  async function load() {
    let query = supabase
      .from("orders")
      .select(
        `id, order_date, pipeline_stage, memo, source_order_id, order_number, category, created_at,
         customers(id, name, phone, address, email),
         salespeople(id, name),
         order_service_lines(id, account_number, expected_commission, actual_commission_amount, status,
           units, units_installed, mobile_lines_ordered, mobile_lines_installed, plan_name,
           services(id, name), providers(id, name))`
      )
      .order("order_date", { ascending: false });
    const { data, error } = await query;
    if (error) {
      container.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
      return;
    }
    orders = data || [];
  }

  function filtered() {
    if (!search.trim()) return orders;
    const q = search.trim().toLowerCase();
    return orders.filter(
      (o) =>
        (o.customers?.name || "").toLowerCase().includes(q) ||
        (o.customers?.phone || "").toLowerCase().includes(q) ||
        (o.salespeople?.name || "").toLowerCase().includes(q) ||
        o.order_service_lines?.some((l) => (l.account_number || "").toLowerCase().includes(q))
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

  function detailModalHtml() {
    if (!detailOrderId) return "";
    const o = orders.find((x) => x.id === detailOrderId);
    if (!o) return "";
    const c = o.customers || {};
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
            <div><span class="muted">Order Date</span><div>${escapeHtml(o.order_date || "-")}</div></div>
            <div><span class="muted">Salesperson</span><div>${escapeHtml(o.salespeople?.name || "-")}</div></div>
            <div><span class="muted">Status</span><div>${escapeHtml(o.pipeline_stage || "-")}</div></div>
            <div><span class="muted">Order Id (source)</span><div>${escapeHtml(o.source_order_id || "-")}</div></div>
            <div><span class="muted">Order Number / Work Order</span><div>${escapeHtml(o.order_number || "-")}</div></div>
            <div><span class="muted">Category</span><div>${escapeHtml(o.category || "-")}</div></div>
          </div>

          <h4>Service Lines</h4>
          <table class="data-table small">
            <thead>
              <tr><th>Service</th><th>Provider</th><th>Account #</th><th>Units</th><th>Units Installed</th><th>Mobile Lines</th><th>Plan</th><th>Commission</th></tr>
            </thead>
            <tbody>
              ${
                (o.order_service_lines || [])
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
        </div>
      </div>
    `;
  }

  function draw() {
    const rows = filtered();
    const allSelected = rows.length > 0 && rows.every((o) => selected.has(o.id));
    container.innerHTML = `
      <div class="screen">
        <h2>Customers / Orders</h2>
        <p class="muted">All staff can view all customer/order data. Click a row to see full details and notes. (Commission-related screens are admin-only)</p>
        <input type="text" id="search" placeholder="Search by name/phone/salesperson/account number" value="${escapeHtml(search)}" />
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
                  (o) => `
              <tr data-order-row="${o.id}" class="clickable-row">
                ${isAdmin ? `<td><input type="checkbox" class="row-select" data-id="${o.id}" ${selected.has(o.id) ? "checked" : ""} /></td>` : ""}
                <td>${o.order_date || "-"}</td>
                <td>${escapeHtml(o.customers?.name || "-")}</td>
                <td>${escapeHtml(o.customers?.phone || "-")}</td>
                <td>${escapeHtml(o.salespeople?.name || "-")}</td>
                <td>${(o.order_service_lines || [])
                  .map(
                    (l) =>
                      `${escapeHtml(l.services?.name || "?")}/${escapeHtml(l.providers?.name || "?")}${
                        l.account_number ? " (" + escapeHtml(l.account_number) + ")" : ""
                      }`
                  )
                  .join(", ") || "-"}</td>
                <td>${escapeHtml(o.pipeline_stage)}</td>
                ${
                  isAdmin
                    ? `<td>${
                        (o.order_service_lines || [])
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
      ${detailModalHtml()}
    `;

    container.querySelector("#search").addEventListener("input", (e) => {
      search = e.target.value;
      draw();
    });

    if (isAdmin) {
      container.querySelector("#select-all")?.addEventListener("change", (e) => {
        if (e.target.checked) rows.forEach((o) => selected.add(o.id));
        else rows.forEach((o) => selected.delete(o.id));
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
        const ok = confirm(`Delete ${selected.size} selected order(s)? This cannot be undone.`);
        if (!ok) return;
        await deleteOrders(Array.from(selected));
      });
      container.querySelector("#delete-all-btn")?.addEventListener("click", async () => {
        if (rows.length === 0) return;
        const ok = confirm(`Delete ALL ${rows.length} order(s) currently shown? This cannot be undone.`);
        if (!ok) return;
        await deleteOrders(rows.map((o) => o.id));
      });
    }

    container.querySelectorAll("tr[data-order-row]").forEach((tr) => {
      tr.addEventListener("click", async () => {
        detailOrderId = tr.dataset.orderRow;
        detailBusy = true;
        draw();
        const o = orders.find((x) => x.id === detailOrderId);
        await loadNotesFor(o?.customers?.id);
        detailBusy = false;
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
        detailOrderId = null;
        draw();
      }
    });
    container.querySelector("#detail-close-btn")?.addEventListener("click", () => {
      detailOrderId = null;
      draw();
    });
    container.querySelector(".modal-card")?.addEventListener("click", (e) => e.stopPropagation());
    container.querySelector("#note-add-btn")?.addEventListener("click", async () => {
      const input = container.querySelector("#note-input");
      const text = input.value.trim();
      if (!text) return;
      const o = orders.find((x) => x.id === detailOrderId);
      const customerId = o?.customers?.id;
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
  }

  async function deleteOrders(ids) {
    const { error } = await supabase.from("orders").delete().in("id", ids);
    if (error) {
      alert("Failed to delete: " + error.message);
      return;
    }
    selected.clear();
    if (detailOrderId && ids.includes(detailOrderId)) detailOrderId = null;
    await load();
    draw();
  }

  await load();
  draw();
}
