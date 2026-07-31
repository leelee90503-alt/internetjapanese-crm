import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import { escapeHtml, fmtMoney, normalizeAccountNumber } from "./normalize.js";

// Missing Commission: tracks commission the company believes it has NOT
// received yet (from a manually-maintained list, e.g. an Excel sheet sent
// back and forth with the provider/agency). Separate from the automatic
// commission_report_rows matching flow -- this is for cases that need
// manual follow-up (chargebacks, "not yet received", account number
// mix-ups after an install issue, etc).
//
// Each item can optionally be linked to one or more order_service_lines
// (matched_line_ids). Linking is informational + lets staff jump to the
// order; it does not by itself change the order line's status. Marking an
// item "Received" here is a separate, explicit action that also updates
// the linked order line(s) to status='received'.

function statusLabel(item) {
  if (item.resolved) return "Received";
  if (item.needs_review) return "Needs review";
  return "Missing";
}

function downloadAsExcel(items) {
  const rows = items.map((item) => ({
    Status: statusLabel(item),
    Customer: item.customer_name || "",
    "Account #": item.account_number || "",
    Missing: item.description || "",
    Price: item.price === null || item.price === undefined ? "" : Number(item.price),
    "Order #": item.source_order_number || "",
    "Sales Date": item.sales_date || "",
    "Claimed To Provider": item.claimed_date_raw || "",
    "Status / Notes": item.status_notes || "",
    "Linked Lines": (item.matched_line_ids || []).length,
    "Review Note": item.review_note || "",
  }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet["!cols"] = [
    { wch: 12 }, // Status
    { wch: 20 }, // Customer
    { wch: 18 }, // Account #
    { wch: 24 }, // Missing
    { wch: 10 }, // Price
    { wch: 16 }, // Order #
    { wch: 12 }, // Sales Date
    { wch: 18 }, // Claimed To Provider
    { wch: 30 }, // Status / Notes
    { wch: 12 }, // Linked Lines
    { wch: 40 }, // Review Note
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Missing Commission");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Missing_Commission_${dateStr}.xlsx`);
}

function emptyForm() {
  return {
    customer_name: "",
    account_number: "",
    description: "",
    source_order_number: "",
    sales_date: "",
    price: "",
    claimed_date_raw: "",
    status_notes: "",
  };
}

export async function renderMissingCommission(container, ctx) {
  let items = [];
  let loadError = null;
  let showResolved = false;
  let expandedId = null;
  let searchResults = {}; // itemId -> array of candidate order_service_lines
  let searchAccount = {}; // itemId -> current search input value
  let addFormOpen = false;
  let addForm = emptyForm();
  let busy = false;

  async function load() {
    const { data, error } = await supabase
      .from("missing_commission_items")
      .select("*")
      .order("resolved", { ascending: true })
      .order("needs_review", { ascending: false })
      .order("created_at", { ascending: false });
    if (error) {
      loadError = error.message;
      items = [];
      return;
    }
    loadError = null;
    items = data || [];
  }

  async function fetchLineDetails(ids) {
    if (!ids || ids.length === 0) return [];
    const { data, error } = await supabase
      .from("order_service_lines")
      .select("id, account_number, plan_name, status, expected_commission, actual_commission_amount, orders(customers(name))")
      .in("id", ids);
    if (error) return [];
    return data || [];
  }

  function badgeForItem(item) {
    if (item.resolved) return `<span class="badge ok">Received</span>`;
    if (item.needs_review) return `<span class="badge warn">Needs review</span>`;
    return `<span class="badge error">Missing</span>`;
  }

  function rowHtml(item) {
    const lineCount = (item.matched_line_ids || []).length;
    return `
      <tr data-id="${item.id}" class="${item.needs_review && !item.resolved ? "needs-review-row" : ""}">
        <td>${badgeForItem(item)}</td>
        <td><button type="button" class="btn-link" data-toggle="${item.id}">${escapeHtml(item.customer_name || "-")}</button></td>
        <td>${escapeHtml(item.account_number || "-")}</td>
        <td>${escapeHtml(item.description || "-")}</td>
        <td>${fmtMoney(item.price)}</td>
        <td>${escapeHtml(item.sales_date || "-")}</td>
        <td>${escapeHtml(item.claimed_date_raw || "-")}</td>
        <td>${lineCount > 0 ? `<span class="badge neutral">${lineCount} linked</span>` : `<span class="badge error">none linked</span>`}</td>
        <td>
          <button class="btn small" data-toggle="${item.id}">${expandedId === item.id ? "Hide" : "Details"}</button>
          ${
            !item.resolved
              ? `<button class="btn small primary" data-mark-received="${item.id}" ${busy ? "disabled" : ""}>Mark Received</button>`
              : `<button class="btn small" data-reopen="${item.id}" ${busy ? "disabled" : ""}>Reopen</button>`
          }
          <button class="btn small danger" data-delete="${item.id}" ${busy ? "disabled" : ""}>Delete</button>
        </td>
      </tr>
      ${expandedId === item.id ? detailRowHtml(item) : ""}
    `;
  }

  function detailRowHtml(item) {
    const lines = searchResults[`linked:${item.id}`] || [];
    const candidates = searchResults[item.id] || [];
    const linkedIds = new Set(item.matched_line_ids || []);
    return `
      <tr class="detail-row">
        <td colspan="9">
          <div class="card">
            <div class="grid3">
              <div><span class="muted">Order #</span><div>${escapeHtml(item.source_order_number || "-")}</div></div>
              <div><span class="muted">Status / Notes (from source list)</span><div>${escapeHtml(item.status_notes || "-")}</div></div>
              <div><span class="muted">Created</span><div>${item.created_at ? escapeHtml(new Date(item.created_at).toLocaleDateString()) : "-"}</div></div>
            </div>
            ${
              item.review_note
                ? `<div class="alert info"><strong>Review note:</strong> ${escapeHtml(item.review_note)}</div>`
                : ""
            }

            <h4>Linked order lines</h4>
            ${
              lines.length === 0
                ? `<p class="muted">${linkedIds.size === 0 ? "No order lines linked yet." : "Loading..."}</p>`
                : `<table class="data-table small">
                    <thead><tr><th>Customer</th><th>Account #</th><th>Plan</th><th>Status</th><th>Expected</th><th>Actual</th><th></th></tr></thead>
                    <tbody>
                      ${lines
                        .map(
                          (l) => `
                        <tr>
                          <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
                          <td>${escapeHtml(l.account_number || "-")}</td>
                          <td>${escapeHtml(l.plan_name || "-")}</td>
                          <td>${escapeHtml(l.status)}</td>
                          <td>${fmtMoney(l.expected_commission)}</td>
                          <td>${fmtMoney(l.actual_commission_amount)}</td>
                          <td><button class="btn small" data-unlink="${item.id}:${l.id}">Unlink</button></td>
                        </tr>`
                        )
                        .join("")}
                    </tbody>
                  </table>`
            }

            <h4>Link another order line</h4>
            <div class="inline-form">
              <input type="text" placeholder="Account number" value="${escapeHtml(searchAccount[item.id] || item.account_number || "")}" data-search-input="${item.id}" />
              <button class="btn small" data-search-btn="${item.id}">Search</button>
            </div>
            ${
              candidates.length > 0
                ? `<table class="data-table small">
                    <thead><tr><th>Customer</th><th>Account #</th><th>Plan</th><th>Status</th><th>Expected</th><th></th></tr></thead>
                    <tbody>
                      ${candidates
                        .map(
                          (l) => `
                        <tr>
                          <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
                          <td>${escapeHtml(l.account_number || "-")}</td>
                          <td>${escapeHtml(l.plan_name || "-")}</td>
                          <td>${escapeHtml(l.status)}</td>
                          <td>${fmtMoney(l.expected_commission)}</td>
                          <td>${
                            linkedIds.has(l.id)
                              ? `<span class="muted">already linked</span>`
                              : `<button class="btn small primary" data-link="${item.id}:${l.id}">Link</button>`
                          }</td>
                        </tr>`
                        )
                        .join("")}
                    </tbody>
                  </table>`
                : ""
            }
          </div>
        </td>
      </tr>
    `;
  }

  function addFormHtml() {
    if (!addFormOpen) return "";
    return `
      <div class="card">
        <h4>Add Missing Commission Item</h4>
        <div class="inline-form">
          <input type="text" placeholder="Customer name" id="mc-name" value="${escapeHtml(addForm.customer_name)}" />
          <input type="text" placeholder="Account number" id="mc-account" value="${escapeHtml(addForm.account_number)}" />
          <input type="text" placeholder="What's missing (e.g. Mobile line)" id="mc-desc" value="${escapeHtml(addForm.description)}" />
        </div>
        <div class="inline-form">
          <input type="text" placeholder="Order #" id="mc-order" value="${escapeHtml(addForm.source_order_number)}" />
          <label>Sales date<input type="date" id="mc-sales-date" value="${escapeHtml(addForm.sales_date)}" /></label>
          <input type="number" step="0.01" placeholder="Price" id="mc-price" value="${escapeHtml(addForm.price)}" />
        </div>
        <div class="inline-form">
          <input type="text" placeholder="Claimed to provider (date or note)" id="mc-claimed" value="${escapeHtml(addForm.claimed_date_raw)}" />
          <input type="text" placeholder="Status / notes" id="mc-status" style="flex:1" value="${escapeHtml(addForm.status_notes)}" />
        </div>
        <div class="btn-row">
          <button class="btn primary" id="mc-save-btn" ${busy ? "disabled" : ""}>${busy ? "Saving..." : "Save"}</button>
          <button class="btn" id="mc-cancel-btn">Cancel</button>
        </div>
      </div>
    `;
  }

  function draw() {
    const visible = items.filter((i) => showResolved || !i.resolved);
    const needsReviewCount = items.filter((i) => i.needs_review && !i.resolved).length;
    const missingCount = items.filter((i) => !i.resolved).length;
    const resolvedCount = items.filter((i) => i.resolved).length;

    container.innerHTML = `
      <div class="screen">
        <h2>Missing Commission</h2>
        <p class="muted">
          Commission we believe has NOT been received yet, tracked separately from the automatic Commission Report matching.
          Link an item to the actual order line(s) it corresponds to, and mark it "Received" once the commission comes in
          (this also updates the linked order line's status).
        </p>

        ${loadError ? `<div class="alert error">Failed to load: ${escapeHtml(loadError)}</div>` : ""}

        <div class="inline-form">
          <span class="badge error">${missingCount} outstanding</span>
          <span class="badge warn">${needsReviewCount} need review</span>
          <span class="badge ok">${resolvedCount} received</span>
          <label class="checkbox-inline"><input type="checkbox" id="show-resolved" ${showResolved ? "checked" : ""} /> Show received</label>
          <button class="btn small primary" id="mc-add-open-btn">+ Add Missing Item</button>
          <button class="btn small" id="mc-download-btn">Download Excel</button>
        </div>

        ${addFormHtml()}

        <table class="data-table">
          <thead>
            <tr>
              <th>Status</th><th>Customer</th><th>Account #</th><th>Missing</th><th>Price</th>
              <th>Sales Date</th><th>Claimed To Provider</th><th>Linked Lines</th><th></th>
            </tr>
          </thead>
          <tbody>
            ${visible.map(rowHtml).join("") || `<tr><td colspan="9" class="muted">Nothing here.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;

    wireEvents();
  }

  function wireEvents() {
    container.querySelector("#show-resolved")?.addEventListener("change", (e) => {
      showResolved = e.target.checked;
      draw();
    });

    container.querySelector("#mc-download-btn")?.addEventListener("click", () => {
      const visible = items.filter((i) => showResolved || !i.resolved);
      if (visible.length === 0) {
        alert("Nothing to download.");
        return;
      }
      downloadAsExcel(visible);
    });

    container.querySelector("#mc-add-open-btn")?.addEventListener("click", () => {
      addFormOpen = !addFormOpen;
      addForm = emptyForm();
      draw();
    });
    container.querySelector("#mc-cancel-btn")?.addEventListener("click", () => {
      addFormOpen = false;
      draw();
    });
    container.querySelector("#mc-save-btn")?.addEventListener("click", async () => {
      addForm.customer_name = container.querySelector("#mc-name")?.value.trim() || "";
      addForm.account_number = container.querySelector("#mc-account")?.value.trim() || "";
      addForm.description = container.querySelector("#mc-desc")?.value.trim() || "";
      addForm.source_order_number = container.querySelector("#mc-order")?.value.trim() || "";
      addForm.sales_date = container.querySelector("#mc-sales-date")?.value || "";
      addForm.price = container.querySelector("#mc-price")?.value || "";
      addForm.claimed_date_raw = container.querySelector("#mc-claimed")?.value.trim() || "";
      addForm.status_notes = container.querySelector("#mc-status")?.value.trim() || "";
      if (!addForm.customer_name) {
        alert("Customer name is required.");
        return;
      }
      busy = true;
      draw();
      const { error } = await supabase.from("missing_commission_items").insert({
        customer_name: addForm.customer_name,
        account_number: addForm.account_number || null,
        description: addForm.description || null,
        source_order_number: addForm.source_order_number || null,
        sales_date: addForm.sales_date || null,
        price: addForm.price === "" ? null : Number(addForm.price),
        claimed_date_raw: addForm.claimed_date_raw || null,
        status_notes: addForm.status_notes || null,
        matched_line_ids: [],
        needs_review: true,
        review_note: "Manually added -- not yet linked to an order line.",
        created_by: ctx.profile?.id || null,
      });
      busy = false;
      if (error) {
        alert("Failed to save: " + error.message);
        draw();
        return;
      }
      addFormOpen = false;
      addForm = emptyForm();
      await load();
      draw();
    });

    container.querySelectorAll("[data-toggle]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.toggle;
        if (expandedId === id) {
          expandedId = null;
          draw();
          return;
        }
        expandedId = id;
        const item = items.find((i) => i.id === id);
        draw();
        if (item && (item.matched_line_ids || []).length > 0) {
          const lines = await fetchLineDetails(item.matched_line_ids);
          searchResults[`linked:${id}`] = lines;
          draw();
        }
      });
    });

    container.querySelectorAll("[data-search-input]").forEach((input) => {
      input.addEventListener("input", (e) => {
        searchAccount[input.dataset.searchInput] = e.target.value;
      });
    });
    container.querySelectorAll("[data-search-btn]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.searchBtn;
        const raw = container.querySelector(`[data-search-input="${id}"]`)?.value || "";
        const normalized = normalizeAccountNumber(raw);
        if (!normalized) {
          alert("Enter an account number to search.");
          return;
        }
        const { data, error } = await supabase
          .from("order_service_lines")
          .select("id, account_number, plan_name, status, expected_commission, orders(customers(name))")
          .ilike("account_number", `%${normalized}%`)
          .limit(25);
        if (error) {
          alert("Search failed: " + error.message);
          return;
        }
        searchResults[id] = data || [];
        draw();
        expandedId = id;
      });
    });

    container.querySelectorAll("[data-link]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const [itemId, lineId] = btn.dataset.link.split(":");
        const item = items.find((i) => i.id === itemId);
        if (!item) return;
        const newIds = Array.from(new Set([...(item.matched_line_ids || []), lineId]));
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").update({ matched_line_ids: newIds }).eq("id", itemId);
        busy = false;
        if (error) {
          alert("Failed to link: " + error.message);
        }
        await load();
        expandedId = itemId;
        const refreshed = items.find((i) => i.id === itemId);
        if (refreshed) searchResults[`linked:${itemId}`] = await fetchLineDetails(refreshed.matched_line_ids);
        draw();
      });
    });

    container.querySelectorAll("[data-unlink]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const [itemId, lineId] = btn.dataset.unlink.split(":");
        const item = items.find((i) => i.id === itemId);
        if (!item) return;
        const newIds = (item.matched_line_ids || []).filter((x) => x !== lineId);
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").update({ matched_line_ids: newIds }).eq("id", itemId);
        busy = false;
        if (error) {
          alert("Failed to unlink: " + error.message);
        }
        await load();
        expandedId = itemId;
        const refreshed = items.find((i) => i.id === itemId);
        if (refreshed) searchResults[`linked:${itemId}`] = await fetchLineDetails(refreshed.matched_line_ids);
        draw();
      });
    });

    container.querySelectorAll("[data-mark-received]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.markReceived;
        const item = items.find((i) => i.id === id);
        if (!item) return;
        const lineIds = item.matched_line_ids || [];
        if (!confirm(`Mark "${item.customer_name}" - ${item.description} as received?${lineIds.length ? ` This will also set ${lineIds.length} linked order line(s) to "received".` : ""}`)) {
          return;
        }
        busy = true;
        draw();
        if (lineIds.length > 0) {
          const { data: lines } = await supabase.from("order_service_lines").select("id, expected_commission").in("id", lineIds);
          for (const l of lines || []) {
            await supabase
              .from("order_service_lines")
              .update({
                status: "received",
                actual_commission_amount: l.expected_commission,
                commission_matched_at: new Date().toISOString(),
                commission_matched_by: ctx.profile?.id || null,
              })
              .eq("id", l.id);
          }
        }
        const { error } = await supabase.from("missing_commission_items").update({ resolved: true }).eq("id", id);
        busy = false;
        if (error) {
          alert("Failed to mark received: " + error.message);
        }
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-reopen]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.reopen;
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").update({ resolved: false }).eq("id", id);
        busy = false;
        if (error) alert("Failed to reopen: " + error.message);
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-delete]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.delete;
        const item = items.find((i) => i.id === id);
        if (!confirm(`Delete this missing-commission item for "${item?.customer_name || ""}"? This does not change any order line.`)) return;
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").delete().eq("id", id);
        busy = false;
        if (error) alert("Failed to delete: " + error.message);
        await load();
        draw();
      });
    });
  }

  await load();
  draw();
}
