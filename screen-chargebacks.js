import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import { escapeHtml, fmtMoney } from "./normalize.js";

// Chargebacks: a read-only report of every confirmed commission_report_rows
// entry with a negative commission_amount -- i.e. a provider clawing back
// commission already paid (customer canceled early, etc). Confirming a
// chargeback row on the Commission Report screen deliberately does NOT
// overwrite the matched order_service_line's actual_commission_amount/status
// (see isChargebackRow() in screen-commissionReports.js) -- it's recorded
// here instead, linked back to the customer/product via
// matched_order_service_line_id, so the original "commission received"
// history stays intact while the chargeback itself is still visible.

const PAGE_SIZE = 50;

function downloadAsExcel(rows) {
  const out = rows.map((r) => ({
    Date: r.created_at ? new Date(r.created_at).toLocaleDateString() : "",
    Customer: r.line?.orders?.customers?.name || r.customer_name_raw || "",
    Phone: r.line?.orders?.customers?.phone || "",
    "Account #": r.account_number || "",
    Provider: r.line?.providers?.name || "",
    Product: r.line?.plan_name || r.raw_data?.commission_product_name || "",
    "Chargeback Amount": r.commission_amount === null || r.commission_amount === undefined ? "" : Number(r.commission_amount),
    "Original Order Date": r.line?.orders?.order_date || "",
    "Source File": r.commission_report_batches?.source_filename || "",
  }));
  const sheet = XLSX.utils.json_to_sheet(out);
  sheet["!cols"] = [
    { wch: 12 }, // Date
    { wch: 20 }, // Customer
    { wch: 14 }, // Phone
    { wch: 18 }, // Account #
    { wch: 12 }, // Provider
    { wch: 20 }, // Product
    { wch: 16 }, // Chargeback Amount
    { wch: 14 }, // Original Order Date
    { wch: 24 }, // Source File
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Chargebacks");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Chargebacks_${dateStr}.xlsx`);
}

export async function renderChargebacks(container, ctx) {
  let rows = [];
  let loadError = null;
  let providers = [];

  // filters
  let filterProvider = "";
  let filterSearch = "";
  let filterSearchDraft = "";
  let page = 0;

  async function load() {
    const { data: providerRows } = await supabase.from("providers").select("id, name").order("name");
    providers = providerRows || [];

    const { data, error } = await supabase
      .from("commission_report_rows")
      .select(
        `id, account_number, customer_name_raw, commission_amount, raw_data, match_status, created_at,
         commission_report_batches(source_filename),
         line:matched_order_service_line_id(id, plan_name, account_number, status, actual_commission_amount, expected_commission,
           providers(name), orders(order_date, customers(name, phone)))`
      )
      .lt("commission_amount", 0)
      .order("created_at", { ascending: false })
      .limit(2000);
    if (error) {
      loadError = error.message;
      rows = [];
      return;
    }
    loadError = null;
    rows = data || [];
  }

  function filtered() {
    let out = rows.slice();
    if (filterProvider) {
      out = out.filter((r) => (r.line?.providers?.name || "") === filterProvider);
    }
    if (filterSearch.trim()) {
      const needle = filterSearch.trim().toLowerCase();
      out = out.filter((r) => {
        const name = (r.line?.orders?.customers?.name || r.customer_name_raw || "").toLowerCase();
        const acct = (r.account_number || "").toString();
        const plan = (r.line?.plan_name || r.raw_data?.commission_product_name || "").toLowerCase();
        return name.includes(needle) || plan.includes(needle) || acct.includes(needle);
      });
    }
    return out;
  }

  function rowHtml(r) {
    const customer = r.line?.orders?.customers?.name || r.customer_name_raw || "-";
    const phone = r.line?.orders?.customers?.phone || "";
    const product = r.line?.plan_name || r.raw_data?.commission_product_name || "-";
    const provider = r.line?.providers?.name || "-";
    return `
      <tr>
        <td>${r.created_at ? new Date(r.created_at).toLocaleDateString() : "-"}</td>
        <td>${escapeHtml(customer)}${phone ? ` <span class="muted">(${escapeHtml(phone)})</span>` : ""}</td>
        <td>${escapeHtml(r.account_number || "-")}</td>
        <td>${escapeHtml(provider)}</td>
        <td>${escapeHtml(product)}</td>
        <td><span class="badge chargeback">${fmtMoney(r.commission_amount)}</span></td>
        <td>${
          r.line
            ? `Received: ${fmtMoney(r.line.actual_commission_amount)} <span class="muted">(status: ${escapeHtml(r.line.status || "-")})</span>`
            : `<span class="muted">Original line not found (may have been deleted)</span>`
        }</td>
        <td class="muted">${escapeHtml(r.commission_report_batches?.source_filename || "-")}</td>
        <td><button class="btn small danger" data-delete="${r.id}">Delete</button></td>
      </tr>
    `;
  }

  function draw() {
    const active = document.activeElement;
    const activeId = active && container.contains(active) ? active.id : null;

    const filteredRows = filtered();
    const totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
    if (page >= totalPages) page = totalPages - 1;
    const pageRows = filteredRows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);
    const totalChargeback = filteredRows.reduce((s, r) => s + (Number(r.commission_amount) || 0), 0);

    container.innerHTML = `
      <div class="screen">
        <h2>Chargebacks</h2>
        <p class="muted">
          Every commission report row with a negative amount -- a provider clawing back commission already paid
          (e.g. the customer canceled early). Confirming a chargeback on the Commission Report screen does not
          overwrite the customer's original received commission; it's recorded here instead, linked to the customer
          and product so you can see exactly who and what was charged back. If the same commission report is
          uploaded more than once, the same chargeback can end up recorded twice -- use Delete to remove the extra
          copy (deleting only removes this record; it never touches the customer's received commission).
        </p>

        ${loadError ? `<div class="alert error">Failed to load: ${escapeHtml(loadError)}</div>` : ""}

        <div class="inline-form">
          <span class="badge neutral">${filteredRows.length} chargeback(s)${filterProvider || filterSearch ? ` (filtered from ${rows.length})` : ""}</span>
          <span class="badge chargeback">Total charged back: ${fmtMoney(totalChargeback)}</span>
          <button class="btn small" id="cb-download-btn">Download Excel</button>
        </div>

        <div class="inline-form">
          <input type="text" placeholder="Search customer / account / product" id="cb-search" value="${escapeHtml(filterSearchDraft)}" style="flex:1" />
          <button class="btn" id="cb-search-btn">Search</button>
          ${filterSearch.trim() ? `<button class="btn" id="cb-search-clear-btn">Clear</button>` : ""}
          <select id="cb-provider">
            <option value="">All providers</option>
            ${providers.map((p) => `<option value="${escapeHtml(p.name)}" ${filterProvider === p.name ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
          </select>
        </div>

        <table class="data-table">
          <thead>
            <tr>
              <th>Date</th><th>Customer</th><th>Account #</th><th>Provider</th><th>Product</th>
              <th>Chargeback Amount</th><th>Original Commission</th><th>Source File</th><th></th>
            </tr>
          </thead>
          <tbody>
            ${pageRows.map(rowHtml).join("") || `<tr><td colspan="9" class="muted">No chargebacks${filterProvider || filterSearch ? " matching these filters" : ""}.</td></tr>`}
          </tbody>
        </table>

        ${
          totalPages > 1
            ? `<div class="inline-form">
                <button class="btn small" id="cb-prev" ${page === 0 ? "disabled" : ""}>&laquo; Prev</button>
                <span class="muted">Page ${page + 1} of ${totalPages}</span>
                <button class="btn small" id="cb-next" ${page >= totalPages - 1 ? "disabled" : ""}>Next &raquo;</button>
              </div>`
            : ""
        }
      </div>
    `;

    wireEvents();
    if (activeId) container.querySelector(`#${activeId}`)?.focus();
  }

  function wireEvents() {
    container.querySelector("#cb-search")?.addEventListener("input", (e) => {
      filterSearchDraft = e.target.value;
    });
    container.querySelector("#cb-search")?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        filterSearch = filterSearchDraft;
        page = 0;
        draw();
      }
    });
    container.querySelector("#cb-search-btn")?.addEventListener("click", () => {
      filterSearch = filterSearchDraft;
      page = 0;
      draw();
    });
    container.querySelector("#cb-search-clear-btn")?.addEventListener("click", () => {
      filterSearch = "";
      filterSearchDraft = "";
      page = 0;
      draw();
    });
    container.querySelector("#cb-provider")?.addEventListener("change", (e) => {
      filterProvider = e.target.value;
      page = 0;
      draw();
    });
    container.querySelector("#cb-prev")?.addEventListener("click", () => {
      page = Math.max(0, page - 1);
      draw();
    });
    container.querySelector("#cb-next")?.addEventListener("click", () => {
      page = page + 1;
      draw();
    });
    container.querySelector("#cb-download-btn")?.addEventListener("click", () => {
      const rowsToExport = filtered();
      if (rowsToExport.length === 0) {
        alert("Nothing to download.");
        return;
      }
      downloadAsExcel(rowsToExport);
    });

    container.querySelectorAll("[data-delete]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.delete;
        const r = rows.find((x) => x.id === id);
        const label = r ? `${r.line?.orders?.customers?.name || r.customer_name_raw || "-"} -- ${fmtMoney(r.commission_amount)}` : id;
        if (!confirm(`Delete this chargeback row (${label})? This only removes the chargeback record itself -- it does not change the customer's original received commission.`)) return;
        btn.disabled = true;
        const { error } = await supabase.from("commission_report_rows").delete().eq("id", id);
        if (error) {
          alert("Failed to delete: " + error.message);
          btn.disabled = false;
          return;
        }
        await load();
        draw();
      });
    });
  }

  await load();
  draw();
}
