import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import { escapeHtml, fmtMoney, normalizeAccountNumber } from "./normalize.js";

// Pending Commission: a straightforward, searchable/filterable listing of
// every order_service_lines row currently sitting in status='pending'.
// This is a read/browse screen (no workflow state of its own) -- it exists
// so staff can see the full pending picture at a glance, separate from
// Missing Commission (which is the curated "we believe this is actually
// missing" follow-up list, populated manually or after STALE_DAYS).
//
// From here staff can jump a line straight into Missing Commission via
// "Flag as Missing" (creates a missing_commission_items row linked to that
// line, same shape as a manual add there).

const PAGE_SIZE = 50;

function daysPending(orderDate) {
  if (!orderDate) return null;
  const start = new Date(orderDate + "T00:00:00");
  const today = new Date();
  return Math.floor((today - start) / 86400000);
}

function ageBadge(days) {
  if (days === null) return `<span class="badge neutral">-</span>`;
  if (days >= 21) return `<span class="badge error">${days}d</span>`;
  if (days >= 14) return `<span class="badge warn">${days}d</span>`;
  return `<span class="badge neutral">${days}d</span>`;
}

function downloadAsExcel(rows) {
  const out = rows.map((l) => ({
    Customer: l.orders?.customers?.name || "",
    "Account #": l.account_number || "",
    Provider: l.providers?.name || "",
    Plan: l.plan_name || "",
    "Order Date": l.orders?.order_date || "",
    "Days Pending": daysPending(l.orders?.order_date) ?? "",
    "Order #": l.orders?.order_number || "",
    "Expected Commission": l.expected_commission === null || l.expected_commission === undefined ? "" : Number(l.expected_commission),
    "Already Flagged Missing": l.__flagged ? "Yes" : "No",
  }));
  const sheet = XLSX.utils.json_to_sheet(out);
  sheet["!cols"] = [
    { wch: 20 }, // Customer
    { wch: 18 }, // Account #
    { wch: 12 }, // Provider
    { wch: 20 }, // Plan
    { wch: 12 }, // Order Date
    { wch: 12 }, // Days Pending
    { wch: 18 }, // Order #
    { wch: 16 }, // Expected Commission
    { wch: 16 }, // Already Flagged Missing
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Pending Commission");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Pending_Commission_${dateStr}.xlsx`);
}

export async function renderPendingCommission(container, ctx) {
  let lines = [];
  let loadError = null;
  let providers = [];
  let flaggedLineIds = new Set(); // line ids already linked to a missing_commission_items row
  let busy = false;
  let page = 0;

  // filters
  let filterProvider = "";
  let filterSearch = "";
  let filterMinDays = "";
  let sortBy = "order_date_asc"; // oldest first by default -- oldest pending is usually most urgent

  async function load() {
    const { data: providerRows } = await supabase.from("providers").select("id, name").order("name");
    providers = providerRows || [];

    const { data, error } = await supabase
      .from("order_service_lines")
      .select("id, account_number, plan_name, expected_commission, order_id, providers(name), orders(order_date, order_number, customers(name))")
      .eq("status", "pending")
      .order("order_id", { ascending: true })
      .limit(5000);
    if (error) {
      loadError = error.message;
      lines = [];
      return;
    }
    loadError = null;
    lines = data || [];

    const { data: missingItems } = await supabase.from("missing_commission_items").select("matched_line_ids");
    flaggedLineIds = new Set((missingItems || []).flatMap((m) => m.matched_line_ids || []));
  }

  function filteredSorted() {
    let rows = lines.slice();

    if (filterProvider) {
      rows = rows.filter((l) => l.providers?.name === filterProvider);
    }
    if (filterSearch.trim()) {
      const needle = filterSearch.trim().toLowerCase();
      const needleAcct = normalizeAccountNumber(filterSearch);
      rows = rows.filter((l) => {
        const name = (l.orders?.customers?.name || "").toLowerCase();
        const acct = (l.account_number || "").toString();
        const plan = (l.plan_name || "").toLowerCase();
        return (
          name.includes(needle) ||
          plan.includes(needle) ||
          (needleAcct && acct.replace(/\D/g, "").includes(needleAcct))
        );
      });
    }
    if (filterMinDays !== "" && !isNaN(Number(filterMinDays))) {
      const min = Number(filterMinDays);
      rows = rows.filter((l) => (daysPending(l.orders?.order_date) ?? -1) >= min);
    }

    rows.sort((a, b) => {
      if (sortBy === "order_date_asc") {
        return (a.orders?.order_date || "") < (b.orders?.order_date || "") ? -1 : 1;
      }
      if (sortBy === "order_date_desc") {
        return (a.orders?.order_date || "") > (b.orders?.order_date || "") ? -1 : 1;
      }
      if (sortBy === "expected_desc") {
        return (Number(b.expected_commission) || 0) - (Number(a.expected_commission) || 0);
      }
      if (sortBy === "customer_asc") {
        return (a.orders?.customers?.name || "").localeCompare(b.orders?.customers?.name || "");
      }
      return 0;
    });

    return rows;
  }

  function rowHtml(l) {
    const days = daysPending(l.orders?.order_date);
    const flagged = flaggedLineIds.has(l.id);
    return `
      <tr>
        <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
        <td>${escapeHtml(l.account_number || "-")}</td>
        <td>${escapeHtml(l.providers?.name || "-")}</td>
        <td>${escapeHtml(l.plan_name || "-")}</td>
        <td>${escapeHtml(l.orders?.order_date || "-")}</td>
        <td>${ageBadge(days)}</td>
        <td>${fmtMoney(l.expected_commission)}</td>
        <td>
          ${
            flagged
              ? `<span class="badge warn">Flagged</span>`
              : `<button class="btn small" data-flag="${l.id}" ${busy ? "disabled" : ""}>Flag as Missing</button>`
          }
        </td>
      </tr>
    `;
  }

  function draw() {
    const rows = filteredSorted();
    const totalPages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    if (page >= totalPages) page = totalPages - 1;
    const pageRows = rows.slice(page * PAGE_SIZE, (page + 1) * PAGE_SIZE);

    const totalExpected = rows.reduce((s, l) => s + (Number(l.expected_commission) || 0), 0);
    const staleCount = rows.filter((l) => (daysPending(l.orders?.order_date) ?? -1) >= 21).length;

    container.innerHTML = `
      <div class="screen">
        <h2>Pending Commission</h2>
        <p class="muted">
          All order lines currently sitting in "pending" status (waiting on the provider's commission report to confirm).
          This is a browse/search view -- for items you believe are actually missing (not just still waiting), use
          "Flag as Missing" to send them to the Missing Commission follow-up list.
        </p>

        ${loadError ? `<div class="alert error">Failed to load: ${escapeHtml(loadError)}</div>` : ""}

        <div class="inline-form">
          <span class="badge neutral">${rows.length} pending${filterProvider || filterSearch || filterMinDays !== "" ? ` (filtered from ${lines.length})` : ""}</span>
          <span class="badge error">${staleCount} over 21 days</span>
          <span class="badge neutral">Total expected: ${fmtMoney(totalExpected)}</span>
          <button class="btn small" id="pc-download-btn">Download Excel</button>
        </div>

        <div class="inline-form">
          <input type="text" placeholder="Search customer / account / plan" id="pc-search" value="${escapeHtml(filterSearch)}" style="flex:1" />
          <select id="pc-provider">
            <option value="">All providers</option>
            ${providers.map((p) => `<option value="${escapeHtml(p.name)}" ${filterProvider === p.name ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
          </select>
          <input type="number" placeholder="Min days pending" id="pc-min-days" value="${escapeHtml(filterMinDays)}" style="width:150px" />
          <select id="pc-sort">
            <option value="order_date_asc" ${sortBy === "order_date_asc" ? "selected" : ""}>Oldest order first</option>
            <option value="order_date_desc" ${sortBy === "order_date_desc" ? "selected" : ""}>Newest order first</option>
            <option value="expected_desc" ${sortBy === "expected_desc" ? "selected" : ""}>Highest expected $ first</option>
            <option value="customer_asc" ${sortBy === "customer_asc" ? "selected" : ""}>Customer name (A-Z)</option>
          </select>
        </div>

        <table class="data-table">
          <thead>
            <tr>
              <th>Customer</th><th>Account #</th><th>Provider</th><th>Plan</th>
              <th>Order Date</th><th>Days Pending</th><th>Expected</th><th></th>
            </tr>
          </thead>
          <tbody>
            ${pageRows.map(rowHtml).join("") || `<tr><td colspan="8" class="muted">Nothing pending${filterProvider || filterSearch || filterMinDays !== "" ? " matching these filters" : ""}.</td></tr>`}
          </tbody>
        </table>

        ${
          totalPages > 1
            ? `<div class="inline-form">
                <button class="btn small" id="pc-prev" ${page === 0 ? "disabled" : ""}>&laquo; Prev</button>
                <span class="muted">Page ${page + 1} of ${totalPages}</span>
                <button class="btn small" id="pc-next" ${page >= totalPages - 1 ? "disabled" : ""}>Next &raquo;</button>
              </div>`
            : ""
        }
      </div>
    `;

    wireEvents();
  }

  function wireEvents() {
    container.querySelector("#pc-search")?.addEventListener("input", (e) => {
      filterSearch = e.target.value;
      page = 0;
      draw();
    });
    container.querySelector("#pc-provider")?.addEventListener("change", (e) => {
      filterProvider = e.target.value;
      page = 0;
      draw();
    });
    container.querySelector("#pc-min-days")?.addEventListener("input", (e) => {
      filterMinDays = e.target.value;
      page = 0;
      draw();
    });
    container.querySelector("#pc-sort")?.addEventListener("change", (e) => {
      sortBy = e.target.value;
      draw();
    });
    container.querySelector("#pc-prev")?.addEventListener("click", () => {
      page = Math.max(0, page - 1);
      draw();
    });
    container.querySelector("#pc-next")?.addEventListener("click", () => {
      page = page + 1;
      draw();
    });
    container.querySelector("#pc-download-btn")?.addEventListener("click", () => {
      const rows = filteredSorted().map((l) => ({ ...l, __flagged: flaggedLineIds.has(l.id) }));
      if (rows.length === 0) {
        alert("Nothing to download.");
        return;
      }
      downloadAsExcel(rows);
    });

    container.querySelectorAll("[data-flag]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.flag;
        const line = lines.find((l) => l.id === id);
        if (!line) return;
        if (!confirm(`Flag "${line.orders?.customers?.name || "this line"}" - ${line.plan_name || ""} as Missing Commission?`)) return;
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").insert({
          customer_name: line.orders?.customers?.name || "(unknown)",
          account_number: line.account_number || null,
          description: line.plan_name || "(no plan name)",
          source_order_number: line.orders?.order_number || null,
          sales_date: line.orders?.order_date || null,
          price: line.expected_commission,
          claimed_date_raw: null,
          status_notes: "Manually flagged from Pending Commission list.",
          matched_line_ids: [line.id],
          needs_review: true,
          review_note: "Flagged from Pending Commission -- confirm before marking received.",
          auto_detected: false,
          created_by: ctx.profile?.id || null,
        });
        busy = false;
        if (error) {
          alert("Failed to flag: " + error.message);
        } else {
          flaggedLineIds.add(id);
        }
        draw();
      });
    });
  }

  await load();
  draw();
}
