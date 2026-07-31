import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import { escapeHtml, fmtMoney, normalizeAccountNumber } from "./normalize.js";

// Pending Commission: shows ALL order_service_lines currently in
// status='pending', with a full detail view per line. Staff can expand any
// line to see the expected commission + ordered service, a live lookup of
// the uploaded commission report for that account number (so they can see
// whether this customer already showed up in a report and, if so, why it
// didn't auto-match), a free-text note, and a 3-way decision:
//   - Keep Pending    -> no status change, just save the note
//   - Mark Received   -> status='received', copies expected_commission into
//                        actual_commission_amount (same effect as a normal match)
//   - Confirmed Missing -> status='missing' -- this is the ONLY way a line
//                        moves to the Missing Commission screen. There is no
//                        automatic 21-day promotion anymore; lines just get
//                        an "21+ days" warning badge here so staff can see it
//                        and decide for themselves.
//
// This replaces the old "Flag as Missing" button, which used to create a
// missing_commission_items row -- that table is now legacy-only.

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
  let busy = false;
  let page = 0;

  // filters
  let filterProvider = "";
  let filterSearch = "";
  let filterMinDays = "";
  let sortBy = "order_date_asc"; // oldest first by default -- oldest pending is usually most urgent

  // detail-expand state
  let expandedId = null;
  let noteDraft = ""; // free-text note being edited for the expanded line

  // commission report lookup cache, keyed by normalized account number.
  // Values: { status: 'loading' | 'done' | 'error', rows: [...], error?: string }
  const reportLookupCache = new Map();

  async function load() {
    const { data: providerRows } = await supabase.from("providers").select("id, name").order("name");
    providers = providerRows || [];

    const { data, error } = await supabase
      .from("order_service_lines")
      .select(
        "id, account_number, plan_name, expected_commission, order_id, providers(name), orders(order_date, order_number, customers(name))"
      )
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
  }

  async function lookupReport(accountNumber) {
    const key = normalizeAccountNumber(accountNumber);
    if (!key) return { status: "done", rows: [] };
    if (reportLookupCache.has(key)) return reportLookupCache.get(key);

    reportLookupCache.set(key, { status: "loading", rows: [] });
    const { data, error } = await supabase
      .from("commission_report_rows")
      .select("id, account_number, customer_name_raw, commission_amount, raw_data, match_status, matched_order_service_line_id, created_at")
      .eq("account_number", accountNumber)
      .order("created_at", { ascending: false })
      .limit(10);

    const result = error ? { status: "error", rows: [], error: error.message } : { status: "done", rows: data || [] };
    reportLookupCache.set(key, result);
    return result;
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

  function reportRowSummary(r) {
    const prodName = r.raw_data?.commission_product_name || "";
    const parts = [];
    if (r.customer_name_raw) parts.push(escapeHtml(r.customer_name_raw));
    if (prodName) parts.push(escapeHtml(prodName));
    parts.push(fmtMoney(r.commission_amount));
    return parts.join(" &middot; ");
  }

  function explainMismatch(line, reportRow) {
    // Best-effort explanation of why this report row didn't auto-match this
    // line, mirroring the comparison the confirm-matches tool uses.
    if (reportRow.matched_order_service_line_id === line.id) {
      return `This report row already links to this exact order line (match: ${escapeHtml(reportRow.match_status || "-")}).`;
    }
    if (reportRow.matched_order_service_line_id) {
      return `This report row matched a different order line instead -- likely because another pending line looked like a closer match at the time (e.g. same account, similar product name/ordinal).`;
    }
    const reasons = [];
    const expected = line.expected_commission === null || line.expected_commission === undefined ? null : Number(line.expected_commission);
    const actual = reportRow.commission_amount === null || reportRow.commission_amount === undefined ? null : Number(reportRow.commission_amount);
    if (expected !== null && actual !== null && Math.abs(expected - actual) > 0.01) {
      reasons.push(`amount differs (expected ${fmtMoney(expected)} vs. report ${fmtMoney(actual)})`);
    }
    const prodName = (reportRow.raw_data?.commission_product_name || "").toLowerCase();
    const planName = (line.plan_name || "").toLowerCase();
    if (prodName && planName && !prodName.includes(planName) && !planName.includes(prodName)) {
      reasons.push(`product name differs (expected "${escapeHtml(line.plan_name || "")}" vs. report "${escapeHtml(reportRow.raw_data?.commission_product_name || "")}")`);
    }
    if (reasons.length === 0) {
      return "Unmatched -- no obvious reason found; may need manual review in the Commission Reports screen.";
    }
    return "Didn't auto-match because " + reasons.join(" and ") + ".";
  }

  function reportLookupHtml(line) {
    const key = normalizeAccountNumber(line.account_number);
    if (!key) {
      return `<p class="muted">No account number on this line -- can't look up the commission report.</p>`;
    }
    const cached = reportLookupCache.get(key);
    if (!cached || cached.status === "loading") {
      return `<p class="muted">Looking up commission report for account ${escapeHtml(line.account_number)}...</p>`;
    }
    if (cached.status === "error") {
      return `<p class="alert error">Failed to look up commission report: ${escapeHtml(cached.error)}</p>`;
    }
    if (cached.rows.length === 0) {
      return `<p class="muted">No commission report rows found for account ${escapeHtml(line.account_number)} -- this customer hasn't appeared in any uploaded report yet.</p>`;
    }
    return `
      <p class="muted">Found ${cached.rows.length} report row(s) for account ${escapeHtml(line.account_number)}:</p>
      <table class="data-table">
        <thead><tr><th>Reported As</th><th>Match Status</th><th>Why it didn't match this line</th></tr></thead>
        <tbody>
          ${cached.rows
            .map(
              (r) => `
            <tr>
              <td>${reportRowSummary(r)}</td>
              <td>${escapeHtml(r.match_status || "-")}</td>
              <td>${explainMismatch(line, r)}</td>
            </tr>`
            )
            .join("")}
        </tbody>
      </table>
    `;
  }

  function detailHtml(line) {
    const days = daysPending(line.orders?.order_date);
    return `
      <tr class="detail-row">
        <td colspan="8">
          <div class="detail-panel">
            <div class="detail-grid">
              <div>
                <strong>Ordered service:</strong> ${escapeHtml(line.plan_name || "-")}<br/>
                <strong>Provider:</strong> ${escapeHtml(line.providers?.name || "-")}<br/>
                <strong>Order #:</strong> ${escapeHtml(line.orders?.order_number || "-")}<br/>
                <strong>Order date:</strong> ${escapeHtml(line.orders?.order_date || "-")}
                ${days !== null ? ` (${days} days pending)` : ""}
              </div>
              <div>
                <strong>Expected commission:</strong> ${fmtMoney(line.expected_commission)}<br/>
                <strong>Account #:</strong> ${escapeHtml(line.account_number || "-")}<br/>
                <strong>Customer:</strong> ${escapeHtml(line.orders?.customers?.name || "-")}
              </div>
            </div>

            <h4>Commission report lookup</h4>
            ${reportLookupHtml(line)}

            <h4>Decision</h4>
            <label class="field-label" for="pc-note-${line.id}">Note</label>
            <textarea id="pc-note-${line.id}" data-note-for="${line.id}" rows="2" style="width:100%" placeholder="Optional note -- e.g. why you're keeping this pending, or what the provider said">${escapeHtml(noteDraft)}</textarea>
            <div class="inline-form" style="margin-top:8px">
              <button class="btn small" data-decide="${line.id}:keep_pending" ${busy ? "disabled" : ""}>Keep Pending</button>
              <button class="btn small primary" data-decide="${line.id}:received" ${busy ? "disabled" : ""}>Mark Received</button>
              <button class="btn small danger" data-decide="${line.id}:missing" ${busy ? "disabled" : ""}>Confirmed Missing</button>
            </div>
          </div>
        </td>
      </tr>
    `;
  }

  function rowHtml(l) {
    const days = daysPending(l.orders?.order_date);
    const isExpanded = expandedId === l.id;
    return `
      <tr class="clickable-row ${isExpanded ? "expanded" : ""}" data-toggle="${l.id}">
        <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
        <td>${escapeHtml(l.account_number || "-")}</td>
        <td>${escapeHtml(l.providers?.name || "-")}</td>
        <td>${escapeHtml(l.plan_name || "-")}</td>
        <td>${escapeHtml(l.orders?.order_date || "-")}</td>
        <td>${ageBadge(days)}</td>
        <td>${fmtMoney(l.expected_commission)}</td>
        <td><button class="btn small" data-toggle-btn="${l.id}">${isExpanded ? "Hide" : "Details"}</button></td>
      </tr>
      ${isExpanded ? detailHtml(l) : ""}
    `;
  }

  function draw() {
    // Preserve focus/cursor position across the innerHTML replacement below --
    // otherwise every keystroke in a text input re-renders the DOM and drops
    // focus, so only the first character of anything typed quickly ever lands.
    const active = document.activeElement;
    const activeId = active && container.contains(active) ? active.id : null;
    const activeSelStart = activeId && "selectionStart" in active ? active.selectionStart : null;
    const activeSelEnd = activeId && "selectionEnd" in active ? active.selectionEnd : null;

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
          All order lines currently waiting on a commission report. Click a row for full detail -- expected commission,
          ordered service, and a live lookup of any commission report already uploaded for that account -- then choose
          Keep Pending, Mark Received, or Confirmed Missing. Only "Confirmed Missing" sends a line to the Missing
          Commission follow-up list; the "21+ days" badge below is just a visual warning, nothing moves automatically.
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

    if (activeId) {
      const toRefocus = container.querySelector(`#${activeId}`);
      if (toRefocus) {
        toRefocus.focus();
        if (activeSelStart !== null && "setSelectionRange" in toRefocus) {
          try {
            toRefocus.setSelectionRange(activeSelStart, activeSelEnd);
          } catch {
            // some input types (e.g. number) don't support setSelectionRange -- ignore
          }
        }
      }
    }

    // If a line is expanded and its report lookup hasn't been kicked off yet,
    // fire it now and redraw when it resolves. This runs after wireEvents so
    // the click handlers above are already attached.
    if (expandedId) {
      const line = lines.find((l) => l.id === expandedId);
      const key = line ? normalizeAccountNumber(line.account_number) : "";
      if (line && key && !reportLookupCache.has(key)) {
        lookupReport(line.account_number).then(() => draw());
      }
    }
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
      const rows = filteredSorted();
      if (rows.length === 0) {
        alert("Nothing to download.");
        return;
      }
      downloadAsExcel(rows);
    });

    container.querySelectorAll("[data-toggle], [data-toggle-btn]").forEach((el) => {
      el.addEventListener("click", (e) => {
        // avoid double-firing when the button inside the row is clicked
        if (el.hasAttribute("data-toggle") && e.target.closest("[data-toggle-btn]")) return;
        const id = el.dataset.toggle || el.dataset.toggleBtn;
        expandedId = expandedId === id ? null : id;
        noteDraft = "";
        draw();
      });
    });

    container.querySelectorAll("[data-note-for]").forEach((el) => {
      el.addEventListener("input", (e) => {
        noteDraft = e.target.value;
      });
    });

    container.querySelectorAll("[data-decide]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const [id, choice] = btn.dataset.decide.split(":");
        const line = lines.find((l) => l.id === id);
        if (!line) return;

        const note = noteDraft.trim() || null;

        if (choice === "keep_pending") {
          busy = true;
          draw();
          const { error } = await supabase
            .from("order_service_lines")
            .update({ resolution_note: note, resolution_at: note ? new Date().toISOString() : null, resolution_by: note ? ctx.profile?.id || null : null })
            .eq("id", id);
          busy = false;
          if (error) alert("Failed to save note: " + error.message);
          await load();
          draw();
          return;
        }

        if (choice === "received") {
          if (!confirm(`Mark "${line.orders?.customers?.name || "this line"}" - ${line.plan_name || ""} as received?`)) return;
          busy = true;
          draw();
          const { error } = await supabase
            .from("order_service_lines")
            .update({
              status: "received",
              actual_commission_amount: line.expected_commission,
              commission_matched_at: new Date().toISOString(),
              commission_matched_by: ctx.profile?.id || null,
              resolution_note: note,
              resolution_at: new Date().toISOString(),
              resolution_by: ctx.profile?.id || null,
            })
            .eq("id", id);
          busy = false;
          if (error) alert("Failed to mark received: " + error.message);
          expandedId = null;
          await load();
          draw();
          return;
        }

        if (choice === "missing") {
          if (
            !confirm(
              `Confirm "${line.orders?.customers?.name || "this line"}" - ${line.plan_name || ""} as genuinely missing commission? This moves it to the Missing Commission follow-up list.`
            )
          )
            return;
          busy = true;
          draw();
          const { error } = await supabase
            .from("order_service_lines")
            .update({
              status: "missing",
              resolution_note: note,
              resolution_at: new Date().toISOString(),
              resolution_by: ctx.profile?.id || null,
            })
            .eq("id", id);
          busy = false;
          if (error) alert("Failed to confirm missing: " + error.message);
          expandedId = null;
          await load();
          draw();
          return;
        }
      });
    });
  }

  await load();
  draw();
}
