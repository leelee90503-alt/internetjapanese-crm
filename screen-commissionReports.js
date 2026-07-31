import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";
import { normalizeAccountNumber, escapeHtml, fmtMoney } from "./normalize.js";

// Admin-only screen: upload a commission report from a provider (Excel/CSV),
// auto-match each row against existing order_service_lines by account number,
// let the admin fix up any unmatched/mismatched rows, then confirm to write
// actual_commission_amount / status back onto the matched order service lines.
//
// Also includes a small manual search/correction tool at the bottom for
// one-off fixes and for providers that don't send commission reports at all
// (has_commission_report = false) — those are simply confirmed by hand.

const FIELD_DEFS = [
  { key: "account_number", label: "Account Number", guesses: ["account", "acc", "계정번호", "계정"] },
  { key: "customer_name", label: "Customer Name (reference only)", guesses: ["name", "customer", "고객"] },
  {
    key: "commission_product_name",
    label: "Commission Product Name (optional — disambiguates multi-line accounts, e.g. multiple mobile lines)",
    guesses: ["commission product name", "commission product", "product name", "package", "plan name", "상품"],
  },
  {
    key: "commission_amount",
    label: "Commission Amount",
    guesses: ["commission", "payout", "payment", "amount", "커미션", "금액"],
  },
];

const AMOUNT_EPSILON = 0.01;

// ---- multi-line-per-account matching helpers ----
// Some providers (e.g. Spectrum) send one commission report row per line
// within an account (Mobile Line 1, Mobile Line 1:Unlimited, Mobile Line 2,
// Mobile Line 2:Unlimited, ...) all sharing a single account number. Plain
// account-number matching can't tell those rows apart, so when the report
// includes a Commission Product Name column we also compare it against each
// candidate order_service_line's plan_name.
function normPlan(s) {
  return String(s ?? "").trim().toLowerCase();
}

function isMobileProductText(s) {
  return /mobile|\bline\s*\d*/.test(normPlan(s));
}

function mobileLineOrdinal(s) {
  const m = normPlan(s).match(/line\s*(\d+)/);
  return m ? Number(m[1]) : 1;
}

// Pick the best unused candidate order_service_line for one report row.
// `candidates` must already be filtered to lines not yet claimed by an
// earlier row in this same upload.
function pickCandidate(candidates, productNameRaw) {
  if (candidates.length === 0) return null;
  if (candidates.length === 1) return candidates[0];

  const product = normPlan(productNameRaw);
  if (!product) {
    // No product name given (or that column wasn't mapped) -- multiple
    // candidates can't be told apart, so fall back to the pre-Part-B
    // best-effort behavior (same as when this account had only one line).
    return candidates.find((l) => l.status === "pending") || candidates[0];
  }

  // 1) exact plan_name match
  let hit = candidates.find((l) => normPlan(l.plan_name) === product);
  if (hit) return hit;

  // 2) substring match either direction (skip bare "mobile" plan_names --
  //    those are ambiguous between multiple lines and handled by ordinal
  //    matching below instead)
  hit = candidates.find((l) => {
    const p = normPlan(l.plan_name);
    if (!p || p === "mobile") return false;
    return p.includes(product) || product.includes(p);
  });
  if (hit) return hit;

  // 3) mobile line ordinal match: "Mobile Line 2" / "Mobile Line 2:Unlimited"
  //    etc. map to the Nth mobile-plan candidate for this account, in
  //    original (created_at) order -- matches how the app numbers mobile
  //    lines during customer upload.
  if (isMobileProductText(product)) {
    const mobileCandidates = candidates.filter((l) => isMobileProductText(l.plan_name));
    const n = mobileLineOrdinal(product);
    if (mobileCandidates[n - 1]) return mobileCandidates[n - 1];
  }

  // Multiple candidates exist and a product name WAS given, but it didn't
  // match any of them (e.g. "Xumo" when no such line was ever created) --
  // don't guess and silently write commission data onto the wrong line;
  // leave this row unmatched so the admin resolves it by hand.
  return null;
}

export async function renderCommissionReports(container, ctx) {
  let step = "upload"; // upload -> mapping -> review
  let providers = [];
  let recentBatches = [];
  let selectedProviderId = "";
  let sourceFilename = "";
  let workbook = null;
  let sheetNames = [];
  let selectedSheet = "";
  let headers = [];
  let dataRows = [];
  let mapping = {};
  let reportRows = [];
  let linesForProvider = [];
  let busyMsg = "";

  // ---- manual search/correction tool state ----
  let manualQuery = "";
  let manualResults = [];
  let manualBusy = "";

  async function loadProviders() {
    const { data } = await supabase.from("providers").select("*").eq("is_active", true).order("name");
    providers = data || [];
  }

  async function loadRecentBatches() {
    const { data } = await supabase
      .from("commission_report_batches")
      .select("id, source_filename, status, created_at, providers(name)")
      .order("created_at", { ascending: false })
      .limit(10);
    recentBatches = data || [];
  }

  function statusBadge(status) {
    if (status === "received") return `<span class="badge ok">Received</span>`;
    if (status === "mismatch") return `<span class="badge error">Mismatch</span>`;
    if (status === "no_report") return `<span class="badge neutral">No Report (manual)</span>`;
    return `<span class="badge warn">Pending</span>`;
  }

  async function drawUploadStep() {
    container.innerHTML = `<div class="screen"><h2>Commission Reports</h2><p class="muted">Loading...</p></div>`;
    await Promise.all([loadProviders(), loadRecentBatches()]);
    container.innerHTML = `
      <div class="screen">
        <h2>Commission Reports</h2>
        <p class="muted">Upload a commission report Excel/CSV from a provider. Rows are automatically matched to existing orders by account number; you review and confirm before anything is saved.</p>
        <div class="card">
          <div class="grid3">
            <label>Provider
              <select id="provider-select">
                <option value="">(Select a provider)</option>
                ${providers
                  .map(
                    (p) =>
                      `<option value="${p.id}">${escapeHtml(p.name)}${p.has_commission_report ? "" : " (no report — manual only)"}</option>`
                  )
                  .join("")}
              </select>
            </label>
            <label>File
              <input type="file" id="file-input" accept=".xlsx,.xls,.csv" />
            </label>
          </div>
        </div>

        <div class="card">
          <h3>Recent Uploads</h3>
          <table class="data-table small">
            <thead><tr><th>Date</th><th>Provider</th><th>File</th><th>Status</th></tr></thead>
            <tbody>
              ${
                recentBatches
                  .map(
                    (b) => `
                <tr>
                  <td>${b.created_at ? new Date(b.created_at).toLocaleString() : "-"}</td>
                  <td>${escapeHtml(b.providers?.name || "-")}</td>
                  <td>${escapeHtml(b.source_filename || "-")}</td>
                  <td>${escapeHtml(b.status)}</td>
                </tr>`
                  )
                  .join("") || `<tr><td colspan="4" class="muted">No uploads yet.</td></tr>`
              }
            </tbody>
          </table>
        </div>

        ${manualToolHtml()}
      </div>
    `;

    container.querySelector("#file-input").addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      selectedProviderId = container.querySelector("#provider-select").value;
      if (!selectedProviderId) {
        alert("Please select a provider first.");
        e.target.value = "";
        return;
      }
      sourceFilename = file.name;
      const buf = await file.arrayBuffer();
      workbook = XLSX.read(buf, { cellDates: true });
      sheetNames = workbook.SheetNames;
      selectedSheet = sheetNames[0];
      loadSheet(selectedSheet);
      step = "mapping";
      drawMappingStep();
    });

    wireManualTool();
  }

  function loadSheet(name) {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    headers = (rows[0] || []).map((h) => String(h ?? "").trim());
    dataRows = rows.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
    mapping = {};
    FIELD_DEFS.forEach((f) => {
      // Exact match first (e.g. a header literally "Commission" or "Account"),
      // then fall back to substring -- exact-first avoids a more specific
      // header (e.g. "Commission Product Name") losing out to an earlier,
      // more generic header that happens to also contain the guess text
      // (e.g. "Commission" the guess matching "Commission Type" first).
      let idx = headers.findIndex((h) => f.guesses.some((g) => h.toLowerCase() === g.toLowerCase()));
      if (idx === -1) {
        idx = headers.findIndex((h) => f.guesses.some((g) => h.toLowerCase().includes(g.toLowerCase())));
      }
      mapping[f.key] = idx;
    });
  }

  function drawMappingStep() {
    const provider = providers.find((p) => p.id === selectedProviderId);
    container.innerHTML = `
      <div class="screen">
        <h2>Step 2: Column Mapping</h2>
        <p class="muted">Provider: <strong>${escapeHtml(provider?.name || "")}</strong> / File: <strong>${escapeHtml(sourceFilename)}</strong></p>
        <div class="card">
          <label>Select Sheet
            <select id="sheet-select">
              ${sheetNames.map((n) => `<option value="${escapeHtml(n)}" ${n === selectedSheet ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">Detected ${headers.length} column(s) / ${dataRows.length} data row(s).</p>
          <table class="data-table">
            <thead><tr><th>System Field</th><th>Excel Column</th></tr></thead>
            <tbody>
              ${FIELD_DEFS.map(
                (f) => `
                <tr>
                  <td>${f.label}</td>
                  <td>
                    <select data-field="${f.key}">
                      <option value="-1" ${mapping[f.key] === -1 || mapping[f.key] === undefined ? "selected" : ""}>(Not used)</option>
                      ${headers
                        .map(
                          (h, i) =>
                            `<option value="${i}" ${mapping[f.key] === i ? "selected" : ""}>${escapeHtml(h) || "(untitled " + i + ")"}</option>`
                        )
                        .join("")}
                    </select>
                  </td>
                </tr>`
              ).join("")}
            </tbody>
          </table>
          <div class="btn-row">
            <button class="btn" id="back-btn">Re-upload</button>
            <button class="btn primary" id="next-btn">Next: Review &amp; Match</button>
          </div>
        </div>
      </div>
    `;

    container.querySelector("#sheet-select").addEventListener("change", (e) => {
      selectedSheet = e.target.value;
      loadSheet(selectedSheet);
      drawMappingStep();
    });
    FIELD_DEFS.forEach((f) => {
      container.querySelector(`[data-field="${f.key}"]`).addEventListener("change", (e) => {
        mapping[f.key] = Number(e.target.value);
      });
    });
    container.querySelector("#back-btn").addEventListener("click", () => {
      step = "upload";
      drawUploadStep();
    });
    container.querySelector("#next-btn").addEventListener("click", async () => {
      await buildRows();
      step = "review";
      drawReviewStep();
    });
  }

  function cell(row, fieldKey) {
    const idx = mapping[fieldKey];
    if (idx === undefined || idx === -1) return "";
    return row[idx] ?? "";
  }

  async function buildRows() {
    const { data: lines, error } = await supabase
      .from("order_service_lines")
      .select(
        "id, account_number, plan_name, expected_commission, status, created_at, orders(order_date, customers(name, phone))"
      )
      .eq("provider_id", selectedProviderId)
      .order("created_at", { ascending: true });
    if (error) {
      alert("Failed to load existing orders for matching: " + error.message);
      linesForProvider = [];
    } else {
      linesForProvider = lines || [];
    }

    const linesByAccount = {};
    linesForProvider.forEach((l) => {
      if (!l.account_number) return;
      (linesByAccount[l.account_number] ||= []).push(l);
    });

    // Parse every data row first, then resolve matches in file order so that,
    // within one account, earlier rows claim their line before later rows
    // (needed for the mobile-line-ordinal fallback in pickCandidate).
    const parsedRows = [];
    for (const row of dataRows) {
      const accountRaw = String(cell(row, "account_number") ?? "").trim();
      const account = normalizeAccountNumber(accountRaw);
      const customerNameRaw = String(cell(row, "customer_name") ?? "").trim();
      const commissionProductNameRaw = String(cell(row, "commission_product_name") ?? "").trim();
      const commRaw = cell(row, "commission_amount");
      const commissionAmount = commRaw === "" ? null : Number(commRaw);
      if (!account && !customerNameRaw && (commissionAmount === null || isNaN(commissionAmount))) continue;
      parsedRows.push({
        accountRaw,
        account,
        customerNameRaw,
        commissionProductNameRaw,
        commissionAmount: isNaN(commissionAmount) ? null : commissionAmount,
      });
    }

    reportRows = [];
    const usedLineIds = new Set();
    for (const pr of parsedRows) {
      const candidates = (pr.account ? linesByAccount[pr.account] || [] : []).filter(
        (l) => !usedLineIds.has(l.id)
      );
      const candidate = pickCandidate(candidates, pr.commissionProductNameRaw);
      if (candidate) usedLineIds.add(candidate.id);

      reportRows.push({
        id: crypto.randomUUID(),
        excluded: false,
        accountNumberRaw: pr.accountRaw,
        accountNumber: pr.account,
        customerNameRaw: pr.customerNameRaw,
        commissionProductNameRaw: pr.commissionProductNameRaw,
        commissionAmount: pr.commissionAmount,
        matchedLineId: candidate ? candidate.id : null,
        matchStatus: candidate ? "matched" : "unmatched",
      });
    }
  }

  function lineLabel(line) {
    const cust = line.orders?.customers;
    const plan = line.plan_name ? ` [${line.plan_name}]` : "";
    return `${cust?.name || "(no name)"}${plan} - ${fmtMoney(line.expected_commission)} (acct: ${line.account_number || "-"})`;
  }

  function drawReviewStep() {
    const usedLineIds = new Set(reportRows.filter((r) => !r.excluded).map((r) => r.matchedLineId).filter(Boolean));
    container.innerHTML = `
      <div class="screen">
        <h2>Step 3: Review &amp; Match</h2>
        <p class="muted">Green = auto-matched and amount agrees. Yellow = mismatched amount (will be saved but flagged). Rows with no match need a manual selection before confirming.</p>
        <div class="btn-row">
          <button class="btn primary" id="confirm-btn">Confirm &amp; Save</button>
          <button class="btn" id="back-to-mapping-btn">Back (Column Mapping)</button>
        </div>
        ${busyMsg ? `<div class="alert info">${escapeHtml(busyMsg)}</div>` : ""}
        <table class="data-table small">
          <thead>
            <tr>
              <th>Include</th><th>Account #</th><th>Customer (from file)</th><th>Product (from file)</th><th>Commission Amount</th><th>Match</th>
            </tr>
          </thead>
          <tbody>
            ${reportRows
              .map((r) => {
                const line = linesForProvider.find((l) => l.id === r.matchedLineId);
                const mismatch =
                  line &&
                  line.expected_commission != null &&
                  r.commissionAmount != null &&
                  Math.abs(Number(line.expected_commission) - Number(r.commissionAmount)) > AMOUNT_EPSILON;
                return `
                <tr data-row="${r.id}" class="${r.excluded ? "excluded" : ""}">
                  <td><input type="checkbox" class="r-include" ${r.excluded ? "" : "checked"} /></td>
                  <td><input type="text" class="r-account" value="${escapeHtml(r.accountNumberRaw)}" /></td>
                  <td>${escapeHtml(r.customerNameRaw) || "-"}</td>
                  <td>${escapeHtml(r.commissionProductNameRaw) || "-"}</td>
                  <td><input type="number" step="0.01" class="r-amount" value="${r.commissionAmount ?? ""}" /></td>
                  <td>
                    ${
                      line
                        ? `<span class="badge ${mismatch ? "error" : "ok"}">${mismatch ? "Mismatch" : "Matched"}</span>
                           <div class="muted">${escapeHtml(lineLabel(line))}</div>`
                        : `<span class="badge warn">No match</span>`
                    }
                    <select class="r-manual-match">
                      <option value="">(No selection)</option>
                      ${linesForProvider
                        .filter((l) => l.id === r.matchedLineId || !usedLineIds.has(l.id))
                        .map(
                          (l) =>
                            `<option value="${l.id}" ${l.id === r.matchedLineId ? "selected" : ""}>${escapeHtml(lineLabel(l))}</option>`
                        )
                        .join("")}
                    </select>
                  </td>
                </tr>`;
              })
              .join("") || `<tr><td colspan="6" class="muted">No rows were loaded from the file.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
    wireReviewEvents();
  }

  function readRowFromTr(tr, r) {
    r.excluded = !tr.querySelector(".r-include").checked;
    r.accountNumberRaw = tr.querySelector(".r-account").value.trim();
    const amtVal = tr.querySelector(".r-amount").value;
    r.commissionAmount = amtVal === "" ? null : Number(amtVal);
    r.matchedLineId = tr.querySelector(".r-manual-match").value || null;
    r.matchStatus = r.matchedLineId ? "manual" : "unmatched";
  }

  function wireReviewEvents() {
    container.querySelector("#back-to-mapping-btn").addEventListener("click", () => {
      step = "mapping";
      drawMappingStep();
    });
    container.querySelectorAll("tr[data-row]").forEach((tr) => {
      const r = reportRows.find((x) => x.id === tr.dataset.row);
      tr.querySelector(".r-manual-match").addEventListener("change", () => {
        readRowFromTr(tr, r);
        drawReviewStep();
      });
    });
    container.querySelector("#confirm-btn").addEventListener("click", async () => {
      container.querySelectorAll("tr[data-row]").forEach((tr) => {
        const r = reportRows.find((x) => x.id === tr.dataset.row);
        readRowFromTr(tr, r);
      });
      await confirmRows();
    });
  }

  async function confirmRows() {
    const toConfirm = reportRows.filter((r) => !r.excluded);
    if (toConfirm.length === 0) {
      alert("No rows are included.");
      return;
    }
    const unresolved = toConfirm.filter((r) => !r.matchedLineId);
    if (unresolved.length > 0) {
      alert(`${unresolved.length} row(s) have no matched order. Select a match or uncheck "Include" for those rows.`);
      return;
    }

    busyMsg = "Saving...";
    drawReviewStep();

    const { data: batch, error: batchErr } = await supabase
      .from("commission_report_batches")
      .insert({
        provider_id: selectedProviderId,
        source_filename: sourceFilename,
        uploaded_by: ctx.profile.id,
        status: "confirmed",
      })
      .select()
      .single();
    if (batchErr) {
      busyMsg = "Failed to save batch: " + batchErr.message;
      drawReviewStep();
      return;
    }

    let successCount = 0;
    const errors = [];
    for (const r of toConfirm) {
      try {
        const line = linesForProvider.find((l) => l.id === r.matchedLineId);
        const mismatch =
          line &&
          line.expected_commission != null &&
          r.commissionAmount != null &&
          Math.abs(Number(line.expected_commission) - Number(r.commissionAmount)) > AMOUNT_EPSILON;

        const { error: updErr } = await supabase
          .from("order_service_lines")
          .update({
            actual_commission_amount: r.commissionAmount,
            commission_matched_at: new Date().toISOString(),
            commission_matched_by: ctx.profile.id,
            status: mismatch ? "mismatch" : "received",
          })
          .eq("id", r.matchedLineId);
        if (updErr) throw updErr;

        const { error: rowErr } = await supabase.from("commission_report_rows").insert({
          batch_id: batch.id,
          raw_data: {
            account_number: r.accountNumberRaw,
            customer_name: r.customerNameRaw,
            commission_product_name: r.commissionProductNameRaw,
            commission_amount: r.commissionAmount,
          },
          account_number: r.accountNumber,
          customer_name_raw: r.customerNameRaw,
          commission_amount: r.commissionAmount,
          matched_order_service_line_id: r.matchedLineId,
          match_status: r.matchStatus,
          confirmed: true,
        });
        if (rowErr) throw rowErr;

        successCount++;
        r.savedOk = true;
      } catch (err) {
        errors.push(`${r.accountNumberRaw || "(no account #)"}: ${err.message || err}`);
      }
    }

    reportRows = reportRows.filter((r) => !r.savedOk);
    busyMsg = `${successCount} saved.` + (errors.length ? ` ${errors.length} failed: ` + errors.join(" / ") : "");
    if (reportRows.length === 0) {
      step = "upload";
      await drawUploadStep();
      return;
    }
    drawReviewStep();
  }

  // ---- manual search / correction tool ----
  function manualToolHtml() {
    return `
      <div class="card">
        <h3>Manual Search / Correction</h3>
        <p class="muted">Look up an order by account number to manually set its commission status — useful for providers with no commission report, or for fixing a single row.</p>
        <div class="inline-form">
          <input type="text" id="manual-search-input" placeholder="Search by account number" value="${escapeHtml(manualQuery)}" />
          <button class="btn" id="manual-search-btn">Search</button>
        </div>
        ${manualBusy ? `<div class="alert info">${escapeHtml(manualBusy)}</div>` : ""}
        ${
          manualResults.length > 0
            ? `<table class="data-table small">
                <thead><tr><th>Customer</th><th>Account #</th><th>Provider</th><th>Service</th><th>Expected</th><th>Actual</th><th>Status</th><th></th></tr></thead>
                <tbody>
                  ${manualResults
                    .map(
                      (l) => `
                    <tr data-line="${l.id}">
                      <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
                      <td>${escapeHtml(l.account_number || "-")}</td>
                      <td>${escapeHtml(l.providers?.name || "-")}</td>
                      <td>${escapeHtml(l.services?.name || "-")}</td>
                      <td>${fmtMoney(l.expected_commission)}</td>
                      <td><input type="number" step="0.01" class="m-amount" value="${l.actual_commission_amount ?? ""}" style="width:90px" /></td>
                      <td>
                        <select class="m-status">
                          <option value="pending" ${l.status === "pending" ? "selected" : ""}>Pending</option>
                          <option value="received" ${l.status === "received" ? "selected" : ""}>Received</option>
                          <option value="mismatch" ${l.status === "mismatch" ? "selected" : ""}>Mismatch</option>
                          <option value="no_report" ${l.status === "no_report" ? "selected" : ""}>No Report (manual)</option>
                        </select>
                      </td>
                      <td><button class="btn small m-save">Save</button></td>
                    </tr>`
                    )
                    .join("")}
                </tbody>
              </table>`
            : ""
        }
      </div>
    `;
  }

  function wireManualTool() {
    container.querySelector("#manual-search-btn").addEventListener("click", async () => {
      manualQuery = container.querySelector("#manual-search-input").value;
      await runManualSearch();
    });
    container.querySelector("#manual-search-input").addEventListener("keydown", async (e) => {
      if (e.key === "Enter") {
        manualQuery = e.target.value;
        await runManualSearch();
      }
    });
    wireManualResultRows();
  }

  function wireManualResultRows() {
    container.querySelectorAll("tr[data-line] .m-save")?.forEach((btn) => {
      btn.addEventListener("click", async () => {
        const tr = btn.closest("tr");
        const id = tr.dataset.line;
        const amtVal = tr.querySelector(".m-amount").value;
        const status = tr.querySelector(".m-status").value;
        const { error } = await supabase
          .from("order_service_lines")
          .update({
            actual_commission_amount: amtVal === "" ? null : Number(amtVal),
            status,
            commission_matched_at: new Date().toISOString(),
            commission_matched_by: ctx.profile.id,
          })
          .eq("id", id);
        if (error) {
          alert("Failed to save: " + error.message);
          return;
        }
        manualBusy = "Saved.";
        await runManualSearch();
      });
    });
  }

  async function runManualSearch() {
    const q = manualQuery.trim();
    if (!q) {
      manualResults = [];
      redrawManualToolOnly();
      return;
    }
    manualBusy = "Searching...";
    redrawManualToolOnly();
    const { data, error } = await supabase
      .from("order_service_lines")
      .select(
        "id, account_number, expected_commission, actual_commission_amount, status, services(name), providers(name), orders(customers(name, phone))"
      )
      .ilike("account_number", `%${q}%`)
      .order("created_at", { ascending: false })
      .limit(15);
    if (error) {
      manualBusy = "Search failed: " + error.message;
      manualResults = [];
      redrawManualToolOnly();
      return;
    }
    manualResults = data || [];
    manualBusy = manualResults.length === 0 ? "No matching orders found." : "";
    redrawManualToolOnly();
  }

  function redrawManualToolOnly() {
    // Only the last card (the manual tool) needs to be replaced; the manual
    // tool only appears on the upload step.
    if (step === "upload") {
      const cards = container.querySelectorAll(".card");
      const last = cards[cards.length - 1];
      if (last) last.outerHTML = manualToolHtml();
      wireManualTool();
    }
  }

  await drawUploadStep();
}
