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

// A negative Commission Amount in a provider's report means a chargeback --
// the provider is clawing back commission already paid (e.g. the customer
// canceled early), not reporting a fresh amount to compare against
// expected_commission. These rows must never overwrite the matched line's
// actual_commission_amount/status (that would erase the record of the
// original commission having been received) -- they're recorded as their
// own commission_report_rows entry (matched_order_service_line_id still
// set, so the line/customer/product is known) and surfaced on the
// Chargebacks report instead.
function isChargebackRow(r) {
  return r.commissionAmount != null && Number(r.commissionAmount) < 0;
}

// Detects a TV-plan rename: the matched line's stored plan_name and the
// report's product name are both TV plans but spelled differently (e.g.
// "TV Stream" on file, "TV Essentials" on the report). Returns the OLD name
// to display/save if this looks like a rename, otherwise null. Shared by
// buildRows() (auto-match via pickCandidate's rename tier) and
// readRowFromTr() (a manual match the admin picked by hand) so both paths
// get the same "update plan_name to the new name" treatment.
function renamedFromIfTvRename(line, productNameRaw) {
  if (!line || !line.plan_name) return null;
  const product = normPlan(productNameRaw);
  if (!product || !isTvServiceText(product) || !isTvServiceText(line.plan_name)) return null;
  if (normPlan(line.plan_name) === product) return null;
  return line.plan_name;
}

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

// Matches "TV Choice" / "TV Stream" / "TV Essentials" / etc. -- any plan
// name with a standalone "TV" word in it. Used both for the rename fallback
// in pickCandidate() and for isChargebackRow's sibling renamedFromIfTvRename()
// above (defined earlier in this file, before this function -- function
// declarations are hoisted, so the forward reference is fine).
function isTvServiceText(s) {
  return /\btv\b/i.test(String(s ?? ""));
}

function mobileLineOrdinal(s) {
  const m = normPlan(s).match(/line\s*(\d+)/);
  return m ? Number(m[1]) : 1;
}

// Coarse service category for a plan name / reported product name, used only
// to sanity-check the "account has exactly one remaining line" shortcut
// below. Returns null when the text doesn't clearly indicate any category
// (e.g. blank, or a generic legacy plan_name like plain "Internet"/"Mobile"
// that doesn't conflict with anything) -- null means "can't rule it out",
// not "definitely matches".
function productCategory(s) {
  const p = normPlan(s);
  if (!p) return null;
  if (isTvServiceText(s)) return "tv";
  if (/xumo/.test(p)) return "xumo";
  if (isMobileProductText(s)) return "mobile";
  if (/internet|gig|advantage|premier/.test(p)) return "internet";
  return null;
}

// Pick the best unused candidate order_service_line for one report row.
// `candidates` must already be filtered to lines not yet claimed by an
// earlier row in this same upload.
function pickCandidate(candidates, productNameRaw) {
  if (candidates.length === 0) return null;

  const product = normPlan(productNameRaw);

  if (candidates.length === 1) {
    const only = candidates[0];
    if (!product) return only;
    // Guard against blindly attaching a report row to an account's one
    // remaining line when the two are obviously different kinds of service
    // (e.g. a "Mobile Line 1" report row landing on an account whose only
    // line on file is "Premier" internet -- that's not a $ mismatch on the
    // same service, it's a missing Mobile line for this account). Only
    // auto-match when neither side has a recognizable, conflicting
    // category -- a generic/legacy plan_name (category unknown) still
    // matches as before.
    const prodCat = productCategory(product);
    const planCat = productCategory(only.plan_name);
    if (!prodCat || !planCat || prodCat === planCat) return only;
    return null;
  }

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

  // 4) TV plan rename: providers occasionally rename a TV plan (e.g. "TV
  //    Stream" becomes "TV Essentials"). If the reported product is a TV
  //    plan and exactly one not-yet-claimed TV-category line remains on
  //    this account, treat it as the same line under its new name rather
  //    than leaving it unmatched -- renamedFromIfTvRename() (used by the
  //    caller) then flags the old->new name so confirmRows() updates
  //    plan_name and the review screen tells the admin what changed. If
  //    more than one TV candidate remains, stay unmatched -- same "don't
  //    guess" rule as above.
  if (isTvServiceText(product)) {
    const tvCandidates = candidates.filter((l) => isTvServiceText(l.plan_name));
    if (tvCandidates.length === 1) return tvCandidates[0];
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
    if (status === "mismatch") return `<span class="badge error">Commission Mismatch</span>`;
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
        renamedFrom: renamedFromIfTvRename(candidate, pr.commissionProductNameRaw),
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
        <p class="muted">Green = auto-matched and amount agrees. Yellow = "Commission Mismatch" -- the account/product matched fine, only the $ amount differs from what's expected (will be saved but flagged). Purple = a negative amount, i.e. a chargeback -- confirming it will NOT overwrite this line's received commission, it's recorded separately and shown on the Chargebacks report. Rows with no match need a manual selection before confirming. If a TV plan's name changed on the provider's report (e.g. renamed), it's noted under the match and the line's saved name is updated to match on save.</p>
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
                const chargeback = isChargebackRow(r);
                const mismatch =
                  !chargeback &&
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
                        ? `<span class="badge ${chargeback ? "chargeback" : mismatch ? "error" : "ok"}">${chargeback ? "Chargeback" : mismatch ? "Commission Mismatch" : "Matched"}</span>
                           <div class="muted">${escapeHtml(lineLabel(line))}</div>
                           ${
                             r.renamedFrom
                               ? `<div class="muted">Product name changed: "${escapeHtml(r.renamedFrom)}" &rarr; "${escapeHtml(r.commissionProductNameRaw)}" (will update on save)</div>`
                               : ""
                           }`
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
    const line = r.matchedLineId ? linesForProvider.find((l) => l.id === r.matchedLineId) : null;
    r.renamedFrom = renamedFromIfTvRename(line, r.commissionProductNameRaw);
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
        const chargeback = isChargebackRow(r);
        const mismatch =
          !chargeback &&
          line &&
          line.expected_commission != null &&
          r.commissionAmount != null &&
          Math.abs(Number(line.expected_commission) - Number(r.commissionAmount)) > AMOUNT_EPSILON;

        // Chargebacks are a separate financial event from the original
        // commission -- leave the matched line's actual_commission_amount
        // and status exactly as they are (don't clobber a real "received"
        // amount with a negative number). The chargeback itself is still
        // recorded below via commission_report_rows, which is what the
        // Chargebacks report reads from.
        //
        // A TV plan rename (renamedFrom set, regardless of chargeback) is a
        // catalog-identity fix, not a commission event, so it's applied
        // independently -- it also fixes plan_name going forward so future
        // reports for this account match on the exact name again instead of
        // needing the rename fallback in pickCandidate() every time.
        const updateFields = {};
        if (r.renamedFrom) updateFields.plan_name = r.commissionProductNameRaw;
        if (!chargeback) {
          Object.assign(updateFields, {
            actual_commission_amount: r.commissionAmount,
            commission_matched_at: new Date().toISOString(),
            commission_matched_by: ctx.profile.id,
            status: mismatch ? "mismatch" : "received",
          });
        }
        if (Object.keys(updateFields).length > 0) {
          const { error: updErr } = await supabase
            .from("order_service_lines")
            .update(updateFields)
            .eq("id", r.matchedLineId);
          if (updErr) throw updErr;
        }

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
                          <option value="mismatch" ${l.status === "mismatch" ? "selected" : ""}>Commission Mismatch</option>
                          <option value="no_report" ${l.status === "no_report" ? "selected" : ""}>No Report (manual)</option>
                          <option value="missing" ${l.status === "missing" ? "selected" : ""}>Missing</option>
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
