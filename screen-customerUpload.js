import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";
import { checkDuplicateSuspect, findExactDuplicateLine } from "./duplicateCheck.js";
import {
  normalizeAccountNumber,
  normalizePhone,
  normalizeDate,
  escapeHtml,
  fmtMoney,
} from "./normalize.js";

const FIELD_DEFS = [
  { key: "name", label: "Name", guesses: ["name", "이름", "고객명", "customer"] },
  { key: "phone", label: "Phone", guesses: ["phone", "전화", "연락처", "tel"] },
  { key: "address", label: "Address", guesses: ["address", "주소"] },
  {
    key: "account_number",
    label: "Account Number",
    guesses: ["account", "acc", "계정번호", "계정"],
  },
  { key: "order_date", label: "Order Date", guesses: ["date", "날짜", "가입일", "order date"] },
  { key: "provider", label: "Provider", guesses: ["provider", "프로바이더", "carrier"] },
  { key: "service", label: "Service / Plan", guesses: ["service", "plan", "서비스", "상품"] },
  {
    key: "salesperson",
    label: "Salesperson",
    guesses: ["sales", "세일즈", "담당자", "sold by", "sale by"],
  },
  {
    key: "expected_commission",
    label: "Expected Commission",
    guesses: ["commission", "커미션", "amount", "금액"],
  },
  {
    key: "memo",
    label: "Memo / Notes",
    guesses: ["memo", "note", "notes", "비고", "메모"],
  },
];

// ---- Two-file merge mode: field definitions per source file. ----
// Order matters for auto-detection: more specific labels (e.g. "units_installed",
// "package_group") must be listed BEFORE the more generic ones they overlap with
// ("units", "package") so the substring-matching fallback pass doesn't let the
// generic field steal a column meant for the specific one.
const FILE_A_FIELDS = [
  { key: "order_id", label: "Order Id", guesses: ["order id", "orderid"], required: true },
  { key: "name", label: "Customer Full Name", guesses: ["customer full name", "full name", "customer name"] },
  { key: "first_name", label: "First Name", guesses: ["first name", "firstname"] },
  { key: "last_name", label: "Last Name", guesses: ["last name", "lastname"] },
  { key: "address", label: "Address", guesses: ["address"] },
  { key: "city", label: "City", guesses: ["city"] },
  { key: "state", label: "State", guesses: ["state"] },
  { key: "zip", label: "ZipCode", guesses: ["zipcode", "zip code", "zip"] },
  { key: "phone", label: "ServicePhone", guesses: ["servicephone", "service phone"] },
  { key: "units_installed", label: "Units Installed", guesses: ["units installed"] },
  { key: "units", label: "Units", guesses: ["units"] },
  { key: "mobile_lines_ordered", label: "Mobile Lines Ordered", guesses: ["mobile lines ordered"] },
  { key: "mobile_lines_installed", label: "Mobile Lines Installed", guesses: ["mobile lines installed"] },
  { key: "package_group", label: "Package Group", guesses: ["package group"] },
  { key: "package", label: "Package", guesses: ["package"] },
  { key: "product", label: "Product", guesses: ["product"] },
  { key: "account_number", label: "Account Number", guesses: ["account number", "account"] },
];

const FILE_B_FIELDS = [
  { key: "order_id", label: "Order Id", guesses: ["order id", "orderid"], required: true },
  { key: "order_date", label: "Order Date", guesses: ["order date"] },
  { key: "name", label: "Customer Full Name", guesses: ["customer full name", "full name", "customer name"] },
  { key: "first_name", label: "First Name", guesses: ["first name", "firstname"] },
  { key: "last_name", label: "Last Name", guesses: ["last name", "lastname"] },
  { key: "address", label: "Address", guesses: ["address"] },
  { key: "city", label: "City", guesses: ["city"] },
  { key: "state", label: "State", guesses: ["state"] },
  { key: "zip", label: "ZipCode", guesses: ["zipcode", "zip code", "zip"] },
  { key: "phone", label: "ServicePhone", guesses: ["servicephone", "service phone"] },
  { key: "units_installed", label: "Units Installed", guesses: ["units installed"] },
  { key: "units", label: "Units", guesses: ["units"] },
  { key: "account_number", label: "Account Number", guesses: ["account number", "account"] },
  { key: "order_number", label: "Order Number / Work Order", guesses: ["work order", "order number"] },
  { key: "category", label: "Category", guesses: ["category"] },
  { key: "salesperson", label: "SalesIdName", guesses: ["salesidname", "sales id name", "sales id", "salesperson"] },
];

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

// Two-pass auto-detect: exact header match first, then substring match, each
// pass respecting field order and never reusing an already-claimed column.
function autoDetectMapping(headers, fieldDefs) {
  const lowerHeaders = headers.map((h) => h.toLowerCase().trim());
  const used = new Set();
  const mapping = {};

  fieldDefs.forEach((f) => {
    for (let i = 0; i < lowerHeaders.length; i++) {
      if (used.has(i)) continue;
      if (f.guesses.some((g) => lowerHeaders[i] === g.toLowerCase())) {
        mapping[f.key] = i;
        used.add(i);
        break;
      }
    }
  });
  fieldDefs.forEach((f) => {
    if (mapping[f.key] !== undefined) return;
    for (let i = 0; i < lowerHeaders.length; i++) {
      if (used.has(i)) continue;
      if (f.guesses.some((g) => lowerHeaders[i].includes(g.toLowerCase()))) {
        mapping[f.key] = i;
        used.add(i);
        break;
      }
    }
  });
  fieldDefs.forEach((f) => {
    if (mapping[f.key] === undefined) mapping[f.key] = -1;
  });
  return mapping;
}

function normText(v) {
  return String(v ?? "").trim().toLowerCase().replace(/\s+/g, " ");
}
function normNum(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}
function joinAddress(address, city, state, zip) {
  const parts = [address, city, [state, zip].filter((p) => p && String(p).trim() !== "").join(" ")].filter(
    (p) => p && String(p).trim() !== ""
  );
  return parts.join(", ");
}

// Fuzzy match a raw text value against a master list (services/providers) by
// name or display_name — exact match first, then substring either direction.
function matchMaster(list, rawValue) {
  if (!rawValue) return null;
  const norm = String(rawValue).trim().toLowerCase();
  if (!norm) return null;
  const hit = list.find(
    (m) =>
      m.name.toLowerCase() === norm ||
      (m.display_name && m.display_name.toLowerCase() === norm) ||
      m.name.toLowerCase().includes(norm) ||
      norm.includes(m.name.toLowerCase())
  );
  return hit ? hit.id : null;
}

// Try a prioritized list of raw text candidates (e.g. Product, then Package,
// then Package Group) against a master list, and report which candidate (if
// any) was actually used and whether it matched — so a failed match can be
// shown to staff with the exact text that didn't match, instead of quietly
// disappearing behind a generic "?" in the saved data.
function matchMasterPrioritized(list, rawValues) {
  const candidates = (rawValues || []).filter((v) => v && String(v).trim() !== "");
  for (const raw of candidates) {
    const id = matchMaster(list, raw);
    if (id) return { id, raw, matched: true };
  }
  return { id: null, raw: candidates[0] || null, matched: candidates.length === 0 };
}

// ---- Commission rate auto-lookup (Master Data > Commission Rates) ----
// Looks up the $ amount for a given provider + plan name so "Expected
// Commission" can be pre-filled during upload instead of typed by hand.
// Staff can still edit the value in the review grid before saving either way.
function isMobileText(s) {
  return String(s || "").trim().toLowerCase() === "mobile";
}

function lookupCommissionRate(commissionRates, providerId, planName) {
  if (!providerId || !planName) return null;
  const target = String(planName).trim().toLowerCase();
  if (!target) return null;
  const match = commissionRates.find(
    (r) => r.provider_id === providerId && String(r.plan_name).trim().toLowerCase() === target
  );
  return match ? Number(match.rate) : null;
}

// `candidates` is a prioritized list of raw text fields that might hold the
// plan/package name (e.g. Product, then Package, then Package Group, then
// the free-text Service value) -- the first one with an exact match in the
// commission rate table wins. Mobile lines are a special case: the source
// data usually just says "Mobile" with no specific plan name, so the $
// amount instead depends on which mobile line number this is within the
// order (1st line priced differently from the 2nd+, per the rate table's
// "Mobile Line 1" / "Mobile Line 2" ... rows) -- `mobileLineIndex` is that
// 1-based position, tracked by the caller per block/order.
function autoLookupCommission(commissionRates, providerId, candidates, mobileLineIndex) {
  const list = (candidates || []).filter((v) => v && String(v).trim() !== "");
  if (list.some((c) => isMobileText(c))) {
    return mobileLineIndex ? lookupCommissionRate(commissionRates, providerId, `Mobile Line ${mobileLineIndex}`) : null;
  }
  for (const c of list) {
    const rate = lookupCommissionRate(commissionRates, providerId, c);
    if (rate !== null) return rate;
  }
  return null;
}

export async function renderCustomerUpload(container, ctx) {
  let mode = "single"; // "single" | "two"
  let step = "upload"; // upload -> mapping -> review (single mode); upload -> mappingTwo -> mergeReview -> review (two-file mode)

  // ---- single-file mode state ----
  let workbook = null;
  let sheetNames = [];
  let selectedSheet = "";
  let headers = [];
  let dataRows = [];
  let mapping = {}; // field key -> column index (or -1)

  // ---- two-file mode state ----
  let fileAWorkbook = null;
  let fileASheetNames = [];
  let fileASelectedSheet = "";
  let fileAHeaders = [];
  let fileADataRows = [];
  let fileAMapping = {};
  let fileBWorkbook = null;
  let fileBSheetNames = [];
  let fileBSelectedSheet = "";
  let fileBHeaders = [];
  let fileBDataRows = [];
  let fileBMapping = {};
  let mergedRows = [];

  let defaultProviderId = "";
  let blocks = []; // built after mapping confirmed
  let services = [];
  let providers = [];
  let salespeople = [];
  let commissionRates = [];
  let busyMsg = "";

  async function loadMasters() {
    const [s, p, sp, cr] = await Promise.all([
      supabase.from("services").select("*").eq("is_active", true).order("sort_order"),
      supabase.from("providers").select("*").eq("is_active", true).order("name"),
      supabase.from("salespeople").select("*").eq("is_active", true).order("name"),
      supabase.from("commission_rates").select("provider_id, plan_name, rate").eq("is_active", true),
    ]);
    services = s.data || [];
    providers = p.data || [];
    salespeople = sp.data || [];
    commissionRates = cr.data || [];
  }

  // Shared header (title + Reset) rendered at the top of every step, so
  // staff can always discard the in-progress upload and start over without
  // having to click "Back" through each step.
  function screenHeaderHtml(title) {
    return `
      <div class="screen-header-row">
        <h2>${title}</h2>
        <button class="btn" id="reset-upload-btn" type="button">Reset</button>
      </div>
    `;
  }

  function wireResetBtn() {
    container.querySelector("#reset-upload-btn")?.addEventListener("click", () => {
      const hasProgress = !!(workbook || fileAWorkbook || fileBWorkbook || blocks.length > 0 || mergedRows.length > 0);
      if (hasProgress && !confirm("Reset and discard the current upload? Nothing will be saved.")) return;
      resetAll();
    });
  }

  function resetAll() {
    mode = "single";
    step = "upload";
    workbook = null;
    sheetNames = [];
    selectedSheet = "";
    headers = [];
    dataRows = [];
    mapping = {};
    fileAWorkbook = null;
    fileASheetNames = [];
    fileASelectedSheet = "";
    fileAHeaders = [];
    fileADataRows = [];
    fileAMapping = {};
    fileBWorkbook = null;
    fileBSheetNames = [];
    fileBSelectedSheet = "";
    fileBHeaders = [];
    fileBDataRows = [];
    fileBMapping = {};
    mergedRows = [];
    defaultProviderId = "";
    blocks = [];
    busyMsg = "";
    drawUploadStep();
  }

  function drawUploadStep() {
    container.innerHTML = `
      <div class="screen">
        ${screenHeaderHtml("Customer Excel Upload")}
        <p class="muted">Upload an Excel (.xlsx/.csv) file, or two files that share an Order Id column and need to be matched and merged first. Everything goes through column mapping and a review/confirm step before being saved as real customer and order data.</p>
        <div class="card">
          <label>Provider for this upload
            <select id="provider-select-upload">
              <option value="">(Not set — choose before uploading)</option>
              ${providers.map((p) => `<option value="${p.id}" ${defaultProviderId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">Applied to every row that doesn't specify (or can't be matched to) a provider of its own. You can still change it per line on the review screen.</p>
        </div>
        <div class="tabs">
          <button data-mode="single" class="${mode === "single" ? "tab active" : "tab"}">Single File</button>
          <button data-mode="two" class="${mode === "two" ? "tab active" : "tab"}">Two Files (Service Info + Order Info)</button>
        </div>
        ${
          mode === "single"
            ? `<div class="card">
                <input type="file" id="file-input" accept=".xlsx,.xls,.csv" />
              </div>`
            : `<div class="card">
                <p class="muted">Upload the "Service Info" file (has Product/Package/Package Group, e.g. a Bundle Orders export) and the "Order Info" file (has Order Date and SalesIdName, e.g. an order list export). Rows are matched by Order Id automatically once both are selected.</p>
                <div class="grid3">
                  <label>Service Info File
                    <input type="file" id="file-input-a" accept=".xlsx,.xls,.csv" />
                  </label>
                  <label>Order Info File
                    <input type="file" id="file-input-b" accept=".xlsx,.xls,.csv" />
                  </label>
                </div>
              </div>`
        }
      </div>
    `;

    wireResetBtn();
    container.querySelector("#provider-select-upload").addEventListener("change", (e) => {
      defaultProviderId = e.target.value;
    });
    container.querySelectorAll("[data-mode]").forEach((btn) => {
      btn.addEventListener("click", () => {
        mode = btn.dataset.mode;
        drawUploadStep();
      });
    });

    if (mode === "single") {
      container.querySelector("#file-input").addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const buf = await file.arrayBuffer();
        workbook = XLSX.read(buf, { cellDates: true });
        sheetNames = workbook.SheetNames;
        selectedSheet = sheetNames[0];
        await loadMasters();
        loadSheet(selectedSheet);
        step = "mapping";
        drawMappingStep();
      });
    } else {
      container.querySelector("#file-input-a").addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const buf = await file.arrayBuffer();
        fileAWorkbook = XLSX.read(buf, { cellDates: true });
        fileASheetNames = fileAWorkbook.SheetNames;
        fileASelectedSheet = fileASheetNames[0];
        await maybeProceedToMappingTwo();
      });
      container.querySelector("#file-input-b").addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const buf = await file.arrayBuffer();
        fileBWorkbook = XLSX.read(buf, { cellDates: true });
        fileBSheetNames = fileBWorkbook.SheetNames;
        fileBSelectedSheet = fileBSheetNames[0];
        await maybeProceedToMappingTwo();
      });
    }
  }

  async function maybeProceedToMappingTwo() {
    if (!fileAWorkbook || !fileBWorkbook) return; // wait until both files are chosen
    await loadMasters();
    loadSheetA(fileASelectedSheet);
    loadSheetB(fileBSelectedSheet);
    step = "mappingTwo";
    drawMappingTwoStep();
  }

  function loadSheet(name) {
    const sheet = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    headers = (rows[0] || []).map((h) => String(h ?? "").trim());
    dataRows = rows.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
    mapping = {};
    FIELD_DEFS.forEach((f) => {
      const idx = headers.findIndex((h) =>
        f.guesses.some((g) => h.toLowerCase().includes(g.toLowerCase()))
      );
      mapping[f.key] = idx;
    });
  }

  function loadSheetA(name) {
    const sheet = fileAWorkbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    fileAHeaders = (rows[0] || []).map((h) => String(h ?? "").trim());
    fileADataRows = rows.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
    fileAMapping = autoDetectMapping(fileAHeaders, FILE_A_FIELDS);
  }

  function loadSheetB(name) {
    const sheet = fileBWorkbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    fileBHeaders = (rows[0] || []).map((h) => String(h ?? "").trim());
    fileBDataRows = rows.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
    fileBMapping = autoDetectMapping(fileBHeaders, FILE_B_FIELDS);
  }

  function drawMappingStep() {
    container.innerHTML = `
      <div class="screen">
        ${screenHeaderHtml("Step 2: Column Mapping &amp; Validation")}
        <div class="card">
          <label>Select Sheet
            <select id="sheet-select">
              ${sheetNames.map((n) => `<option value="${escapeHtml(n)}" ${n === selectedSheet ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">Detected ${headers.length} column(s) / ${dataRows.length} data row(s). Check and adjust which column maps to each system field below.</p>
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
          <label>Default provider for rows with no/empty provider column
            <select id="default-provider">
              <option value="">(Not set)</option>
              ${providers.map((p) => `<option value="${p.id}" ${defaultProviderId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
            </select>
          </label>
          <div class="btn-row">
            <button class="btn" id="back-btn">Re-upload</button>
            <button class="btn primary" id="next-btn">Next: Review &amp; Confirm</button>
          </div>
        </div>
      </div>
    `;

    wireResetBtn();
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
    container.querySelector("#default-provider").addEventListener("change", (e) => {
      defaultProviderId = e.target.value;
    });
    container.querySelector("#back-btn").addEventListener("click", () => {
      step = "upload";
      drawUploadStep();
    });
    container.querySelector("#next-btn").addEventListener("click", async () => {
      defaultProviderId = container.querySelector("#default-provider").value;
      await buildBlocks();
      step = "review";
      drawReviewStep();
    });
  }

  function drawMappingTwoStep() {
    container.innerHTML = `
      <div class="screen">
        ${screenHeaderHtml("Step 2: Column Mapping (Two Files)")}
        <div class="card">
          <h3>Service Info File</h3>
          <label>Select Sheet
            <select id="sheet-select-a">
              ${fileASheetNames.map((n) => `<option value="${escapeHtml(n)}" ${n === fileASelectedSheet ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">Detected ${fileAHeaders.length} column(s) / ${fileADataRows.length} data row(s).</p>
          <table class="data-table">
            <thead><tr><th>System Field</th><th>Excel Column</th></tr></thead>
            <tbody>
              ${FILE_A_FIELDS.map(
                (f) => `
                <tr>
                  <td>${f.label}${f.required ? " *" : ""}</td>
                  <td>
                    <select data-file="a" data-field="${f.key}">
                      <option value="-1" ${fileAMapping[f.key] === -1 ? "selected" : ""}>(Not used)</option>
                      ${fileAHeaders
                        .map(
                          (h, i) =>
                            `<option value="${i}" ${fileAMapping[f.key] === i ? "selected" : ""}>${escapeHtml(h) || "(untitled " + i + ")"}</option>`
                        )
                        .join("")}
                    </select>
                  </td>
                </tr>`
              ).join("")}
            </tbody>
          </table>
        </div>
        <div class="card">
          <h3>Order Info File</h3>
          <label>Select Sheet
            <select id="sheet-select-b">
              ${fileBSheetNames.map((n) => `<option value="${escapeHtml(n)}" ${n === fileBSelectedSheet ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">Detected ${fileBHeaders.length} column(s) / ${fileBDataRows.length} data row(s).</p>
          <table class="data-table">
            <thead><tr><th>System Field</th><th>Excel Column</th></tr></thead>
            <tbody>
              ${FILE_B_FIELDS.map(
                (f) => `
                <tr>
                  <td>${f.label}${f.required ? " *" : ""}</td>
                  <td>
                    <select data-file="b" data-field="${f.key}">
                      <option value="-1" ${fileBMapping[f.key] === -1 ? "selected" : ""}>(Not used)</option>
                      ${fileBHeaders
                        .map(
                          (h, i) =>
                            `<option value="${i}" ${fileBMapping[f.key] === i ? "selected" : ""}>${escapeHtml(h) || "(untitled " + i + ")"}</option>`
                        )
                        .join("")}
                    </select>
                  </td>
                </tr>`
              ).join("")}
            </tbody>
          </table>
        </div>
        <div class="card">
          <label>Default Provider (applied to every matched order — these files have no provider column)
            <select id="default-provider-two">
              <option value="">(Not set)</option>
              ${providers.map((p) => `<option value="${p.id}" ${defaultProviderId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
            </select>
          </label>
        </div>
        <div class="btn-row">
          <button class="btn" id="back-to-upload-btn">Re-upload</button>
          <button class="btn primary" id="merge-btn">Next: Match &amp; Merge by Order Id</button>
        </div>
      </div>
    `;

    wireResetBtn();
    container.querySelector("#sheet-select-a").addEventListener("change", (e) => {
      fileASelectedSheet = e.target.value;
      loadSheetA(fileASelectedSheet);
      drawMappingTwoStep();
    });
    container.querySelector("#sheet-select-b").addEventListener("change", (e) => {
      fileBSelectedSheet = e.target.value;
      loadSheetB(fileBSelectedSheet);
      drawMappingTwoStep();
    });
    container.querySelectorAll("[data-file]").forEach((sel) => {
      sel.addEventListener("change", (e) => {
        const val = Number(e.target.value);
        if (sel.dataset.file === "a") fileAMapping[sel.dataset.field] = val;
        else fileBMapping[sel.dataset.field] = val;
      });
    });
    container.querySelector("#default-provider-two").addEventListener("change", (e) => {
      defaultProviderId = e.target.value;
    });
    container.querySelector("#back-to-upload-btn").addEventListener("click", () => {
      step = "upload";
      fileAWorkbook = null;
      fileBWorkbook = null;
      drawUploadStep();
    });
    container.querySelector("#merge-btn").addEventListener("click", () => {
      if (fileAMapping.order_id === -1 || fileBMapping.order_id === -1) {
        alert("Please map the Order Id column in both files before merging.");
        return;
      }
      buildMergedRows();
      step = "mergeReview";
      drawMergeReviewStep();
    });
  }

  function cell(row, fieldKey) {
    const idx = mapping[fieldKey];
    if (idx === undefined || idx === -1) return "";
    return row[idx] ?? "";
  }

  function cellByMapping(row, fieldMapping, fieldKey) {
    const idx = fieldMapping[fieldKey];
    if (idx === undefined || idx === -1) return "";
    return row[idx] ?? "";
  }

  async function buildBlocks() {
    blocks = [];
    let current = null;
    let mobileLineCounter = 0; // resets per block -- see autoLookupCommission
    for (const row of dataRows) {
      const name = String(cell(row, "name") ?? "").trim();
      const isNewBlock = name !== "";
      if (isNewBlock) {
        mobileLineCounter = 0;
        current = {
          id: crypto.randomUUID(),
          name,
          phone: String(cell(row, "phone") ?? "").trim(),
          address: String(cell(row, "address") ?? "").trim(),
          orderDate: normalizeDate(cell(row, "order_date")),
          salesperson: String(cell(row, "salesperson") ?? "").trim(),
          memo: String(cell(row, "memo") ?? "").trim(),
          excluded: false,
          duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
          linkToExisting: false,
          sourceOrderId: null,
          orderNumber: null,
          category: null,
          needsReview: false,
          reviewReasons: [],
          followUpNeeded: false,
          followUpDate: null,
          followUpReason: "",
          lines: [],
        };
        blocks.push(current);
      }
      if (!current) continue; // skip rows with no name if they appear before any block starts (formatting error)
      const providerRaw = String(cell(row, "provider") ?? "").trim();
      const providerMatch = matchMasterPrioritized(providers, [providerRaw]);
      const provider = providerMatch.id || (defaultProviderId || null);
      // Service is taken directly from the uploaded file's Service/Product
      // column, verbatim — it is no longer matched against the Services
      // master list. The exact text is what gets saved (and auto-registered
      // into Master Data on save if it's new), so it always reflects the
      // source file instead of "?" or a fuzzy-matched substitute.
      const serviceName = String(cell(row, "service") ?? "").trim();
      if (isMobileText(serviceName)) mobileLineCounter += 1;
      const expectedCommissionFromFile = Number(cell(row, "expected_commission")) || null;
      current.lines.push({
        id: crypto.randomUUID(),
        excluded: false,
        serviceName,
        providerId: provider,
        providerMatched: providerMatch.matched,
        providerRaw: providerMatch.raw,
        accountNumberRaw: String(cell(row, "account_number") ?? "").trim(),
        accountNumber: normalizeAccountNumber(cell(row, "account_number")),
        // The uploaded file's own Expected Commission column wins if present;
        // otherwise auto-look-up the $ amount from Master Data > Commission
        // Rates by provider + plan name (staff can still edit it below).
        expectedCommission:
          expectedCommissionFromFile ??
          autoLookupCommission(commissionRates, provider, [serviceName], mobileLineCounter),
        units: null,
        unitsInstalled: null,
        mobileLinesOrdered: null,
        mobileLinesInstalled: null,
        planName: null,
      });
    }

    // Surface Provider values that didn't match anything in the master list
    // — these previously saved silently as "?"; now they're flagged so
    // staff can fix the master list or pick manually. (Service no longer
    // needs this: it's saved as free text from the source file directly.)
    for (const b of blocks) {
      const unmatchedProviders = [
        ...new Set(b.lines.filter((l) => l.providerRaw && !l.providerMatched).map((l) => l.providerRaw)),
      ];
      const reasons = [];
      if (unmatchedProviders.length) reasons.push(`Provider not matched: ${unmatchedProviders.join(", ")}`);
      b.reviewReasons = reasons;
      b.needsReview = reasons.length > 0;
    }

    // Duplicate-suspect check (name + phone + account number combination; order date is excluded from the criteria)
    for (const b of blocks) {
      const firstAccount = b.lines[0]?.accountNumber || "";
      const res = await checkDuplicateSuspect({
        name: b.name,
        phone: b.phone,
        accountNumber: firstAccount,
      });
      b.duplicate = res;
    }

    sortBlocksByReviewPriority();
  }

  // Blocks that need a human look — merge mismatches/missing data from the
  // two-file merge, or a possible duplicate customer — float to the top so
  // they aren't missed among rows that are already clean.
  function sortBlocksByReviewPriority() {
    blocks.sort((a, b) => {
      const aFlag = a.needsReview || a.duplicate.suspect ? 1 : 0;
      const bFlag = b.needsReview || b.duplicate.suspect ? 1 : 0;
      return bFlag - aFlag;
    });
  }

  // ---- two-file merge ----
  // A single Order Id can appear more than once in either file — most commonly
  // in mobile orders with 2-5 lines, where the Order Info file has one row per
  // line with its own Order Number / Work Order (e.g. 1-5) but the Service Info
  // file may still have just one (aggregate) row for that Order Id. pairRows()
  // pairs up same-Order-Id rows across both files: if one side has exactly one
  // row and the other has several, the single row is paired with every row on
  // the other side (broadcast); otherwise rows are paired by position.
  function pairRows(aRows, bRows) {
    if (aRows.length <= 1 && bRows.length > 1) {
      const aRow = aRows[0] || null;
      return bRows.map((bRow) => [aRow, bRow]);
    }
    if (bRows.length <= 1 && aRows.length > 1) {
      const bRow = bRows[0] || null;
      return aRows.map((aRow) => [aRow, bRow]);
    }
    const maxLen = Math.max(aRows.length, bRows.length, 1);
    const pairs = [];
    for (let i = 0; i < maxLen; i++) pairs.push([aRows[i] || null, bRows[i] || null]);
    return pairs;
  }

  function buildMergedRows() {
    const aGroups = new Map();
    fileADataRows.forEach((row) => {
      const oid = normText(cellByMapping(row, fileAMapping, "order_id"));
      if (!oid) return;
      if (!aGroups.has(oid)) aGroups.set(oid, []);
      aGroups.get(oid).push(row);
    });
    const bGroups = new Map();
    fileBDataRows.forEach((row) => {
      const oid = normText(cellByMapping(row, fileBMapping, "order_id"));
      if (!oid) return;
      if (!bGroups.has(oid)) bGroups.set(oid, []);
      bGroups.get(oid).push(row);
    });

    const getA = (row, key) => (row ? cellByMapping(row, fileAMapping, key) : "");
    const getB = (row, key) => (row ? cellByMapping(row, fileBMapping, key) : "");
    const resolveName = (aRow, bRow) => {
      const full = getA(aRow, "name") || getB(bRow, "name");
      if (full) return full;
      const first = getA(aRow, "first_name") || getB(bRow, "first_name");
      const last = getA(aRow, "last_name") || getB(bRow, "last_name");
      return [first, last].filter((v) => v && String(v).trim() !== "").join(" ");
    };

    const allIds = Array.from(new Set([...aGroups.keys(), ...bGroups.keys()])).sort();
    mergedRows = allIds.map((oid) => {
      const aRows = aGroups.get(oid) || [];
      const bRows = bGroups.get(oid) || [];
      const pairs = pairRows(aRows, bRows);
      const [firstA, firstB] = pairs[0] || [null, null];

      // order-level info (name/phone/address/date/salesperson) taken from the
      // first pair — every line of the same order is assumed to share it.
      const name = resolveName(firstA, firstB);
      const aPhone = getA(firstA, "phone");
      const bPhone = getB(firstB, "phone");
      const address = joinAddress(
        getA(firstA, "address") || getB(firstB, "address"),
        getA(firstA, "city") || getB(firstB, "city"),
        getA(firstA, "state") || getB(firstB, "state"),
        getA(firstA, "zip") || getB(firstB, "zip")
      );

      const aNameOnly = resolveName(firstA, null);
      const bNameOnly = resolveName(null, firstB);
      const nameMismatch = !!(
        normText(aNameOnly) && normText(bNameOnly) && normText(aNameOnly) !== normText(bNameOnly)
      );
      const phoneMismatch = !!(
        normalizePhone(aPhone) && normalizePhone(bPhone) && normalizePhone(aPhone) !== normalizePhone(bPhone)
      );

      // per-line info (one entry per paired row — e.g. one per mobile work order)
      const lines = pairs.map(([aRow, bRow]) => {
        const aAccount = getA(aRow, "account_number");
        const bAccount = getB(bRow, "account_number");
        const aUnits = getA(aRow, "units");
        const bUnits = getB(bRow, "units");
        const aUnitsInstalled = getA(aRow, "units_installed");
        const bUnitsInstalled = getB(bRow, "units_installed");
        const planName =
          [getA(aRow, "product"), getA(aRow, "package"), getA(aRow, "package_group")]
            .filter((v) => v && String(v).trim() !== "")
            .join(" / ") || null;
        // Service is taken directly from the Product column (per the mapping
        // reference the provider gave us), falling back to Package / Package
        // Group only when Product is blank — no matching against the
        // Services master list, so the exact source text is what gets saved.
        const serviceName = String(
          [getA(aRow, "product"), getA(aRow, "package"), getA(aRow, "package_group")].find(
            (v) => v && String(v).trim() !== ""
          ) || ""
        ).trim();
        return {
          serviceName,
          accountNumber: normalizeAccountNumber(aAccount || bAccount),
          accountMismatch: !!(
            normalizeAccountNumber(aAccount) &&
            normalizeAccountNumber(bAccount) &&
            normalizeAccountNumber(aAccount) !== normalizeAccountNumber(bAccount)
          ),
          units: normNum(aUnits) !== null ? normNum(aUnits) : normNum(bUnits),
          unitsMismatch: !!(normNum(aUnits) !== null && normNum(bUnits) !== null && normNum(aUnits) !== normNum(bUnits)),
          unitsInstalled: normNum(aUnitsInstalled) !== null ? normNum(aUnitsInstalled) : normNum(bUnitsInstalled),
          unitsInstalledMismatch: !!(
            normNum(aUnitsInstalled) !== null &&
            normNum(bUnitsInstalled) !== null &&
            normNum(aUnitsInstalled) !== normNum(bUnitsInstalled)
          ),
          mobileLinesOrdered: normNum(getA(aRow, "mobile_lines_ordered")),
          mobileLinesInstalled: normNum(getA(aRow, "mobile_lines_installed")),
          packageGroup: getA(aRow, "package_group"),
          package: getA(aRow, "package"),
          product: getA(aRow, "product"),
          planName,
          orderNumber: getB(bRow, "order_number"),
        };
      });

      return {
        id: crypto.randomUUID(),
        orderId: oid,
        onlyInA: bRows.length === 0,
        onlyInB: aRows.length === 0,
        excluded: false,
        name,
        phone: aPhone || bPhone,
        address,
        orderDate: normalizeDate(getB(firstB, "order_date")),
        salesperson: getB(firstB, "salesperson"),
        category: getB(firstB, "category"),
        nameMismatch,
        phoneMismatch,
        lines,
      };
    });
  }

  // A merged row needs a human look if it came from only one of the two
  // source files, if a field the two files both reported disagrees, or if a
  // line's Product/Package didn't match anything in the Services master list.
  function rowNeedsReview(r) {
    const anyAccountMismatch = r.lines.some((l) => l.accountMismatch);
    const anyUnitsMismatch = r.lines.some((l) => l.unitsMismatch);
    const anyUnitsInstalledMismatch = r.lines.some((l) => l.unitsInstalledMismatch);
    return !!(
      r.onlyInA ||
      r.onlyInB ||
      r.nameMismatch ||
      r.phoneMismatch ||
      anyAccountMismatch ||
      anyUnitsMismatch ||
      anyUnitsInstalledMismatch
    );
  }

  function drawMergeReviewStep() {
    const includedCount = mergedRows.filter((r) => !r.excluded).length;
    const matchedBothCount = mergedRows.filter((r) => !r.onlyInA && !r.onlyInB).length;
    const onlyACount = mergedRows.filter((r) => r.onlyInA).length;
    const onlyBCount = mergedRows.filter((r) => r.onlyInB).length;
    const multiLineCount = mergedRows.filter((r) => r.lines.length > 1).length;
    const reviewCount = mergedRows.filter((r) => rowNeedsReview(r)).length;

    // Rows that need review (missing from one file, or a value mismatch
    // between the two files) are listed first so they aren't missed among
    // the already-clean rows.
    const sortedRows = [...mergedRows].sort((a, b) => {
      const aNeeds = rowNeedsReview(a) ? 1 : 0;
      const bNeeds = rowNeedsReview(b) ? 1 : 0;
      return bNeeds - aNeeds;
    });

    container.innerHTML = `
      <div class="screen">
        ${screenHeaderHtml("Step 3: Merge Result Review")}
        <p class="muted">Matched ${mergedRows.length} order(s) by Order Id — ${matchedBothCount} found in both files, ${onlyACount} only in the Service Info file (missing Order Date), ${onlyBCount} only in the Order Info file (missing Product/Package)${
      multiLineCount ? `, ${multiLineCount} order(s) expanded into multiple service lines (e.g. mobile orders with several Order Number / Work Order values)` : ""
    }. ${reviewCount} row(s) marked "Please review" below need a look and are listed first — you can fix values here, or on the next screen where Units / Units Installed and Service/Provider selection per line can also be edited.</p>
        <div class="btn-row">
          <button class="btn primary" id="merge-next-btn">Next: Review &amp; Confirm (${includedCount})</button>
          <button class="btn" id="merge-back-btn">Back (Column Mapping)</button>
        </div>
        <table class="data-table small">
          <thead>
            <tr><th>Review</th><th>Include</th><th>Order Id</th><th>Customer</th><th>Lines</th><th>Status</th><th>Flags</th></tr>
          </thead>
          <tbody>
            ${
              sortedRows
                .map((r) => {
                  const anyAccountMismatch = r.lines.some((l) => l.accountMismatch);
                  const anyUnitsMismatch = r.lines.some((l) => l.unitsMismatch);
                  const anyUnitsInstalledMismatch = r.lines.some((l) => l.unitsInstalledMismatch);
                  const noFlags =
                    !r.nameMismatch &&
                    !r.phoneMismatch &&
                    !anyAccountMismatch &&
                    !anyUnitsMismatch &&
                    !anyUnitsInstalledMismatch;
                  const needsReview = rowNeedsReview(r);
                  return `
              <tr data-mrow="${r.id}" class="${r.excluded ? "excluded" : ""} ${needsReview ? "needs-review-row" : ""}">
                <td>${needsReview ? `<span class="badge error">Please review</span>` : `<span class="muted">-</span>`}</td>
                <td><input type="checkbox" class="mr-include" ${r.excluded ? "" : "checked"} /></td>
                <td>${escapeHtml(r.orderId)}</td>
                <td>${escapeHtml(r.name || "-")}</td>
                <td>${r.lines.length}${r.lines.length > 1 ? ` <span class="badge neutral">multi-line</span>` : ""}</td>
                <td>
                  ${r.onlyInA ? `<span class="badge warn">Only in Service Info file</span>` : ""}
                  ${r.onlyInB ? `<span class="badge warn">Only in Order Info file</span>` : ""}
                  ${!r.onlyInA && !r.onlyInB ? `<span class="badge ok">Matched</span>` : ""}
                </td>
                <td>
                  ${r.nameMismatch ? `<span class="badge error">Name differs</span>` : ""}
                  ${r.phoneMismatch ? `<span class="badge error">Phone differs</span>` : ""}
                  ${anyAccountMismatch ? `<span class="badge error">Account # differs</span>` : ""}
                  ${anyUnitsMismatch ? `<span class="badge error">Units differ</span>` : ""}
                  ${anyUnitsInstalledMismatch ? `<span class="badge error">Units Installed differ</span>` : ""}
                  ${noFlags ? `<span class="muted">-</span>` : ""}
                </td>
              </tr>`;
                })
                .join("") || `<tr><td colspan="7" class="muted">No rows to merge.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    `;
    wireResetBtn();
    wireMergeReviewEvents();
  }

  function wireMergeReviewEvents() {
    container.querySelector("#merge-back-btn").addEventListener("click", () => {
      step = "mappingTwo";
      drawMappingTwoStep();
    });
    container.querySelectorAll("tr[data-mrow]").forEach((tr) => {
      const r = mergedRows.find((x) => x.id === tr.dataset.mrow);
      tr.querySelector(".mr-include").addEventListener("change", (e) => {
        r.excluded = !e.target.checked;
      });
    });
    container.querySelector("#merge-next-btn").addEventListener("click", async () => {
      await buildBlocksFromMerge();
      step = "review";
      drawReviewStep();
    });
  }

  async function buildBlocksFromMerge() {
    blocks = [];
    for (const r of mergedRows.filter((x) => !x.excluded)) {
      const anyAccountMismatch = r.lines.some((l) => l.accountMismatch);
      const anyUnitsMismatch = r.lines.some((l) => l.unitsMismatch);
      const anyUnitsInstalledMismatch = r.lines.some((l) => l.unitsInstalledMismatch);
      const reviewReasons = [];
      if (r.onlyInA) reviewReasons.push("Only in Service Info file (Order Date missing)");
      if (r.onlyInB) reviewReasons.push("Only in Order Info file (Product/Package missing)");
      if (r.nameMismatch) reviewReasons.push("Name differs between files");
      if (r.phoneMismatch) reviewReasons.push("Phone differs between files");
      if (anyAccountMismatch) reviewReasons.push("Account # differs between files");
      if (anyUnitsMismatch) reviewReasons.push("Units differ between files");
      if (anyUnitsInstalledMismatch) reviewReasons.push("Units Installed differ between files");
      if (!defaultProviderId) reviewReasons.push("Provider not selected — pick a default provider or set it per line below");
      blocks.push({
        id: crypto.randomUUID(),
        name: r.name || "",
        phone: r.phone || "",
        address: r.address || "",
        orderDate: r.orderDate,
        salesperson: r.salesperson || "",
        excluded: false,
        duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
        linkToExisting: false,
        sourceOrderId: r.orderId,
        orderNumber: r.lines.map((l) => l.orderNumber).filter(Boolean).join(", ") || null,
        category: r.category || null,
        needsReview: reviewReasons.length > 0,
        reviewReasons,
        followUpNeeded: false,
        followUpDate: null,
        followUpReason: "",
        lines: (() => {
          let mobileLineCounter = 0; // resets per order/block -- see autoLookupCommission
          return r.lines.map((l) => {
            if (isMobileText(l.product) || isMobileText(l.package) || isMobileText(l.packageGroup) || isMobileText(l.serviceName)) {
              mobileLineCounter += 1;
            }
            return {
              id: crypto.randomUUID(),
              excluded: false,
              serviceName: l.serviceName,
              providerId: defaultProviderId || null,
              providerMatched: true,
              providerRaw: null,
              accountNumberRaw: l.accountNumber || "",
              accountNumber: l.accountNumber || "",
              // No Expected Commission column in this format -- always
              // auto-look-up the $ amount from Master Data > Commission
              // Rates by provider + plan name (staff can still edit it below).
              expectedCommission: autoLookupCommission(
                commissionRates,
                defaultProviderId || null,
                [l.product, l.package, l.packageGroup, l.planName, l.serviceName],
                mobileLineCounter
              ),
              units: l.units,
              unitsInstalled: l.unitsInstalled,
              mobileLinesOrdered: l.mobileLinesOrdered,
              mobileLinesInstalled: l.mobileLinesInstalled,
              planName: l.planName,
            };
          });
        })(),
      });
    }

    // Duplicate-suspect check (same criteria as single-file mode)
    for (const b of blocks) {
      const firstAccount = b.lines[0]?.accountNumber || "";
      const res = await checkDuplicateSuspect({
        name: b.name,
        phone: b.phone,
        accountNumber: firstAccount,
      });
      b.duplicate = res;
    }

    sortBlocksByReviewPriority();
  }

  function salespersonDatalist() {
    return `<datalist id="sp-list">${salespeople
      .map((s) => `<option value="${escapeHtml(s.name)}"></option>`)
      .join("")}</datalist>`;
  }

  // Service is a free-text field (saved exactly as typed / as it came from
  // the uploaded file), but this datalist suggests existing Services master
  // names so staff naturally reuse "Internet" instead of accidentally
  // creating near-duplicates like "internet" / "Internet ".
  function serviceNameDatalist() {
    return `<datalist id="service-name-list">${services
      .map((s) => `<option value="${escapeHtml(s.name)}"></option>`)
      .join("")}</datalist>`;
  }

  function drawReviewStep() {
    container.innerHTML = `
      <div class="screen">
        ${screenHeaderHtml("Step 3: Review &amp; Confirm")}
        <p class="muted">Check and edit values for each customer block. Blocks that need a look — data missing from one file, a mismatch between the two files, or a possible duplicate customer — are marked "Please review" and listed first. Only entries you want saved need to stay checked. Use "Remove" on a block to drop it from this list entirely (it won't be saved and won't be counted); "Reset" above discards the whole upload and starts over.</p>
        ${salespersonDatalist()}
        ${serviceNameDatalist()}
        <div class="btn-row">
          <button class="btn" id="add-block-btn">+ Add Customer Manually</button>
          <button class="btn primary" id="confirm-btn">Confirm Selected Customers</button>
          <button class="btn" id="back-to-mapping-btn">Back (Column Mapping)</button>
        </div>
        ${busyMsg ? `<div class="alert info">${escapeHtml(busyMsg)}</div>` : ""}
        <div id="blocks-wrap">${blocks.map((b) => blockCardHtml(b)).join("") || `<p class="muted">No customer data was loaded.</p>`}</div>
      </div>
    `;
    wireResetBtn();
    wireReviewEvents();
  }

  function blockCardHtml(b) {
    return `
    <div class="card block-card ${b.excluded ? "excluded" : ""} ${b.needsReview ? "needs-review" : ""}" data-block="${b.id}">
      <div class="block-head">
        <label class="checkbox-inline"><input type="checkbox" class="b-include" ${b.excluded ? "" : "checked"} /> Include this customer</label>
        <button class="btn small danger b-remove">Remove</button>
        ${
          b.needsReview
            ? `<span class="badge error">Please review: ${escapeHtml((b.reviewReasons || []).join(", "))}</span>`
            : ""
        }
        ${
          b.duplicate.suspect
            ? `<span class="badge warn">Possible duplicate: ${escapeHtml(b.duplicate.reason)}${
                b.duplicate.matchedCustomerLabel ? " &rarr; existing: " + escapeHtml(b.duplicate.matchedCustomerLabel) : ""
              }</span>`
            : ""
        }
      </div>
      ${
        b.duplicate.suspect
          ? `<label class="checkbox-inline"><input type="checkbox" class="b-link-existing" ${b.linkToExisting ? "checked" : ""} /> Treat as the same as the existing customer (don't create a new one)</label>`
          : ""
      }
      <div class="grid4">
        <label>Name<input type="text" class="b-name" value="${escapeHtml(b.name)}" /></label>
        <label>Phone<input type="text" class="b-phone" value="${escapeHtml(b.phone)}" /></label>
        <label>Order Date<input type="date" class="b-date" value="${b.orderDate || ""}" /></label>
        <label>Salesperson<input type="text" class="b-sales" list="sp-list" value="${escapeHtml(b.salesperson)}" /></label>
        <label>Address<input type="text" class="b-address" value="${escapeHtml(b.address || "")}" /></label>
      </div>
      ${
        b.memo
          ? `<label>Memo / Notes<textarea class="b-memo" rows="2" style="width:100%">${escapeHtml(b.memo)}</textarea></label>`
          : ""
      }
      <label class="checkbox-inline"><input type="checkbox" class="b-followup-needed" ${b.followUpNeeded ? "checked" : ""} /> Follow-up needed</label>
      ${
        b.followUpNeeded
          ? `<div class="inline-form">
              <label>Follow-up Date<input type="date" class="b-followup-date" value="${b.followUpDate || ""}" /></label>
              <button type="button" class="btn small b-followup-1w">+1 Week</button>
              <button type="button" class="btn small b-followup-2w">+2 Weeks</button>
            </div>
            <div class="inline-form">
              <textarea class="b-followup-reason" placeholder="Why is a follow-up needed? (e.g. confirm credit was applied, call back about...)" rows="2" style="flex:1">${escapeHtml(b.followUpReason || "")}</textarea>
            </div>`
          : ""
      }
      <table class="data-table small">
        <thead><tr><th>Include</th><th>Service</th><th>Provider</th><th>Account #</th><th>Units</th><th>Units Installed</th><th>Expected Commission</th><th></th></tr></thead>
        <tbody>
          ${b.lines
            .map(
              (l) => `
            <tr data-line="${l.id}">
              <td><input type="checkbox" class="l-include" ${l.excluded ? "" : "checked"} /></td>
              <td><input type="text" class="l-service-name" list="service-name-list" value="${escapeHtml(l.serviceName || "")}" placeholder="Service / Product" />
              </td>
              <td><select class="l-provider">
                <option value="">(Select)</option>
                ${providers.map((p) => `<option value="${p.id}" ${l.providerId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
              </select>
              ${
                l.providerRaw && !l.providerMatched
                  ? `<div class="hint-error">No match for "${escapeHtml(l.providerRaw)}" — select manually or add it in Master Data</div>`
                  : ""
              }
              ${
                !l.providerId
                  ? `<div class="hint-error">No provider selected</div>`
                  : ""
              }
              </td>
              <td><input type="text" class="l-account" value="${escapeHtml(l.accountNumber)}" /></td>
              <td><input type="number" class="l-units" value="${l.units ?? ""}" style="width:60px" /></td>
              <td><input type="number" class="l-units-installed" value="${l.unitsInstalled ?? ""}" style="width:60px" /></td>
              <td><input type="number" step="0.01" class="l-commission" value="${l.expectedCommission ?? ""}" /></td>
              <td><button class="btn small danger l-remove">Delete</button></td>
            </tr>`
            )
            .join("")}
        </tbody>
      </table>
      <button class="btn small b-add-line">+ Add Service Line</button>
    </div>`;
  }

  function readBlockFromCard(card, b) {
    b.excluded = !card.querySelector(".b-include").checked;
    b.name = card.querySelector(".b-name").value.trim();
    b.phone = card.querySelector(".b-phone").value.trim();
    b.orderDate = card.querySelector(".b-date").value || null;
    b.salesperson = card.querySelector(".b-sales").value.trim();
    const addressEl = card.querySelector(".b-address");
    if (addressEl) b.address = addressEl.value.trim();
    const memoEl = card.querySelector(".b-memo");
    if (memoEl) b.memo = memoEl.value.trim();
    const linkEl = card.querySelector(".b-link-existing");
    if (linkEl) b.linkToExisting = linkEl.checked;
    b.followUpNeeded = card.querySelector(".b-followup-needed").checked;
    const followUpDateEl = card.querySelector(".b-followup-date");
    if (followUpDateEl) b.followUpDate = followUpDateEl.value || null;
    const followUpReasonEl = card.querySelector(".b-followup-reason");
    if (followUpReasonEl) b.followUpReason = followUpReasonEl.value.trim();
    card.querySelectorAll("tr[data-line]").forEach((tr) => {
      const line = b.lines.find((l) => l.id === tr.dataset.line);
      if (!line) return;
      line.excluded = !tr.querySelector(".l-include").checked;
      line.serviceName = tr.querySelector(".l-service-name").value.trim();
      line.providerId = tr.querySelector(".l-provider").value || null;
      line.accountNumber = normalizeAccountNumber(tr.querySelector(".l-account").value);
      const unitsVal = tr.querySelector(".l-units").value;
      line.units = unitsVal === "" ? null : Number(unitsVal);
      const unitsInstalledVal = tr.querySelector(".l-units-installed").value;
      line.unitsInstalled = unitsInstalledVal === "" ? null : Number(unitsInstalledVal);
      const commVal = tr.querySelector(".l-commission").value;
      line.expectedCommission = commVal === "" ? null : Number(commVal);
    });
  }

  function wireReviewEvents() {
    container.querySelector("#add-block-btn").addEventListener("click", () => {
      blocks.push({
        id: crypto.randomUUID(),
        name: "",
        phone: "",
        address: "",
        orderDate: null,
        salesperson: "",
        memo: "",
        excluded: false,
        duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
        linkToExisting: false,
        sourceOrderId: null,
        orderNumber: null,
        category: null,
        needsReview: false,
        reviewReasons: [],
        followUpNeeded: false,
        followUpDate: null,
        followUpReason: "",
        lines: [
          {
            id: crypto.randomUUID(),
            excluded: false,
            serviceName: "",
            providerId: null,
            providerMatched: true,
            providerRaw: null,
            accountNumberRaw: "",
            accountNumber: "",
            expectedCommission: null,
            units: null,
            unitsInstalled: null,
            mobileLinesOrdered: null,
            mobileLinesInstalled: null,
            planName: null,
          },
        ],
      });
      drawReviewStep();
    });

    container.querySelector("#back-to-mapping-btn").addEventListener("click", () => {
      step = mode === "two" ? "mappingTwo" : "mapping";
      if (mode === "two") drawMappingTwoStep();
      else drawMappingStep();
    });

    container.querySelectorAll(".block-card").forEach((card) => {
      const b = blocks.find((x) => x.id === card.dataset.block);
      card.querySelector(".b-remove").addEventListener("click", () => {
        blocks = blocks.filter((x) => x.id !== b.id);
        drawReviewStep();
      });
      card.querySelector(".b-add-line").addEventListener("click", () => {
        readBlockFromCard(card, b);
        b.lines.push({
          id: crypto.randomUUID(),
          excluded: false,
          serviceName: "",
          providerId: null,
          providerMatched: true,
          providerRaw: null,
          accountNumberRaw: "",
          accountNumber: "",
          expectedCommission: null,
          units: null,
          unitsInstalled: null,
          mobileLinesOrdered: null,
          mobileLinesInstalled: null,
          planName: null,
        });
        drawReviewStep();
      });
      card.querySelectorAll(".l-remove").forEach((btn) => {
        btn.addEventListener("click", () => {
          readBlockFromCard(card, b);
          const lineId = btn.closest("tr").dataset.line;
          b.lines = b.lines.filter((l) => l.id !== lineId);
          drawReviewStep();
        });
      });
      card.querySelector(".b-followup-needed").addEventListener("change", (e) => {
        readBlockFromCard(card, b);
        b.followUpNeeded = e.target.checked;
        if (b.followUpNeeded && !b.followUpDate) {
          b.followUpDate = addDays(b.orderDate || todayStr(), 7);
        }
        drawReviewStep();
      });
      card.querySelector(".b-followup-1w")?.addEventListener("click", () => {
        const dateInput = card.querySelector(".b-followup-date");
        dateInput.value = addDays(dateInput.value || b.orderDate || todayStr(), 7);
      });
      card.querySelector(".b-followup-2w")?.addEventListener("click", () => {
        const dateInput = card.querySelector(".b-followup-date");
        dateInput.value = addDays(dateInput.value || b.orderDate || todayStr(), 14);
      });
    });

    container.querySelector("#confirm-btn").addEventListener("click", async () => {
      container.querySelectorAll(".block-card").forEach((card) => {
        const b = blocks.find((x) => x.id === card.dataset.block);
        readBlockFromCard(card, b);
      });
      await confirmBlocks();
    });
  }

  // Resolves a Service master row id for the given free-text name (as typed
  // on the review screen / taken verbatim from the uploaded file), reusing
  // an existing master row (case-insensitive exact match) or auto-creating
  // one if it doesn't exist yet. This is what lets Service be free text on
  // the review screen while still ending up linked to a real Services
  // master row for reporting. If the insert fails (e.g. a non-admin
  // account, since Services master rows are admin-only to insert per RLS),
  // the line is still saved with service_id left null rather than failing
  // the whole customer/order save — the raw text is preserved in plan_name.
  async function getOrCreateServiceId(rawName) {
    const name = (rawName || "").trim();
    if (!name) return null;
    const existing = services.find((s) => s.name.trim().toLowerCase() === name.toLowerCase());
    if (existing) return existing.id;
    const { data: newSvc, error } = await supabase.from("services").insert({ name }).select().single();
    if (error) {
      const { data: retryData } = await supabase.from("services").select("id, name").ilike("name", name);
      const retryMatch = (retryData || []).find((s) => s.name.trim().toLowerCase() === name.toLowerCase());
      return retryMatch ? retryMatch.id : null;
    }
    services.push(newSvc);
    return newSvc.id;
  }

  async function confirmBlocks() {
    const toConfirm = blocks.filter((b) => !b.excluded);
    if (toConfirm.length === 0) {
      alert("No customers are included.");
      return;
    }
    busyMsg = "Saving...";
    drawReviewStep();

    let successCount = 0;
    let skippedBlockCount = 0;
    let skippedLineCount = 0;
    const errors = [];
    const followUpErrors = [];

    for (const b of toConfirm) {
      try {
        if (!b.name) throw new Error("Name is required.");

        // Exact-duplicate check: an Account Number + Customer Full Name +
        // Service (Product/Package) combination that already exists in the
        // system is skipped instead of saved again.
        const candidateLines = b.lines.filter((l) => !l.excluded);
        const nonDupLines = [];
        for (const l of candidateLines) {
          // Resolve (and cache) the Service master id up front so both the
          // duplicate check and the later insert use the same id instead of
          // creating/looking it up twice.
          l._resolvedServiceId = await getOrCreateServiceId(l.serviceName);
          const isDup = await findExactDuplicateLine({
            accountNumber: l.accountNumber,
            customerName: b.name,
            serviceId: l._resolvedServiceId,
          });
          if (isDup) {
            skippedLineCount++;
            continue;
          }
          nonDupLines.push(l);
        }
        if (candidateLines.length > 0 && nonDupLines.length === 0) {
          // Every included line on this block is already saved — skip the
          // whole block rather than creating an empty duplicate order.
          skippedBlockCount++;
          b.savedOk = true;
          continue;
        }

        // Salesperson: reuse an existing name (case-insensitive match), otherwise create a new one
        let salespersonId = null;
        if (b.salesperson) {
          const existing = salespeople.find(
            (s) => s.name.trim().toLowerCase() === b.salesperson.trim().toLowerCase()
          );
          if (existing) {
            salespersonId = existing.id;
          } else {
            const { data: newSp, error: spErr } = await supabase
              .from("salespeople")
              .insert({ name: b.salesperson })
              .select()
              .single();
            if (spErr) throw spErr;
            salespeople.push(newSp);
            salespersonId = newSp.id;
          }
        }

        // Customer: reuse the existing customer if flagged as a duplicate and confirmed by staff; otherwise create new
        let customerId = null;
        let customerWasNewlyCreated = false;
        if (b.duplicate.suspect && b.linkToExisting && b.duplicate.matchedCustomerId) {
          customerId = b.duplicate.matchedCustomerId;
        } else {
          const { data: newCust, error: custErr } = await supabase
            .from("customers")
            .insert({ name: b.name, phone: b.phone, address: b.address || null, registered_by: ctx.profile.id })
            .select()
            .single();
          if (custErr) throw custErr;
          customerId = newCust.id;
          customerWasNewlyCreated = true;
        }

        let newOrder = null;
        try {
          // Order
          const { data: orderData, error: orderErr } = await supabase
            .from("orders")
            .insert({
              customer_id: customerId,
              salesperson_id: salespersonId,
              order_date: b.orderDate,
              pipeline_stage: "confirmed",
              memo: b.memo || null,
              created_by: ctx.profile.id,
              source_order_id: b.sourceOrderId ?? null,
              order_number: b.orderNumber ?? null,
              category: b.category ?? null,
            })
            .select()
            .single();
          if (orderErr) throw orderErr;
          newOrder = orderData;

          const lineRows = [];
          for (const l of nonDupLines) {
            const serviceId = l._resolvedServiceId ?? (await getOrCreateServiceId(l.serviceName));
            lineRows.push({
              order_id: newOrder.id,
              service_id: serviceId,
              provider_id: l.providerId || null,
              plan_name: l.planName ?? l.serviceName ?? null,
              account_number: l.accountNumber || null,
              expected_commission: l.expectedCommission,
              units: l.units ?? null,
              units_installed: l.unitsInstalled ?? null,
              mobile_lines_ordered: l.mobileLinesOrdered ?? null,
              mobile_lines_installed: l.mobileLinesInstalled ?? null,
              status: "pending",
            });
          }
          if (lineRows.length > 0) {
            const { error: lineErr } = await supabase.from("order_service_lines").insert(lineRows);
            if (lineErr) throw lineErr;
          }
        } catch (innerErr) {
          // If the order (or its lines) fails to save after the customer was
          // newly created, roll the customer back too -- otherwise a
          // transient failure here (e.g. a dropped connection) leaves an
          // orphan customer with no order behind, invisible to the
          // Customers/Orders screen's old orders-only view.
          if (customerWasNewlyCreated) {
            await supabase.from("customers").delete().eq("id", customerId);
          }
          throw innerErr;
        }

        // Follow-up (optional): the order/lines above are already saved at
        // this point, so a follow-up save failure is reported alongside the
        // summary instead of rolling back the whole customer/order.
        if (b.followUpNeeded && b.followUpDate) {
          const { error: followUpErr } = await supabase.from("follow_ups").insert({
            customer_id: customerId,
            order_id: newOrder?.id || null,
            due_date: b.followUpDate,
            reason: b.followUpReason || null,
            created_by: ctx.profile?.id || null,
          });
          if (followUpErr) {
            followUpErrors.push(`${b.name || "(no name)"}: ${followUpErr.message}`);
          }
        }

        successCount++;
        b.savedOk = true;
      } catch (err) {
        errors.push(`${b.name || "(no name)"}: ${err.message || err}`);
      }
    }

    blocks = blocks.filter((b) => !b.savedOk);
    busyMsg =
      `${successCount} saved.` +
      (skippedBlockCount ? ` ${skippedBlockCount} customer(s) skipped (already saved — same Account Number + Name + Service).` : "") +
      (skippedLineCount ? ` ${skippedLineCount} line(s) skipped as duplicates.` : "") +
      (errors.length ? ` ${errors.length} failed: ` + errors.join(" / ") : "") +
      (followUpErrors.length
        ? ` ${followUpErrors.length} follow-up(s) failed to save (customer/order were saved fine): ` + followUpErrors.join(" / ")
        : "");
    drawReviewStep();
  }

  await loadMasters();
  drawUploadStep();
}
