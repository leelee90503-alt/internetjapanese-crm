import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";
import { checkDuplicateSuspect } from "./duplicateCheck.js";
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
];

// ---- Two-file merge mode: field definitions per source file. ----
// Order matters for auto-detection: more specific labels (e.g. "units_installed",
// "package_group") must be listed BEFORE the more generic ones they overlap with
// ("units", "package") so the substring-matching fallback pass doesn't let the
// generic field steal a column meant for the specific one.
const FILE_A_FIELDS = [
  { key: "order_id", label: "Order Id", guesses: ["order id", "orderid"], required: true },
  { key: "name", label: "Customer Full Name", guesses: ["customer full name", "full name", "customer name"] },
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
  let busyMsg = "";

  async function loadMasters() {
    const [s, p, sp] = await Promise.all([
      supabase.from("services").select("*").eq("is_active", true).order("sort_order"),
      supabase.from("providers").select("*").eq("is_active", true).order("name"),
      supabase.from("salespeople").select("*").eq("is_active", true).order("name"),
    ]);
    services = s.data || [];
    providers = p.data || [];
    salespeople = sp.data || [];
  }

  function drawUploadStep() {
    container.innerHTML = `
      <div class="screen">
        <h2>Customer Excel Upload</h2>
        <p class="muted">Upload an Excel (.xlsx/.csv) file, or two files that share an Order Id column and need to be matched and merged first. Everything goes through column mapping and a review/confirm step before being saved as real customer and order data.</p>
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
        <h2>Step 2: Column Mapping &amp; Validation</h2>
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
              ${providers.map((p) => `<option value="${p.id}">${escapeHtml(p.name)}</option>`).join("")}
            </select>
          </label>
          <div class="btn-row">
            <button class="btn" id="back-btn">Re-upload</button>
            <button class="btn primary" id="next-btn">Next: Review &amp; Confirm</button>
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
      defaultProviderId = container.querySelector("#default-provider").value;
      await buildBlocks();
      step = "review";
      drawReviewStep();
    });
  }

  function drawMappingTwoStep() {
    container.innerHTML = `
      <div class="screen">
        <h2>Step 2: Column Mapping (Two Files)</h2>
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
    for (const row of dataRows) {
      const name = String(cell(row, "name") ?? "").trim();
      const isNewBlock = name !== "";
      if (isNewBlock) {
        current = {
          id: crypto.randomUUID(),
          name,
          phone: String(cell(row, "phone") ?? "").trim(),
          orderDate: normalizeDate(cell(row, "order_date")),
          salesperson: String(cell(row, "salesperson") ?? "").trim(),
          excluded: false,
          duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
          linkToExisting: false,
          sourceOrderId: null,
          orderNumber: null,
          category: null,
          lines: [],
        };
        blocks.push(current);
      }
      if (!current) continue; // skip rows with no name if they appear before any block starts (formatting error)
      const providerRaw = String(cell(row, "provider") ?? "").trim();
      const provider = matchMaster(providers, providerRaw) || (defaultProviderId || null);
      const service = matchMaster(services, String(cell(row, "service") ?? "").trim());
      current.lines.push({
        id: crypto.randomUUID(),
        excluded: false,
        serviceId: service,
        providerId: provider,
        accountNumberRaw: String(cell(row, "account_number") ?? "").trim(),
        accountNumber: normalizeAccountNumber(cell(row, "account_number")),
        expectedCommission: Number(cell(row, "expected_commission")) || null,
        units: null,
        unitsInstalled: null,
        mobileLinesOrdered: null,
        mobileLinesInstalled: null,
        planName: null,
      });
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
  }

  // ---- two-file merge ----
  function buildMergedRows() {
    const aIndex = new Map();
    fileADataRows.forEach((row) => {
      const oid = normText(cellByMapping(row, fileAMapping, "order_id"));
      if (!oid) return;
      aIndex.set(oid, row);
    });
    const bIndex = new Map();
    fileBDataRows.forEach((row) => {
      const oid = normText(cellByMapping(row, fileBMapping, "order_id"));
      if (!oid) return;
      bIndex.set(oid, row);
    });

    const allIds = Array.from(new Set([...aIndex.keys(), ...bIndex.keys()])).sort();
    mergedRows = allIds.map((oid) => {
      const aRow = aIndex.get(oid) || null;
      const bRow = bIndex.get(oid) || null;
      const getA = (key) => (aRow ? cellByMapping(aRow, fileAMapping, key) : "");
      const getB = (key) => (bRow ? cellByMapping(bRow, fileBMapping, key) : "");

      const aName = getA("name");
      const bName = getB("name");
      const aPhone = getA("phone");
      const bPhone = getB("phone");
      const aAccount = getA("account_number");
      const bAccount = getB("account_number");
      const aUnits = getA("units");
      const bUnits = getB("units");
      const aUnitsInstalled = getA("units_installed");
      const bUnitsInstalled = getB("units_installed");

      const mismatches = {};
      if (normText(aName) && normText(bName) && normText(aName) !== normText(bName)) {
        mismatches.name = { a: aName, b: bName };
      }
      if (normalizePhone(aPhone) && normalizePhone(bPhone) && normalizePhone(aPhone) !== normalizePhone(bPhone)) {
        mismatches.phone = { a: aPhone, b: bPhone };
      }
      if (
        normalizeAccountNumber(aAccount) &&
        normalizeAccountNumber(bAccount) &&
        normalizeAccountNumber(aAccount) !== normalizeAccountNumber(bAccount)
      ) {
        mismatches.accountNumber = { a: aAccount, b: bAccount };
      }
      if (normNum(aUnits) !== null && normNum(bUnits) !== null && normNum(aUnits) !== normNum(bUnits)) {
        mismatches.units = { a: aUnits, b: bUnits };
      }
      if (
        normNum(aUnitsInstalled) !== null &&
        normNum(bUnitsInstalled) !== null &&
        normNum(aUnitsInstalled) !== normNum(bUnitsInstalled)
      ) {
        mismatches.unitsInstalled = { a: aUnitsInstalled, b: bUnitsInstalled };
      }

      return {
        id: crypto.randomUUID(),
        orderId: oid,
        onlyInA: !bRow,
        onlyInB: !aRow,
        excluded: false,
        mismatches,
        name: aName || bName,
        phone: aPhone || bPhone,
        address: joinAddress(
          getA("address") || getB("address"),
          getA("city") || getB("city"),
          getA("state") || getB("state"),
          getA("zip") || getB("zip")
        ),
        orderDate: normalizeDate(getB("order_date")),
        accountNumber: normalizeAccountNumber(aAccount || bAccount),
        units: normNum(aUnits) !== null ? normNum(aUnits) : normNum(bUnits),
        unitsInstalled: normNum(aUnitsInstalled) !== null ? normNum(aUnitsInstalled) : normNum(bUnitsInstalled),
        product: getA("product"),
        package: getA("package"),
        packageGroup: getA("package_group"),
        mobileLinesOrdered: normNum(getA("mobile_lines_ordered")),
        mobileLinesInstalled: normNum(getA("mobile_lines_installed")),
        orderNumber: getB("order_number"),
        category: getB("category"),
        salesperson: getB("salesperson"),
      };
    });
  }

  function drawMergeReviewStep() {
    const includedCount = mergedRows.filter((r) => !r.excluded).length;
    const matchedBothCount = mergedRows.filter((r) => !r.onlyInA && !r.onlyInB).length;
    const onlyACount = mergedRows.filter((r) => r.onlyInA).length;
    const onlyBCount = mergedRows.filter((r) => r.onlyInB).length;
    container.innerHTML = `
      <div class="screen">
        <h2>Step 3: Merge Result Review</h2>
        <p class="muted">Matched ${mergedRows.length} order(s) by Order Id — ${matchedBothCount} found in both files, ${onlyACount} only in the Service Info file (missing Order Date), ${onlyBCount} only in the Order Info file (missing Product/Package). Fix Units / Units Installed here if they differ between the two files; missing Order Date and Service/Provider selections can still be filled in on the next screen.</p>
        <div class="btn-row">
          <button class="btn primary" id="merge-next-btn">Next: Review &amp; Confirm (${includedCount})</button>
          <button class="btn" id="merge-back-btn">Back (Column Mapping)</button>
        </div>
        <table class="data-table small">
          <thead>
            <tr><th>Include</th><th>Order Id</th><th>Customer</th><th>Status</th><th>Flags</th><th>Units</th><th>Units Installed</th></tr>
          </thead>
          <tbody>
            ${
              mergedRows
                .map(
                  (r) => `
              <tr data-mrow="${r.id}" class="${r.excluded ? "excluded" : ""}">
                <td><input type="checkbox" class="mr-include" ${r.excluded ? "" : "checked"} /></td>
                <td>${escapeHtml(r.orderId)}</td>
                <td>${escapeHtml(r.name || "-")}</td>
                <td>
                  ${r.onlyInA ? `<span class="badge warn">Only in Service Info file</span>` : ""}
                  ${r.onlyInB ? `<span class="badge warn">Only in Order Info file</span>` : ""}
                  ${!r.onlyInA && !r.onlyInB ? `<span class="badge ok">Matched</span>` : ""}
                </td>
                <td>
                  ${r.mismatches.name ? `<span class="badge error">Name differs</span>` : ""}
                  ${r.mismatches.phone ? `<span class="badge error">Phone differs</span>` : ""}
                  ${r.mismatches.accountNumber ? `<span class="badge error">Account # differs</span>` : ""}
                  ${r.mismatches.units ? `<span class="badge error">Units: ${escapeHtml(String(r.mismatches.units.a))} / ${escapeHtml(String(r.mismatches.units.b))}</span>` : ""}
                  ${r.mismatches.unitsInstalled ? `<span class="badge error">Units Installed: ${escapeHtml(String(r.mismatches.unitsInstalled.a))} / ${escapeHtml(String(r.mismatches.unitsInstalled.b))}</span>` : ""}
                  ${Object.keys(r.mismatches).length === 0 && !r.onlyInA && !r.onlyInB ? `<span class="muted">-</span>` : ""}
                </td>
                <td><input type="number" class="mr-units" value="${r.units ?? ""}" style="width:70px" /></td>
                <td><input type="number" class="mr-units-installed" value="${r.unitsInstalled ?? ""}" style="width:70px" /></td>
              </tr>`
                )
                .join("") || `<tr><td colspan="7" class="muted">No rows to merge.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    `;
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
      tr.querySelector(".mr-units").addEventListener("change", (e) => {
        r.units = e.target.value === "" ? null : Number(e.target.value);
      });
      tr.querySelector(".mr-units-installed").addEventListener("change", (e) => {
        r.unitsInstalled = e.target.value === "" ? null : Number(e.target.value);
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
      const planName =
        [r.product, r.package, r.packageGroup].filter((v) => v && String(v).trim() !== "").join(" / ") || null;
      const serviceId =
        matchMaster(services, r.packageGroup) || matchMaster(services, r.package) || matchMaster(services, r.product);

      blocks.push({
        id: crypto.randomUUID(),
        name: r.name || "",
        phone: r.phone || "",
        orderDate: r.orderDate,
        salesperson: r.salesperson || "",
        excluded: false,
        duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
        linkToExisting: false,
        sourceOrderId: r.orderId,
        orderNumber: r.orderNumber || null,
        category: r.category || null,
        lines: [
          {
            id: crypto.randomUUID(),
            excluded: false,
            serviceId,
            providerId: defaultProviderId || null,
            accountNumberRaw: r.accountNumber || "",
            accountNumber: r.accountNumber || "",
            expectedCommission: null,
            units: r.units,
            unitsInstalled: r.unitsInstalled,
            mobileLinesOrdered: r.mobileLinesOrdered,
            mobileLinesInstalled: r.mobileLinesInstalled,
            planName,
          },
        ],
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
  }

  function matchMaster(list, rawValue) {
    if (!rawValue) return null;
    const norm = rawValue.trim().toLowerCase();
    const hit = list.find(
      (m) =>
        m.name.toLowerCase() === norm ||
        (m.display_name && m.display_name.toLowerCase() === norm) ||
        m.name.toLowerCase().includes(norm) ||
        norm.includes(m.name.toLowerCase())
    );
    return hit ? hit.id : null;
  }

  function salespersonDatalist() {
    return `<datalist id="sp-list">${salespeople
      .map((s) => `<option value="${escapeHtml(s.name)}"></option>`)
      .join("")}</datalist>`;
  }

  function drawReviewStep() {
    container.innerHTML = `
      <div class="screen">
        <h2>Step 3: Review &amp; Confirm</h2>
        <p class="muted">Check and edit values for each customer block. If a "Possible Duplicate" badge appears, compare it yourself and decide whether it's really a duplicate. Only entries you want saved need to stay checked.</p>
        ${salespersonDatalist()}
        <div class="btn-row">
          <button class="btn" id="add-block-btn">+ Add Customer Manually</button>
          <button class="btn primary" id="confirm-btn">Confirm Selected Customers</button>
          <button class="btn" id="back-to-mapping-btn">Back (Column Mapping)</button>
        </div>
        ${busyMsg ? `<div class="alert info">${escapeHtml(busyMsg)}</div>` : ""}
        <div id="blocks-wrap">${blocks.map((b) => blockCardHtml(b)).join("") || `<p class="muted">No customer data was loaded.</p>`}</div>
      </div>
    `;
    wireReviewEvents();
  }

  function blockCardHtml(b) {
    return `
    <div class="card block-card ${b.excluded ? "excluded" : ""}" data-block="${b.id}">
      <div class="block-head">
        <label class="checkbox-inline"><input type="checkbox" class="b-include" ${b.excluded ? "" : "checked"} /> Include this customer</label>
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
      </div>
      <table class="data-table small">
        <thead><tr><th>Include</th><th>Service</th><th>Provider</th><th>Account #</th><th>Units</th><th>Units Installed</th><th>Expected Commission</th><th></th></tr></thead>
        <tbody>
          ${b.lines
            .map(
              (l) => `
            <tr data-line="${l.id}">
              <td><input type="checkbox" class="l-include" ${l.excluded ? "" : "checked"} /></td>
              <td><select class="l-service">
                <option value="">(Select)</option>
                ${services.map((s) => `<option value="${s.id}" ${l.serviceId === s.id ? "selected" : ""}>${escapeHtml(s.name)}</option>`).join("")}
              </select></td>
              <td><select class="l-provider">
                <option value="">(Select)</option>
                ${providers.map((p) => `<option value="${p.id}" ${l.providerId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
              </select></td>
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
    const linkEl = card.querySelector(".b-link-existing");
    if (linkEl) b.linkToExisting = linkEl.checked;
    card.querySelectorAll("tr[data-line]").forEach((tr) => {
      const line = b.lines.find((l) => l.id === tr.dataset.line);
      if (!line) return;
      line.excluded = !tr.querySelector(".l-include").checked;
      line.serviceId = tr.querySelector(".l-service").value || null;
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
        orderDate: null,
        salesperson: "",
        excluded: false,
        duplicate: { suspect: false, reason: "", matchedCustomerId: null, matchedCustomerLabel: null },
        linkToExisting: false,
        sourceOrderId: null,
        orderNumber: null,
        category: null,
        lines: [
          {
            id: crypto.randomUUID(),
            excluded: false,
            serviceId: null,
            providerId: null,
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
      card.querySelector(".b-add-line").addEventListener("click", () => {
        readBlockFromCard(card, b);
        b.lines.push({
          id: crypto.randomUUID(),
          excluded: false,
          serviceId: null,
          providerId: null,
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
    });

    container.querySelector("#confirm-btn").addEventListener("click", async () => {
      container.querySelectorAll(".block-card").forEach((card) => {
        const b = blocks.find((x) => x.id === card.dataset.block);
        readBlockFromCard(card, b);
      });
      await confirmBlocks();
    });
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
    const errors = [];

    for (const b of toConfirm) {
      try {
        if (!b.name) throw new Error("Name is required.");

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
        if (b.duplicate.suspect && b.linkToExisting && b.duplicate.matchedCustomerId) {
          customerId = b.duplicate.matchedCustomerId;
        } else {
          const { data: newCust, error: custErr } = await supabase
            .from("customers")
            .insert({ name: b.name, phone: b.phone, registered_by: ctx.profile.id })
            .select()
            .single();
          if (custErr) throw custErr;
          customerId = newCust.id;
        }

        // Order
        const { data: newOrder, error: orderErr } = await supabase
          .from("orders")
          .insert({
            customer_id: customerId,
            salesperson_id: salespersonId,
            order_date: b.orderDate,
            pipeline_stage: "confirmed",
            created_by: ctx.profile.id,
            source_order_id: b.sourceOrderId ?? null,
            order_number: b.orderNumber ?? null,
            category: b.category ?? null,
          })
          .select()
          .single();
        if (orderErr) throw orderErr;

        const lineRows = b.lines
          .filter((l) => !l.excluded)
          .map((l) => ({
            order_id: newOrder.id,
            service_id: l.serviceId || null,
            provider_id: l.providerId || null,
            plan_name: l.planName ?? null,
            account_number: l.accountNumber || null,
            expected_commission: l.expectedCommission,
            units: l.units ?? null,
            units_installed: l.unitsInstalled ?? null,
            mobile_lines_ordered: l.mobileLinesOrdered ?? null,
            mobile_lines_installed: l.mobileLinesInstalled ?? null,
            status: "pending",
          }));
        if (lineRows.length > 0) {
          const { error: lineErr } = await supabase.from("order_service_lines").insert(lineRows);
          if (lineErr) throw lineErr;
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
      (errors.length ? ` ${errors.length} failed: ` + errors.join(" / ") : "");
    drawReviewStep();
  }

  drawUploadStep();
}
