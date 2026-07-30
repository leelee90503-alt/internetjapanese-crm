import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";
import { checkDuplicateSuspect } from "./duplicateCheck.js";
import {
  normalizeAccountNumber,
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

export async function renderCustomerUpload(container, ctx) {
  let step = "upload"; // upload -> mapping -> review
  let workbook = null;
  let sheetNames = [];
  let selectedSheet = "";
  let headers = [];
  let dataRows = [];
  let mapping = {}; // field key -> column index (or -1)
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
        <p class="muted">Upload an Excel (.xlsx/.csv) file. It will go through column mapping and a review/confirm step before being saved as real customer and order data.</p>
        <div class="card">
          <input type="file" id="file-input" accept=".xlsx,.xls,.csv" />
        </div>
      </div>
    `;
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

  function cell(row, fieldKey) {
    const idx = mapping[fieldKey];
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
        <thead><tr><th>Include</th><th>Service</th><th>Provider</th><th>Account #</th><th>Expected Commission</th><th></th></tr></thead>
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
        lines: [
          {
            id: crypto.randomUUID(),
            excluded: false,
            serviceId: null,
            providerId: null,
            accountNumberRaw: "",
            accountNumber: "",
            expectedCommission: null,
          },
        ],
      });
      drawReviewStep();
    });

    container.querySelector("#back-to-mapping-btn").addEventListener("click", () => {
      step = "mapping";
      drawMappingStep();
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
            account_number: l.accountNumber || null,
            expected_commission: l.expectedCommission,
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
