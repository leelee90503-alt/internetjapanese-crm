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
  { key: "name", label: "이름", guesses: ["name", "이름", "고객명", "customer"] },
  { key: "phone", label: "전화번호", guesses: ["phone", "전화", "연락처", "tel"] },
  {
    key: "account_number",
    label: "계정번호",
    guesses: ["account", "acc", "계정번호", "계정"],
  },
  { key: "order_date", label: "오더날짜", guesses: ["date", "날짜", "가입일", "order date"] },
  { key: "provider", label: "프로바이더", guesses: ["provider", "프로바이더", "carrier"] },
  { key: "service", label: "서비스/상품", guesses: ["service", "plan", "서비스", "상품"] },
  {
    key: "salesperson",
    label: "세일즈 담당자",
    guesses: ["sales", "세일즈", "담당자", "sold by", "sale by"],
  },
  {
    key: "expected_commission",
    label: "예상 커미션",
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
        <h2>고객 정보 엑셀 업로드</h2>
        <p class="muted">엑셀(.xlsx/.csv) 파일을 업로드하면, 컬럼 매핑 → 검토·확인(컨펌) 단계를 거쳐 정식 고객·오더 데이터로 저장됩니다.</p>
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
        <h2>2단계: 컬럼 매핑 및 검증</h2>
        <div class="card">
          <label>시트 선택
            <select id="sheet-select">
              ${sheetNames.map((n) => `<option value="${escapeHtml(n)}" ${n === selectedSheet ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
            </select>
          </label>
          <p class="muted">감지된 컬럼(${headers.length}개) / 데이터 행 ${dataRows.length}개. 아래에서 각 시스템 필드가 어느 컬럼에 해당하는지 확인·수정하세요.</p>
          <table class="data-table">
            <thead><tr><th>시스템 필드</th><th>엑셀 컬럼</th></tr></thead>
            <tbody>
              ${FIELD_DEFS.map(
                (f) => `
                <tr>
                  <td>${f.label}</td>
                  <td>
                    <select data-field="${f.key}">
                      <option value="-1" ${mapping[f.key] === -1 || mapping[f.key] === undefined ? "selected" : ""}>(사용 안 함)</option>
                      ${headers
                        .map(
                          (h, i) =>
                            `<option value="${i}" ${mapping[f.key] === i ? "selected" : ""}>${escapeHtml(h) || "(제목없음 " + i + ")"}</option>`
                        )
                        .join("")}
                    </select>
                  </td>
                </tr>`
              ).join("")}
            </tbody>
          </table>
          <label>프로바이더 컬럼이 없거나 비어있는 행의 기본 프로바이더
            <select id="default-provider">
              <option value="">(지정 안 함)</option>
              ${providers.map((p) => `<option value="${p.id}">${escapeHtml(p.name)}</option>`).join("")}
            </select>
          </label>
          <div class="btn-row">
            <button class="btn" id="back-btn">다시 업로드</button>
            <button class="btn primary" id="next-btn">다음: 검토·확인</button>
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
      if (!current) continue; // 이름 없는 행이 첫 행일 경우(형식 오류) 건너뜀
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

    // 중복 의심 검사 (이름+전화번호+계정번호 조합, 오더날짜는 기준에서 제외)
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
        <h2>3단계: 검토·확인(컨펌)</h2>
        <p class="muted">고객 블록별로 값을 확인·수정하고, "중복 의심" 표시가 있으면 직접 비교해서 실제 중복인지 판단하세요. 문제 없는 항목만 선택되어 있으면 됩니다.</p>
        ${salespersonDatalist()}
        <div class="btn-row">
          <button class="btn" id="add-block-btn">+ 고객 수동 추가</button>
          <button class="btn primary" id="confirm-btn">선택한 고객 확인(컨펌)</button>
          <button class="btn" id="back-to-mapping-btn">뒤로(컬럼 매핑)</button>
        </div>
        ${busyMsg ? `<div class="alert info">${escapeHtml(busyMsg)}</div>` : ""}
        <div id="blocks-wrap">${blocks.map((b) => blockCardHtml(b)).join("") || `<p class="muted">불러온 고객 데이터가 없습니다.</p>`}</div>
      </div>
    `;
    wireReviewEvents();
  }

  function blockCardHtml(b) {
    return `
    <div class="card block-card ${b.excluded ? "excluded" : ""}" data-block="${b.id}">
      <div class="block-head">
        <label class="checkbox-inline"><input type="checkbox" class="b-include" ${b.excluded ? "" : "checked"} /> 이 고객 포함</label>
        ${
          b.duplicate.suspect
            ? `<span class="badge warn">중복 의심: ${escapeHtml(b.duplicate.reason)}${
                b.duplicate.matchedCustomerLabel ? " → 기존: " + escapeHtml(b.duplicate.matchedCustomerLabel) : ""
              }</span>`
            : ""
        }
      </div>
      ${
        b.duplicate.suspect
          ? `<label class="checkbox-inline"><input type="checkbox" class="b-link-existing" ${b.linkToExisting ? "checked" : ""} /> 기존 고객과 동일 건으로 처리(신규 등록하지 않음)</label>`
          : ""
      }
      <div class="grid4">
        <label>이름<input type="text" class="b-name" value="${escapeHtml(b.name)}" /></label>
        <label>전화번호<input type="text" class="b-phone" value="${escapeHtml(b.phone)}" /></label>
        <label>오더날짜<input type="date" class="b-date" value="${b.orderDate || ""}" /></label>
        <label>세일즈 담당자<input type="text" class="b-sales" list="sp-list" value="${escapeHtml(b.salesperson)}" /></label>
      </div>
      <table class="data-table small">
        <thead><tr><th>포함</th><th>서비스</th><th>프로바이더</th><th>계정번호</th><th>예상 커미션</th><th></th></tr></thead>
        <tbody>
          ${b.lines
            .map(
              (l) => `
            <tr data-line="${l.id}">
              <td><input type="checkbox" class="l-include" ${l.excluded ? "" : "checked"} /></td>
              <td><select class="l-service">
                <option value="">(선택)</option>
                ${services.map((s) => `<option value="${s.id}" ${l.serviceId === s.id ? "selected" : ""}>${escapeHtml(s.name)}</option>`).join("")}
              </select></td>
              <td><select class="l-provider">
                <option value="">(선택)</option>
                ${providers.map((p) => `<option value="${p.id}" ${l.providerId === p.id ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}
              </select></td>
              <td><input type="text" class="l-account" value="${escapeHtml(l.accountNumber)}" /></td>
              <td><input type="number" step="0.01" class="l-commission" value="${l.expectedCommission ?? ""}" /></td>
              <td><button class="btn small danger l-remove">삭제</button></td>
            </tr>`
            )
            .join("")}
        </tbody>
      </table>
      <button class="btn small b-add-line">+ 서비스 항목 추가</button>
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
      alert("포함된 고객이 없습니다.");
      return;
    }
    busyMsg = "저장 중입니다...";
    drawReviewStep();

    let successCount = 0;
    const errors = [];

    for (const b of toConfirm) {
      try {
        if (!b.name) throw new Error("이름이 비어 있습니다.");

        // 세일즈 담당자: 기존 이름과 대소문자 무관 일치 시 재사용, 없으면 새로 등록
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

        // 고객: 중복 의심 + "기존 고객과 동일" 체크 시 기존 고객 재사용, 아니면 신규 등록
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

        // 오더
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
        errors.push(`${b.name || "(이름없음)"}: ${err.message || err}`);
      }
    }

    blocks = blocks.filter((b) => !b.savedOk);
    busyMsg =
      `${successCount}건 저장 완료.` +
      (errors.length ? ` 실패 ${errors.length}건: ` + errors.join(" / ") : "");
    drawReviewStep();
  }

  drawUploadStep();
}
