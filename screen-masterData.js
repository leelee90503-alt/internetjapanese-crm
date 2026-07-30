import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

const TABS = [
  { key: "services", table: "services", label: "서비스 상품", extra: [] },
  {
    key: "providers",
    table: "providers",
    label: "프로바이더",
    extra: [{ key: "has_commission_report", label: "커미션 리포트 제공", type: "checkbox" }],
  },
  { key: "salespeople", table: "salespeople", label: "세일즈 담당자", extra: [] },
];

export async function renderMasterData(container, ctx) {
  let activeTab = TABS[0].key;
  const isAdmin = ctx.profile?.role === "admin";

  async function draw() {
    const tab = TABS.find((t) => t.key === activeTab);
    container.innerHTML = `
      <div class="screen">
        <h2>마스터 데이터 관리</h2>
        <p class="muted">서비스 상품 / 프로바이더 / 세일즈 담당자를 관리합니다. ${
          isAdmin ? "" : "(일반 직원은 조회만 가능합니다)"
        }</p>
        <div class="tabs">
          ${TABS.map(
            (t) =>
              `<button data-tab="${t.key}" class="${t.key === activeTab ? "tab active" : "tab"}">${t.label}</button>`
          ).join("")}
        </div>
        <div id="tab-body">불러오는 중...</div>
      </div>
    `;
    TABS.forEach((t) => {
      container.querySelector(`[data-tab="${t.key}"]`)?.addEventListener("click", () => {
        activeTab = t.key;
        draw();
      });
    });
    await drawTabBody(tab);
  }

  async function drawTabBody(tab) {
    const body = container.querySelector("#tab-body");
    const { data, error } = await supabase
      .from(tab.table)
      .select("*")
      .order("created_at", { ascending: true });

    if (error) {
      body.innerHTML = `<div class="alert error">불러오기 실패: ${escapeHtml(error.message)}</div>`;
      return;
    }

    body.innerHTML = `
      ${isAdmin ? `
      <form id="add-form" class="inline-form">
        <input type="text" name="name" placeholder="${tab.label} 이름 추가" required />
        ${tab.extra
          .map((f) =>
            f.type === "checkbox"
              ? `<label class="checkbox-inline"><input type="checkbox" name="${f.key}" checked /> ${f.label}</label>`
              : ""
          )
          .join("")}
        <button type="submit" class="btn primary">추가</button>
      </form>` : ""}
      <table class="data-table">
        <thead>
          <tr>
            <th>이름</th>
            ${tab.extra.map((f) => `<th>${f.label}</th>`).join("")}
            <th>사용여부</th>
            ${isAdmin ? "<th>관리</th>" : ""}
          </tr>
        </thead>
        <tbody>
          ${data
            .map(
              (row) => `
            <tr data-id="${row.id}">
              <td>${
                isAdmin
                  ? `<input type="text" class="edit-name" value="${escapeHtml(row.name)}" />`
                  : escapeHtml(row.name)
              }</td>
              ${tab.extra
                .map((f) =>
                  f.type === "checkbox"
                    ? `<td><input type="checkbox" class="edit-${f.key}" ${
                        row[f.key] ? "checked" : ""
                      } ${isAdmin ? "" : "disabled"} /></td>`
                    : `<td>${escapeHtml(row[f.key])}</td>`
                )
                .join("")}
              <td><input type="checkbox" class="edit-active" ${row.is_active ? "checked" : ""} ${
                isAdmin ? "" : "disabled"
              } /></td>
              ${isAdmin ? `<td><button class="btn small save-row">저장</button></td>` : ""}
            </tr>
          `
            )
            .join("") || `<tr><td colspan="8" class="muted">데이터가 없습니다.</td></tr>`}
        </tbody>
      </table>
    `;

    if (isAdmin) {
      body.querySelector("#add-form")?.addEventListener("submit", async (e) => {
        e.preventDefault();
        const fd = new FormData(e.target);
        const payload = { name: fd.get("name") };
        tab.extra.forEach((f) => {
          if (f.type === "checkbox") payload[f.key] = fd.get(f.key) === "on";
        });
        const { error: insErr } = await supabase.from(tab.table).insert(payload);
        if (insErr) {
          alert("추가 실패: " + insErr.message);
          return;
        }
        await drawTabBody(tab);
      });

      body.querySelectorAll(".save-row").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const tr = btn.closest("tr");
          const id = tr.dataset.id;
          const payload = { name: tr.querySelector(".edit-name").value };
          tab.extra.forEach((f) => {
            if (f.type === "checkbox") payload[f.key] = tr.querySelector(`.edit-${f.key}`).checked;
          });
          payload.is_active = tr.querySelector(".edit-active").checked;
          const { error: updErr } = await supabase.from(tab.table).update(payload).eq("id", id);
          if (updErr) {
            alert("저장 실패: " + updErr.message);
            return;
          }
          await drawTabBody(tab);
        });
      });
    }
  }

  await draw();
}
