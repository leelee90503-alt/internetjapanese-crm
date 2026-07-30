import { supabase } from "./supabaseClient.js";
import { escapeHtml, fmtMoney } from "./normalize.js";

export async function renderOrders(container, ctx) {
  let search = "";
  let orders = [];

  async function load() {
    let query = supabase
      .from("orders")
      .select(
        `id, order_date, pipeline_stage, memo,
         customers(id, name, phone),
         salespeople(id, name),
         order_service_lines(id, account_number, expected_commission, status, services(name), providers(name))`
      )
      .order("order_date", { ascending: false });
    const { data, error } = await query;
    if (error) {
      container.innerHTML = `<div class="alert error">불러오기 실패: ${escapeHtml(error.message)}</div>`;
      return;
    }
    orders = data || [];
  }

  function filtered() {
    if (!search.trim()) return orders;
    const q = search.trim().toLowerCase();
    return orders.filter(
      (o) =>
        (o.customers?.name || "").toLowerCase().includes(q) ||
        (o.customers?.phone || "").toLowerCase().includes(q) ||
        (o.salespeople?.name || "").toLowerCase().includes(q) ||
        o.order_service_lines?.some((l) => (l.account_number || "").toLowerCase().includes(q))
    );
  }

  function draw() {
    const rows = filtered();
    container.innerHTML = `
      <div class="screen">
        <h2>고객 / 오더 관리</h2>
        <p class="muted">전 직원이 전체 고객·오더 데이터를 조회할 수 있습니다. (커미션 관련 화면은 관리자 전용입니다)</p>
        <input type="text" id="search" placeholder="이름/전화번호/세일즈담당자/계정번호 검색" value="${escapeHtml(search)}" />
        <table class="data-table">
          <thead>
            <tr><th>오더날짜</th><th>고객명</th><th>전화번호</th><th>세일즈담당자</th><th>서비스 항목</th><th>상태</th></tr>
          </thead>
          <tbody>
            ${
              rows
                .map(
                  (o) => `
              <tr>
                <td>${o.order_date || "-"}</td>
                <td>${escapeHtml(o.customers?.name || "-")}</td>
                <td>${escapeHtml(o.customers?.phone || "-")}</td>
                <td>${escapeHtml(o.salespeople?.name || "-")}</td>
                <td>${(o.order_service_lines || [])
                  .map(
                    (l) =>
                      `${escapeHtml(l.services?.name || "?")}/${escapeHtml(l.providers?.name || "?")}${
                        l.account_number ? " (" + escapeHtml(l.account_number) + ")" : ""
                      }`
                  )
                  .join(", ") || "-"}</td>
                <td>${escapeHtml(o.pipeline_stage)}</td>
              </tr>`
                )
                .join("") || `<tr><td colspan="6" class="muted">데이터가 없습니다.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    `;
    container.querySelector("#search").addEventListener("input", (e) => {
      search = e.target.value;
      draw();
    });
  }

  await load();
  draw();
}
