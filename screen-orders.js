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
      container.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
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
        <h2>Customers / Orders</h2>
        <p class="muted">All staff can view all customer/order data. (Commission-related screens are admin-only)</p>
        <input type="text" id="search" placeholder="Search by name/phone/salesperson/account number" value="${escapeHtml(search)}" />
        <table class="data-table">
          <thead>
            <tr><th>Order Date</th><th>Customer Name</th><th>Phone</th><th>Salesperson</th><th>Service Items</th><th>Status</th></tr>
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
                .join("") || `<tr><td colspan="6" class="muted">No data.</td></tr>`
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
