import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

export async function renderDashboard(container, ctx) {
  const [{ count: customerCount }, { count: orderCount }, { count: lineCount }] = await Promise.all([
    supabase.from("customers").select("*", { count: "exact", head: true }),
    supabase.from("orders").select("*", { count: "exact", head: true }),
    supabase.from("order_service_lines").select("*", { count: "exact", head: true }),
  ]);

  container.innerHTML = `
    <div class="screen">
      <h2>Dashboard</h2>
      <p class="muted">Welcome, ${escapeHtml(ctx.profile?.full_name || ctx.profile?.email || "")}. (Role: ${
    ctx.profile?.role === "admin" ? "Admin" : "Staff"
  })</p>
      <div class="stat-grid">
        <div class="stat-card"><div class="stat-num">${customerCount ?? "-"}</div><div class="stat-label">Total Customers</div></div>
        <div class="stat-card"><div class="stat-num">${orderCount ?? "-"}</div><div class="stat-label">Total Orders</div></div>
        <div class="stat-card"><div class="stat-num">${lineCount ?? "-"}</div><div class="stat-label">Total Service Items</div></div>
      </div>
      <div class="card">
        <h3>Notes</h3>
        <p>This version covers the first (core) phase: master data management, customer Excel upload &amp; confirmation, order management, and staff permission management.</p>
        <p class="muted">Commission report matching, SMS/email sending, support tickets, and incentive payouts will be added in a later phase.</p>
      </div>
    </div>
  `;
}
