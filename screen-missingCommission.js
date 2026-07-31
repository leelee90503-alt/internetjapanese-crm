import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import { escapeHtml, fmtMoney } from "./normalize.js";

// Missing Commission: shows ONLY order_service_lines that staff have
// explicitly confirmed as missing (status = 'missing'), set from the
// Pending Commission screen's "Confirmed Missing" decision. This screen is
// intentionally narrow -- it is the "needs provider follow-up" list, not a
// general pending/waiting view. Everything still waiting on a report
// (including anything past the 21-day mark that hasn't been confirmed
// missing yet) lives in Pending Commission instead.
//
// This also still shows any legacy missing_commission_items rows that
// haven't been resolved yet (from the earlier manual-upload workflow),
// so nothing from before this screen was split gets silently hidden.

function daysSince(dateStr) {
  if (!dateStr) return null;
  const start = new Date(dateStr + "T00:00:00");
  const today = new Date();
  return Math.floor((today - start) / 86400000);
}

function downloadAsExcel(rows) {
  const out = rows.map((l) => ({
    Customer: l.orders?.customers?.name || "",
    "Account #": l.account_number || "",
    Provider: l.providers?.name || "",
    Plan: l.plan_name || "",
    "Order Date": l.orders?.order_date || "",
    "Order #": l.orders?.order_number || "",
    "Expected Commission": l.expected_commission === null || l.expected_commission === undefined ? "" : Number(l.expected_commission),
    "Confirmed Missing On": l.resolution_at ? new Date(l.resolution_at).toLocaleDateString() : "",
    Note: l.resolution_note || "",
  }));
  const sheet = XLSX.utils.json_to_sheet(out);
  sheet["!cols"] = [
    { wch: 20 }, // Customer
    { wch: 18 }, // Account #
    { wch: 12 }, // Provider
    { wch: 20 }, // Plan
    { wch: 12 }, // Order Date
    { wch: 18 }, // Order #
    { wch: 16 }, // Expected Commission
    { wch: 16 }, // Confirmed Missing On
    { wch: 40 }, // Note
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Missing Commission");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Missing_Commission_${dateStr}.xlsx`);
}

export async function renderMissingCommission(container, ctx) {
  let lines = [];
  let loadError = null;
  let busy = false;
  let searchTerm = "";

  // Legacy items from the old missing_commission_items workflow (manual
  // 24-item upload etc.) that are still open. Shown in a separate section
  // below the main list so old data isn't lost, without mixing its
  // different shape into the new order_service_lines-driven table.
  let legacyItems = [];

  async function load() {
    const { data, error } = await supabase
      .from("order_service_lines")
      .select(
        "id, account_number, plan_name, expected_commission, status, resolution_note, resolution_at, order_id, providers(name), orders(order_date, order_number, customers(name))"
      )
      .eq("status", "missing")
      .order("resolution_at", { ascending: false });
    if (error) {
      loadError = error.message;
      lines = [];
      return;
    }
    loadError = null;
    lines = data || [];

    const { data: legacy } = await supabase
      .from("missing_commission_items")
      .select("*")
      .eq("resolved", false)
      .order("created_at", { ascending: false });
    legacyItems = legacy || [];
  }

  function filteredLines() {
    if (!searchTerm.trim()) return lines;
    const needle = searchTerm.trim().toLowerCase();
    return lines.filter((l) => {
      const name = (l.orders?.customers?.name || "").toLowerCase();
      const acct = (l.account_number || "").toString().toLowerCase();
      const plan = (l.plan_name || "").toLowerCase();
      const provider = (l.providers?.name || "").toLowerCase();
      return name.includes(needle) || acct.includes(needle) || plan.includes(needle) || provider.includes(needle);
    });
  }

  function rowHtml(l) {
    const days = daysSince(l.orders?.order_date);
    return `
      <tr>
        <td>${escapeHtml(l.orders?.customers?.name || "-")}</td>
        <td>${escapeHtml(l.account_number || "-")}</td>
        <td>${escapeHtml(l.providers?.name || "-")}</td>
        <td>${escapeHtml(l.plan_name || "-")}</td>
        <td>${escapeHtml(l.orders?.order_date || "-")}${days !== null ? ` <span class="muted">(${days}d ago)</span>` : ""}</td>
        <td>${fmtMoney(l.expected_commission)}</td>
        <td>${escapeHtml(l.resolution_note || "-")}</td>
        <td>
          <button class="btn small" data-back-to-pending="${l.id}" ${busy ? "disabled" : ""}>Back to Pending</button>
          <button class="btn small primary" data-mark-received="${l.id}" ${busy ? "disabled" : ""}>Mark Received</button>
        </td>
      </tr>
    `;
  }

  function legacyRowHtml(item) {
    return `
      <tr>
        <td>${escapeHtml(item.customer_name || "-")}</td>
        <td>${escapeHtml(item.account_number || "-")}</td>
        <td>-</td>
        <td>${escapeHtml(item.description || "-")}</td>
        <td>${escapeHtml(item.sales_date || "-")}</td>
        <td>${fmtMoney(item.price)}</td>
        <td>${escapeHtml(item.review_note || item.status_notes || "-")}</td>
        <td>
          <button class="btn small primary" data-legacy-received="${item.id}" ${busy ? "disabled" : ""}>Mark Received</button>
          <button class="btn small danger" data-legacy-delete="${item.id}" ${busy ? "disabled" : ""}>Delete</button>
        </td>
      </tr>
    `;
  }

  function draw() {
    // Preserve focus/cursor position in the search box across the innerHTML
    // replacement below -- otherwise every keystroke re-renders the DOM and
    // drops focus, so only the first character of anything typed lands.
    const active = document.activeElement;
    const activeId = active && container.contains(active) ? active.id : null;
    const activeSelStart = activeId && "selectionStart" in active ? active.selectionStart : null;
    const activeSelEnd = activeId && "selectionEnd" in active ? active.selectionEnd : null;

    const rows = filteredLines();
    const totalExpected = rows.reduce((s, l) => s + (Number(l.expected_commission) || 0), 0);

    container.innerHTML = `
      <div class="screen">
        <h2>Missing Commission</h2>
        <p class="muted">
          Orders staff have explicitly confirmed as missing commission -- these need direct follow-up with the provider,
          they are not just "still waiting". To confirm a still-pending order as missing (or to change your mind), use
          the "Confirmed Missing" decision on the Pending Commission screen.
        </p>

        ${loadError ? `<div class="alert error">Failed to load: ${escapeHtml(loadError)}</div>` : ""}

        <div class="inline-form">
          <span class="badge error">Total: ${rows.length} confirmed missing${searchTerm.trim() ? ` (filtered from ${lines.length})` : ""}</span>
          <span class="badge neutral">Total expected: ${fmtMoney(totalExpected)}</span>
          <button class="btn small" id="mc-download-btn">Download Excel</button>
        </div>

        <div class="inline-form">
          <input type="text" placeholder="Search customer / account / plan / provider" id="mc-search" value="${escapeHtml(searchTerm)}" style="flex:1" />
        </div>

        <table class="data-table">
          <thead>
            <tr>
              <th>Customer</th><th>Account #</th><th>Provider</th><th>Plan</th>
              <th>Order Date</th><th>Expected</th><th>Note</th><th></th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(rowHtml).join("") || `<tr><td colspan="8" class="muted">${searchTerm.trim() ? "No matches for this search." : "Nothing confirmed missing right now."}</td></tr>`}
          </tbody>
        </table>

        ${
          legacyItems.length > 0
            ? `
          <h3>Older manually-tracked items</h3>
          <p class="muted">
            From before this screen was split from Pending Commission -- not yet resolved. New items won't appear here going forward.
          </p>
          <table class="data-table">
            <thead>
              <tr>
                <th>Customer</th><th>Account #</th><th>Provider</th><th>Missing</th>
                <th>Sales Date</th><th>Price</th><th>Note</th><th></th>
              </tr>
            </thead>
            <tbody>
              ${legacyItems.map(legacyRowHtml).join("")}
            </tbody>
          </table>`
            : ""
        }
      </div>
    `;

    wireEvents();

    if (activeId) {
      const toRefocus = container.querySelector(`#${activeId}`);
      if (toRefocus) {
        toRefocus.focus();
        if (activeSelStart !== null && "setSelectionRange" in toRefocus) {
          try {
            toRefocus.setSelectionRange(activeSelStart, activeSelEnd);
          } catch {
            // some input types don't support setSelectionRange -- ignore
          }
        }
      }
    }
  }

  function wireEvents() {
    container.querySelector("#mc-search")?.addEventListener("input", (e) => {
      searchTerm = e.target.value;
      draw();
    });

    container.querySelector("#mc-download-btn")?.addEventListener("click", () => {
      const rows = filteredLines();
      if (rows.length === 0) {
        alert("Nothing to download.");
        return;
      }
      downloadAsExcel(rows);
    });

    container.querySelectorAll("[data-back-to-pending]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.backToPending;
        busy = true;
        draw();
        const { error } = await supabase
          .from("order_service_lines")
          .update({ status: "pending", resolution_note: null, resolution_at: null, resolution_by: null })
          .eq("id", id);
        busy = false;
        if (error) alert("Failed to move back to pending: " + error.message);
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-mark-received]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.markReceived;
        const line = lines.find((l) => l.id === id);
        if (!line) return;
        if (!confirm(`Mark "${line.orders?.customers?.name || "this line"}" - ${line.plan_name || ""} as received?`)) return;
        busy = true;
        draw();
        const { error } = await supabase
          .from("order_service_lines")
          .update({
            status: "received",
            actual_commission_amount: line.expected_commission,
            commission_matched_at: new Date().toISOString(),
            commission_matched_by: ctx.profile?.id || null,
          })
          .eq("id", id);
        busy = false;
        if (error) alert("Failed to mark received: " + error.message);
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-legacy-received]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.legacyReceived;
        const item = legacyItems.find((i) => i.id === id);
        if (!item) return;
        const lineIds = item.matched_line_ids || [];
        if (
          !confirm(
            `Mark "${item.customer_name}" - ${item.description} as received?${
              lineIds.length ? ` This will also set ${lineIds.length} linked order line(s) to "received".` : ""
            }`
          )
        ) {
          return;
        }
        busy = true;
        draw();
        if (lineIds.length > 0) {
          const { data: relatedLines } = await supabase.from("order_service_lines").select("id, expected_commission").in("id", lineIds);
          for (const l of relatedLines || []) {
            await supabase
              .from("order_service_lines")
              .update({
                status: "received",
                actual_commission_amount: l.expected_commission,
                commission_matched_at: new Date().toISOString(),
                commission_matched_by: ctx.profile?.id || null,
              })
              .eq("id", l.id);
          }
        }
        const { error } = await supabase.from("missing_commission_items").update({ resolved: true }).eq("id", id);
        busy = false;
        if (error) alert("Failed to mark received: " + error.message);
        await load();
        draw();
      });
    });

    container.querySelectorAll("[data-legacy-delete]").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.legacyDelete;
        const item = legacyItems.find((i) => i.id === id);
        if (!confirm(`Delete this older tracked item for "${item?.customer_name || ""}"? This does not change any order line.`)) return;
        busy = true;
        draw();
        const { error } = await supabase.from("missing_commission_items").delete().eq("id", id);
        busy = false;
        if (error) alert("Failed to delete: " + error.message);
        await load();
        draw();
      });
    });
  }

  await load();
  draw();
}
