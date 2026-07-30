import { supabase } from "./supabaseClient.js";
import { escapeHtml, fmtMoney } from "./normalize.js";

const TABS = [
  { key: "services", table: "services", label: "Services", extra: [] },
  {
    key: "providers",
    table: "providers",
    label: "Providers",
    extra: [{ key: "has_commission_report", label: "Provides Commission Report", type: "checkbox" }],
  },
  { key: "salespeople", table: "salespeople", label: "Salespeople", extra: [] },
  { key: "commission_rates", label: "Commission Rates" },
];

// Parses text pasted straight from a provider's commission sheet, in the
// "Plan Name" / "$123.45" alternating-line format (e.g. copy-pasted from a
// PDF or spreadsheet): every line is either a plan name or an amount: the
// first line that looks like a dollar amount is paired with the plan-name
// line(s) that came before it since the last amount. Blank lines are
// ignored.
function parseBulkCommissionText(text) {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l !== "");
  const amountRe = /^\$?\s*-?[\d,]+(\.\d+)?$/;
  const rows = [];
  let pendingName = [];
  for (const line of lines) {
    if (amountRe.test(line)) {
      const rate = Number(line.replace(/[$,]/g, ""));
      const name = pendingName.join(" ").trim();
      if (name && !isNaN(rate)) rows.push({ plan_name: name, rate });
      pendingName = [];
    } else {
      pendingName.push(line);
    }
  }
  return rows;
}

export async function renderMasterData(container, ctx) {
  let activeTab = TABS[0].key;
  const isAdmin = ctx.profile?.role === "admin";

  // Commission Rates tab state
  let providers = [];
  let selectedProviderId = null;
  let rates = [];
  let bulkText = "";
  let bulkBusy = false;

  async function draw() {
    const tab = TABS.find((t) => t.key === activeTab);
    container.innerHTML = `
      <div class="screen">
        <h2>Master Data Management</h2>
        <p class="muted">Manage Services / Providers / Salespeople / Commission Rates. ${
          isAdmin ? "" : "(Staff can view only)"
        }</p>
        <div class="tabs">
          ${TABS.map(
            (t) =>
              `<button data-tab="${t.key}" class="${t.key === activeTab ? "tab active" : "tab"}">${t.label}</button>`
          ).join("")}
        </div>
        <div id="tab-body">Loading...</div>
      </div>
    `;
    TABS.forEach((t) => {
      container.querySelector(`[data-tab="${t.key}"]`)?.addEventListener("click", () => {
        activeTab = t.key;
        draw();
      });
    });
    if (tab.key === "commission_rates") {
      await drawCommissionRatesTab();
    } else {
      await drawTabBody(tab);
    }
  }

  async function drawTabBody(tab) {
    const body = container.querySelector("#tab-body");
    const { data, error } = await supabase
      .from(tab.table)
      .select("*")
      .order("created_at", { ascending: true });

    if (error) {
      body.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
      return;
    }

    body.innerHTML = `
      ${isAdmin ? `
      <form id="add-form" class="inline-form">
        <input type="text" name="name" placeholder="Add a new ${tab.label} name" required />
        ${tab.extra
          .map((f) =>
            f.type === "checkbox"
              ? `<label class="checkbox-inline"><input type="checkbox" name="${f.key}" checked /> ${f.label}</label>`
              : ""
          )
          .join("")}
        <button type="submit" class="btn primary">Add</button>
      </form>` : ""}
      <table class="data-table">
        <thead>
          <tr>
            <th>Name</th>
            ${tab.extra.map((f) => `<th>${f.label}</th>`).join("")}
            <th>Active</th>
            ${isAdmin ? "<th>Actions</th>" : ""}
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
              ${isAdmin ? `<td><button class="btn small save-row">Save</button></td>` : ""}
            </tr>
          `
            )
            .join("") || `<tr><td colspan="8" class="muted">No data.</td></tr>`}
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
          alert("Failed to add: " + insErr.message);
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
            alert("Failed to save: " + updErr.message);
            return;
          }
          await drawTabBody(tab);
        });
      });
    }
  }

  async function loadProviders() {
    const { data, error } = await supabase.from("providers").select("id, name").order("name");
    providers = error ? [] : data || [];
    if (!selectedProviderId && providers.length > 0) selectedProviderId = providers[0].id;
  }

  async function loadRates() {
    if (!selectedProviderId) {
      rates = [];
      return;
    }
    const { data, error } = await supabase
      .from("commission_rates")
      .select("id, plan_name, rate, is_active")
      .eq("provider_id", selectedProviderId)
      .order("plan_name", { ascending: true });
    rates = error ? [] : data || [];
  }

  async function drawCommissionRatesTab() {
    const body = container.querySelector("#tab-body");
    body.innerHTML = `<p class="muted">Loading...</p>`;
    if (providers.length === 0) await loadProviders();
    await loadRates();

    body.innerHTML = `
      <p class="muted">
        The $ amount expected per plan/package, by provider. Used to calculate expected commission.
        ${isAdmin ? "" : "(Staff can view only)"}
      </p>
      <div class="inline-form">
        <label>Provider
          <select id="rate-provider-select">
            ${providers
              .map(
                (p) => `<option value="${p.id}" ${p.id === selectedProviderId ? "selected" : ""}>${escapeHtml(p.name)}</option>`
              )
              .join("")}
          </select>
        </label>
      </div>

      ${
        isAdmin
          ? `<form id="add-rate-form" class="inline-form">
              <input type="text" name="plan_name" placeholder="Plan / Package name" required />
              <input type="number" step="0.01" name="rate" placeholder="Rate ($)" required style="width:120px" />
              <button type="submit" class="btn primary">Add</button>
            </form>`
          : ""
      }

      <table class="data-table">
        <thead>
          <tr><th>Plan Name</th><th>Rate</th><th>Active</th>${isAdmin ? "<th>Actions</th>" : ""}</tr>
        </thead>
        <tbody>
          ${
            rates
              .map(
                (r) => `
            <tr data-id="${r.id}">
              <td>${isAdmin ? `<input type="text" class="edit-plan-name" value="${escapeHtml(r.plan_name)}" />` : escapeHtml(r.plan_name)}</td>
              <td>${isAdmin ? `<input type="number" step="0.01" class="edit-rate" value="${r.rate}" style="width:120px" />` : fmtMoney(r.rate)}</td>
              <td><input type="checkbox" class="edit-rate-active" ${r.is_active ? "checked" : ""} ${isAdmin ? "" : "disabled"} /></td>
              ${
                isAdmin
                  ? `<td><button class="btn small save-rate-row">Save</button> <button class="btn small danger delete-rate-row">Delete</button></td>`
                  : ""
              }
            </tr>`
              )
              .join("") || `<tr><td colspan="${isAdmin ? 4 : 3}" class="muted">No rates yet for this provider.</td></tr>`
          }
        </tbody>
      </table>

      ${
        isAdmin
          ? `<h4>Bulk Import</h4>
             <p class="muted">Paste a plan name and its $ amount on alternating lines (e.g. copied straight from a rate sheet) -- an existing plan with the same name for this provider gets its rate updated, a new one gets added.</p>
             <div class="inline-form">
               <textarea id="bulk-rate-input" placeholder="2 Gig&#10;$560.00&#10;Advantage&#10;$480.00" rows="6" style="flex:1;font-family:monospace">${escapeHtml(bulkText)}</textarea>
             </div>
             <div class="inline-form">
               <button type="button" class="btn primary" id="bulk-import-btn" ${bulkBusy ? "disabled" : ""}>${bulkBusy ? "Importing..." : "Import"}</button>
             </div>`
          : ""
      }
    `;

    container.querySelector("#rate-provider-select")?.addEventListener("change", async (e) => {
      selectedProviderId = e.target.value;
      await loadRates();
      await drawCommissionRatesTab();
    });

    if (!isAdmin) return;

    container.querySelector("#add-rate-form")?.addEventListener("submit", async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const planName = String(fd.get("plan_name") || "").trim();
      const rate = Number(fd.get("rate"));
      if (!planName || isNaN(rate)) return;
      const { error } = await supabase
        .from("commission_rates")
        .insert({ provider_id: selectedProviderId, plan_name: planName, rate });
      if (error) {
        alert("Failed to add: " + error.message);
        return;
      }
      await loadRates();
      await drawCommissionRatesTab();
    });

    container.querySelectorAll(".save-rate-row").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const tr = btn.closest("tr");
        const id = tr.dataset.id;
        const payload = {
          plan_name: tr.querySelector(".edit-plan-name").value.trim(),
          rate: Number(tr.querySelector(".edit-rate").value),
          is_active: tr.querySelector(".edit-rate-active").checked,
        };
        const { error } = await supabase.from("commission_rates").update(payload).eq("id", id);
        if (error) {
          alert("Failed to save: " + error.message);
          return;
        }
        await loadRates();
        await drawCommissionRatesTab();
      });
    });

    container.querySelectorAll(".delete-rate-row").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const tr = btn.closest("tr");
        const id = tr.dataset.id;
        const ok = confirm("Delete this commission rate? This cannot be undone.");
        if (!ok) return;
        const { error } = await supabase.from("commission_rates").delete().eq("id", id);
        if (error) {
          alert("Failed to delete: " + error.message);
          return;
        }
        await loadRates();
        await drawCommissionRatesTab();
      });
    });

    container.querySelector("#bulk-rate-input")?.addEventListener("input", (e) => {
      bulkText = e.target.value;
    });

    container.querySelector("#bulk-import-btn")?.addEventListener("click", async () => {
      const parsed = parseBulkCommissionText(bulkText);
      if (parsed.length === 0) {
        alert("Couldn't find any plan name / amount pairs in the pasted text.");
        return;
      }
      bulkBusy = true;
      await drawCommissionRatesTab();

      const existingByName = new Map(rates.map((r) => [r.plan_name.trim().toLowerCase(), r]));
      const errors = [];
      for (const row of parsed) {
        const existing = existingByName.get(row.plan_name.trim().toLowerCase());
        if (existing) {
          const { error } = await supabase.from("commission_rates").update({ rate: row.rate }).eq("id", existing.id);
          if (error) errors.push(`${row.plan_name}: ${error.message}`);
        } else {
          const { error } = await supabase
            .from("commission_rates")
            .insert({ provider_id: selectedProviderId, plan_name: row.plan_name, rate: row.rate });
          if (error) errors.push(`${row.plan_name}: ${error.message}`);
        }
      }

      bulkBusy = false;
      bulkText = "";
      await loadRates();
      await drawCommissionRatesTab();
      if (errors.length > 0) {
        alert(`Imported with ${errors.length} error(s):\n` + errors.join("\n"));
      } else {
        alert(`Imported ${parsed.length} plan(s).`);
      }
    });
  }

  await draw();
}
