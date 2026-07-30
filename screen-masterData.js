import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

const TABS = [
  { key: "services", table: "services", label: "Services", extra: [] },
  {
    key: "providers",
    table: "providers",
    label: "Providers",
    extra: [{ key: "has_commission_report", label: "Provides Commission Report", type: "checkbox" }],
  },
  { key: "salespeople", table: "salespeople", label: "Salespeople", extra: [] },
];

export async function renderMasterData(container, ctx) {
  let activeTab = TABS[0].key;
  const isAdmin = ctx.profile?.role === "admin";

  async function draw() {
    const tab = TABS.find((t) => t.key === activeTab);
    container.innerHTML = `
      <div class="screen">
        <h2>Master Data Management</h2>
        <p class="muted">Manage Services / Providers / Salespeople. ${
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
    await drawTabBody(tab);
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

  await draw();
}
