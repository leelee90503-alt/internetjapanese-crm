import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

export async function renderStaff(container, ctx) {
  async function draw() {
    const { data, error } = await supabase
      .from("profiles")
      .select("*")
      .order("created_at", { ascending: true });
    if (error) {
      container.innerHTML = `<div class="alert error">Failed to load: ${escapeHtml(error.message)}</div>`;
      return;
    }
    container.innerHTML = `
      <div class="screen">
        <h2>Staff Account Management (Admin Only)</h2>
        <p class="muted">Staff sign up directly from the screen and are all created with "Staff" permission initially. Admins can change permissions here.</p>
        <table class="data-table">
          <thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Active</th><th>Actions</th></tr></thead>
          <tbody>
            ${data
              .map(
                (p) => `
              <tr data-id="${p.id}">
                <td>${escapeHtml(p.full_name || "-")}</td>
                <td>${escapeHtml(p.email || "-")}</td>
                <td>
                  <select class="p-role">
                    <option value="staff" ${p.role === "staff" ? "selected" : ""}>Staff</option>
                    <option value="admin" ${p.role === "admin" ? "selected" : ""}>Admin</option>
                  </select>
                </td>
                <td><input type="checkbox" class="p-active" ${p.is_active ? "checked" : ""} /></td>
                <td><button class="btn small save-row">Save</button></td>
              </tr>`
              )
              .join("")}
          </tbody>
        </table>
      </div>
    `;
    container.querySelectorAll(".save-row").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const tr = btn.closest("tr");
        const id = tr.dataset.id;
        const role = tr.querySelector(".p-role").value;
        const is_active = tr.querySelector(".p-active").checked;
        const { error: updErr } = await supabase.from("profiles").update({ role, is_active }).eq("id", id);
        if (updErr) {
          alert("Failed to save: " + updErr.message);
          return;
        }
        await draw();
      });
    });
  }
  await draw();
}
