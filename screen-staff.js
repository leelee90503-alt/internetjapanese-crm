import { createClient } from "@supabase/supabase-js";
import { supabase } from "./supabaseClient.js";
import { SUPABASE_URL, SUPABASE_ANON_KEY } from "./config.js";
import { escapeHtml } from "./normalize.js";

// Staff Account Management (admin only). There is no public sign-up screen
// any more -- this is the only way new employee accounts get created. New
// accounts are created with a throwaway, non-session-persisting Supabase
// client (persistSession: false) so calling auth.signUp() here never
// overwrites the admin's own logged-in session in this browser tab.
export async function renderStaff(container, ctx) {
  let busy = false;
  let formError = "";
  let formInfo = "";

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
        <p class="muted">There is no public sign-up -- add new employee accounts here. New accounts are created with "Staff" permission by default; change permissions below any time.</p>

        <h3>Add Staff Account</h3>
        ${formError ? `<div class="alert error">${escapeHtml(formError)}</div>` : ""}
        ${formInfo ? `<div class="alert info">${escapeHtml(formInfo)}</div>` : ""}
        <form id="add-staff-form" class="inline-form">
          <label>Name<input type="text" name="full_name" required /></label>
          <label>Email<input type="email" name="email" required autocomplete="off" /></label>
          <label>Password<input type="password" name="password" required minlength="6" autocomplete="new-password" /></label>
          <label>Role
            <select name="role">
              <option value="staff" selected>Staff</option>
              <option value="admin">Admin</option>
            </select>
          </label>
          <button type="submit" class="btn primary" ${busy ? "disabled" : ""}>${busy ? "Creating..." : "Create Account"}</button>
        </form>

        <h3 style="margin-top:24px">Existing Accounts</h3>
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

    container.querySelector("#add-staff-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      formError = "";
      formInfo = "";
      const fd = new FormData(e.target);
      const full_name = fd.get("full_name");
      const email = fd.get("email");
      const password = fd.get("password");
      const role = fd.get("role");

      busy = true;
      await draw();

      // A separate, non-persisting client -- signUp() logs the *new* user in
      // by default, and if we used the shared `supabase` client here that
      // would kick the currently-logged-in admin out of their own session in
      // this tab. persistSession: false keeps this whole exchange off to the
      // side; we throw the client away right after.
      const tempClient = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
        auth: { persistSession: false, autoRefreshToken: false },
      });
      const { data: signUpData, error: signUpError } = await tempClient.auth.signUp({
        email,
        password,
        options: { data: { full_name } },
      });

      if (signUpError) {
        busy = false;
        formError = "Failed to create account: " + signUpError.message;
        await draw();
        return;
      }

      const newUserId = signUpData?.user?.id;
      if (role === "admin" && newUserId) {
        // profiles row is created with role='staff' by the on_auth_user_created
        // trigger -- promote it to admin right away using the admin's own
        // (already authenticated, RLS-permitted) client.
        const { error: roleErr } = await supabase.from("profiles").update({ role: "admin" }).eq("id", newUserId);
        if (roleErr) {
          formError = `Account created, but failed to set role to Admin: ${roleErr.message}. You can change it below.`;
        }
      }

      busy = false;
      if (!formError) {
        formInfo = `Account created for ${email}. If email confirmation is enabled on this project, they'll need to confirm their email before they can log in.`;
      }
      await draw();
    });

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
