import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

// Renders the login screen. There is no public sign-up here -- new staff
// accounts are created by an admin from the Staff Account Management screen
// (screen-staff.js) inside the app, using a throwaway secondary Supabase
// client so creating a new user doesn't touch the admin's own session.
export function renderAuthScreen(container, { onAuthed }) {
  let errorMsg = "";

  function draw() {
    container.innerHTML = `
      <div class="auth-wrap">
        <div class="auth-card">
          <h1>Customer Management System (CRM)</h1>
          <p class="muted">Internal system for employees only.</p>
          ${errorMsg ? `<div class="alert error">${escapeHtml(errorMsg)}</div>` : ""}
          <form id="auth-form">
            <label>Email<input type="email" name="email" required autocomplete="username" /></label>
            <label>Password<input type="password" name="password" required minlength="6" autocomplete="current-password" /></label>
            <button type="submit" class="btn primary full">Log In</button>
          </form>
          <p class="muted" style="margin-top:12px">Don't have an account? Ask an admin to add you from Staff Account Management.</p>
        </div>
      </div>
    `;

    container.querySelector("#auth-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      errorMsg = "";
      const fd = new FormData(e.target);
      const email = fd.get("email");
      const password = fd.get("password");

      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) {
        errorMsg = "Login failed: " + error.message;
        draw();
        return;
      }
      onAuthed();
    });
  }

  draw();
}
