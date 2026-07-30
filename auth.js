import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

// Renders the login/signup screen. The very first person to sign up must have
// their profiles.role switched to 'admin' directly in the Supabase SQL Editor
// to become an admin (see README).
export function renderAuthScreen(container, { onAuthed }) {
  let mode = "login"; // "login" | "signup"
  let errorMsg = "";
  let infoMsg = "";

  function draw() {
    container.innerHTML = `
      <div class="auth-wrap">
        <div class="auth-card">
          <h1>Customer Management System (CRM)</h1>
          <p class="muted">Internal system for employees only.</p>
          <div class="tabs">
            <button data-mode="login" class="${mode === "login" ? "tab active" : "tab"}">Log In</button>
            <button data-mode="signup" class="${mode === "signup" ? "tab active" : "tab"}">Sign Up</button>
          </div>
          ${errorMsg ? `<div class="alert error">${escapeHtml(errorMsg)}</div>` : ""}
          ${infoMsg ? `<div class="alert info">${escapeHtml(infoMsg)}</div>` : ""}
          <form id="auth-form">
            ${
              mode === "signup"
                ? `<label>Name<input type="text" name="full_name" required /></label>`
                : ""
            }
            <label>Email<input type="email" name="email" required autocomplete="username" /></label>
            <label>Password<input type="password" name="password" required minlength="6" autocomplete="current-password" /></label>
            <button type="submit" class="btn primary full">${mode === "login" ? "Log In" : "Sign Up"}</button>
          </form>
        </div>
      </div>
    `;

    container.querySelectorAll("[data-mode]").forEach((btn) => {
      btn.addEventListener("click", () => {
        mode = btn.dataset.mode;
        errorMsg = "";
        infoMsg = "";
        draw();
      });
    });

    container.querySelector("#auth-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      errorMsg = "";
      infoMsg = "";
      const fd = new FormData(e.target);
      const email = fd.get("email");
      const password = fd.get("password");

      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) {
          errorMsg = "Login failed: " + error.message;
          draw();
          return;
        }
        onAuthed();
      } else {
        const full_name = fd.get("full_name");
        const { error } = await supabase.auth.signUp({
          email,
          password,
          options: { data: { full_name } },
        });
        if (error) {
          errorMsg = "Sign up failed: " + error.message;
          draw();
          return;
        }
        infoMsg =
          "Sign-up request complete. (If email confirmation is enabled, check your inbox.) The first admin account must have its role switched to admin in the Supabase SQL Editor.";
        mode = "login";
        draw();
      }
    });
  }

  draw();
}
