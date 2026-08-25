import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

// Renders the login screen. There is no public sign-up here -- new staff
// accounts are created by an admin from the Staff Account Management screen
// (screen-staff.js) inside the app, using a throwaway secondary Supabase
// client so creating a new user doesn't touch the admin's own session.
export function renderAuthScreen(container, { onAuthed }) {
  let mode = "login"; // "login" | "forgot"
  let errorMsg = "";
  let infoMsg = "";

  function draw() {
    if (mode === "forgot") {
      container.innerHTML = `
        <div class="auth-wrap">
          <div class="auth-card">
            <h1>Reset your password</h1>
            <p class="muted">Enter your email and we'll send you a password reset link.</p>
            ${errorMsg ? `<div class="alert error">${escapeHtml(errorMsg)}</div>` : ""}
            ${infoMsg ? `<div class="alert info">${escapeHtml(infoMsg)}</div>` : ""}
            <form id="forgot-form">
              <label>Email<input type="email" name="email" required autocomplete="username" /></label>
              <button type="submit" class="btn primary full">Send Reset Link</button>
            </form>
            <p class="muted" style="margin-top:12px"><a href="#" id="back-to-login">Back to log in</a></p>
          </div>
        </div>
      `;
      container.querySelector("#back-to-login").addEventListener("click", (e) => {
        e.preventDefault();
        mode = "login";
        errorMsg = "";
        infoMsg = "";
        draw();
      });
      container.querySelector("#forgot-form").addEventListener("submit", async (e) => {
        e.preventDefault();
        errorMsg = "";
        infoMsg = "";
        const fd = new FormData(e.target);
        const email = fd.get("email");
        const { error } = await supabase.auth.resetPasswordForEmail(email, {
          redirectTo: window.location.origin + window.location.pathname,
        });
        if (error) {
          errorMsg = "Failed to send reset link: " + error.message;
          draw();
          return;
        }
        infoMsg = "If an account exists for that email, a reset link has been sent. Check your inbox (and spam folder).";
        draw();
      });
      return;
    }

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
          <p class="muted" style="margin-top:12px"><a href="#" id="forgot-link">Forgot your password?</a></p>
          <p class="muted">Don't have an account? Ask an admin to add you from Staff Account Management.</p>
        </div>
      </div>
    `;

    container.querySelector("#forgot-link").addEventListener("click", (e) => {
      e.preventDefault();
      mode = "forgot";
      errorMsg = "";
      infoMsg = "";
      draw();
    });

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

// Shown when Supabase detects a password-recovery link (the user clicked the
// "reset password" email). Lets them set a new password, then hands control
// back to the caller to resume normal app boot.
export function renderSetNewPasswordScreen(container, { onDone }) {
  let errorMsg = "";
  let done = false;

  function draw() {
    container.innerHTML = `
      <div class="auth-wrap">
        <div class="auth-card">
          <h1>Set a new password</h1>
          ${
            done
              ? `
            <div class="alert info">Your password has been updated.</div>
            <button class="btn primary full" id="continue-btn">Continue</button>
          `
              : `
            ${errorMsg ? `<div class="alert error">${escapeHtml(errorMsg)}</div>` : ""}
            <form id="new-password-form">
              <label>New password<input type="password" name="password" required minlength="6" autocomplete="new-password" /></label>
              <label>Confirm new password<input type="password" name="confirm" required minlength="6" autocomplete="new-password" /></label>
              <button type="submit" class="btn primary full">Set Password</button>
            </form>
          `
          }
        </div>
      </div>
    `;

    if (done) {
      container.querySelector("#continue-btn").addEventListener("click", () => onDone());
      return;
    }

    container.querySelector("#new-password-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      errorMsg = "";
      const fd = new FormData(e.target);
      const password = fd.get("password");
      const confirm = fd.get("confirm");
      if (password !== confirm) {
        errorMsg = "Passwords do not match.";
        draw();
        return;
      }
      const { error } = await supabase.auth.updateUser({ password });
      if (error) {
        errorMsg = "Failed to set password: " + error.message;
        draw();
        return;
      }
      done = true;
      draw();
    });
  }

  draw();
}
