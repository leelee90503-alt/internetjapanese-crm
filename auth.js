import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

// 로그인/회원가입 화면을 그린다. 최초 1명은 가입 후 Supabase SQL Editor에서
// profiles.role 을 'admin' 으로 직접 바꿔줘야 관리자가 된다. (README 참고)
export function renderAuthScreen(container, { onAuthed }) {
  let mode = "login"; // "login" | "signup"
  let errorMsg = "";
  let infoMsg = "";

  function draw() {
    container.innerHTML = `
      <div class="auth-wrap">
        <div class="auth-card">
          <h1>고객관리 시스템(CRM)</h1>
          <p class="muted">직원 전용 내부 시스템입니다.</p>
          <div class="tabs">
            <button data-mode="login" class="${mode === "login" ? "tab active" : "tab"}">로그인</button>
            <button data-mode="signup" class="${mode === "signup" ? "tab active" : "tab"}">직원 가입</button>
          </div>
          ${errorMsg ? `<div class="alert error">${escapeHtml(errorMsg)}</div>` : ""}
          ${infoMsg ? `<div class="alert info">${escapeHtml(infoMsg)}</div>` : ""}
          <form id="auth-form">
            ${
              mode === "signup"
                ? `<label>이름<input type="text" name="full_name" required /></label>`
                : ""
            }
            <label>이메일<input type="email" name="email" required autocomplete="username" /></label>
            <label>비밀번호<input type="password" name="password" required minlength="6" autocomplete="current-password" /></label>
            <button type="submit" class="btn primary full">${mode === "login" ? "로그인" : "가입하기"}</button>
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
          errorMsg = "로그인 실패: " + error.message;
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
          errorMsg = "가입 실패: " + error.message;
          draw();
          return;
        }
        infoMsg =
          "가입 요청이 완료되었습니다. (이메일 확인이 켜져 있다면 메일함을 확인하세요) 최초 관리자 계정은 Supabase SQL Editor에서 role 을 admin 으로 바꿔야 합니다.";
        mode = "login";
        draw();
      }
    });
  }

  draw();
}
