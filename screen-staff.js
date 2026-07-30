import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

export async function renderStaff(container, ctx) {
  async function draw() {
    const { data, error } = await supabase
      .from("profiles")
      .select("*")
      .order("created_at", { ascending: true });
    if (error) {
      container.innerHTML = `<div class="alert error">불러오기 실패: ${escapeHtml(error.message)}</div>`;
      return;
    }
    container.innerHTML = `
      <div class="screen">
        <h2>직원 계정 관리 (관리자 전용)</h2>
        <p class="muted">직원은 화면에서 직접 가입하며, 최초에는 모두 "일반" 권한으로 생성됩니다. 관리자가 여기서 권한을 바꿔줄 수 있습니다.</p>
        <table class="data-table">
          <thead><tr><th>이름</th><th>이메일</th><th>권한</th><th>사용여부</th><th>관리</th></tr></thead>
          <tbody>
            ${data
              .map(
                (p) => `
              <tr data-id="${p.id}">
                <td>${escapeHtml(p.full_name || "-")}</td>
                <td>${escapeHtml(p.email || "-")}</td>
                <td>
                  <select class="p-role">
                    <option value="staff" ${p.role === "staff" ? "selected" : ""}>일반</option>
                    <option value="admin" ${p.role === "admin" ? "selected" : ""}>관리자</option>
                  </select>
                </td>
                <td><input type="checkbox" class="p-active" ${p.is_active ? "checked" : ""} /></td>
                <td><button class="btn small save-row">저장</button></td>
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
          alert("저장 실패: " + updErr.message);
          return;
        }
        await draw();
      });
    });
  }
  await draw();
}
