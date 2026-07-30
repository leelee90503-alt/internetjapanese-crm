import { supabase } from "./supabaseClient.js";
import { escapeHtml } from "./normalize.js";

export async function renderDashboard(container, ctx) {
  const [{ count: customerCount }, { count: orderCount }, { count: lineCount }] = await Promise.all([
    supabase.from("customers").select("*", { count: "exact", head: true }),
    supabase.from("orders").select("*", { count: "exact", head: true }),
    supabase.from("order_service_lines").select("*", { count: "exact", head: true }),
  ]);

  container.innerHTML = `
    <div class="screen">
      <h2>대시보드</h2>
      <p class="muted">${escapeHtml(ctx.profile?.full_name || ctx.profile?.email || "")}님, 안녕하세요. (권한: ${
    ctx.profile?.role === "admin" ? "관리자" : "일반 직원"
  })</p>
      <div class="stat-grid">
        <div class="stat-card"><div class="stat-num">${customerCount ?? "-"}</div><div class="stat-label">전체 고객 수</div></div>
        <div class="stat-card"><div class="stat-num">${orderCount ?? "-"}</div><div class="stat-label">전체 오더 수</div></div>
        <div class="stat-card"><div class="stat-num">${lineCount ?? "-"}</div><div class="stat-label">전체 서비스 항목 수</div></div>
      </div>
      <div class="card">
        <h3>안내</h3>
        <p>이번 버전은 1차(핵심) 범위입니다: 마스터 데이터 관리, 고객 엑셀 업로드·컨펌, 오더 관리, 직원 권한 관리.</p>
        <p class="muted">커미션 리포트 매칭, SMS/이메일 발송, AS 티켓, 인센티브 지급 등은 기획안 3장 범위를 따라 다음 단계에서 추가될 예정입니다.</p>
      </div>
    </div>
  `;
}
