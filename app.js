import { supabase } from "./supabaseClient.js";
import { renderAuthScreen } from "./auth.js";
import { renderDashboard } from "./screen-dashboard.js";
import { renderMasterData } from "./screen-masterData.js";
import { renderCustomerUpload } from "./screen-customerUpload.js";
import { renderOrders } from "./screen-orders.js";
import { renderStaff } from "./screen-staff.js";
import { escapeHtml } from "./normalize.js";

const ROUTES = [
  { path: "#/dashboard", label: "대시보드", render: renderDashboard, adminOnly: false },
  { path: "#/upload", label: "고객 엑셀 업로드", render: renderCustomerUpload, adminOnly: false },
  { path: "#/orders", label: "고객/오더 관리", render: renderOrders, adminOnly: false },
  { path: "#/master-data", label: "마스터 데이터", render: renderMasterData, adminOnly: false },
  { path: "#/staff", label: "직원 관리", render: renderStaff, adminOnly: true },
];

const root = document.getElementById("app");
let session = null;
let profile = null;

async function loadProfile() {
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    profile = null;
    return;
  }
  const { data, error } = await supabase.from("profiles").select("*").eq("id", user.id).single();
  if (error) {
    console.error("프로필 조회 실패", error);
    profile = null;
    return;
  }
  profile = data;
}

function currentRoute() {
  const hash = location.hash || "#/dashboard";
  return ROUTES.find((r) => r.path === hash) || ROUTES[0];
}

async function renderApp() {
  if (!session) {
    renderAuthScreen(root, {
      onAuthed: async () => {
        await bootAfterLogin();
      },
    });
    return;
  }

  if (!profile) {
    await loadProfile();
  }

  if (!profile) {
    root.innerHTML = `<div class="auth-wrap"><div class="auth-card"><p>프로필을 불러오지 못했습니다. 잠시 후 새로고침 해주세요.</p></div></div>`;
    return;
  }

  const route = currentRoute();
  if (route.adminOnly && profile.role !== "admin") {
    location.hash = "#/dashboard";
    return;
  }

  root.innerHTML = `
    <div class="layout">
      <aside class="sidebar">
        <div class="brand">CRM</div>
        <nav>
          ${ROUTES.filter((r) => !r.adminOnly || profile.role === "admin")
            .map(
              (r) =>
                `<a href="${r.path}" class="${r.path === route.path ? "nav-link active" : "nav-link"}">${r.label}</a>`
            )
            .join("")}
        </nav>
        <div class="sidebar-footer">
          <div class="who">${escapeHtml(profile.full_name || profile.email)}<br/><span class="muted">${
    profile.role === "admin" ? "관리자" : "일반 직원"
  }</span></div>
          <button class="btn small" id="logout-btn">로그아웃</button>
        </div>
      </aside>
      <main class="content" id="content"></main>
    </div>
  `;
  document.getElementById("logout-btn").addEventListener("click", async () => {
    await supabase.auth.signOut();
    session = null;
    profile = null;
    renderApp();
  });

  const contentEl = document.getElementById("content");
  await route.render(contentEl, { session, profile, navigate: (path) => (location.hash = path) });
}

async function bootAfterLogin() {
  const {
    data: { session: s },
  } = await supabase.auth.getSession();
  session = s;
  await loadProfile();
  if (!location.hash) location.hash = "#/dashboard";
  await renderApp();
}

window.addEventListener("hashchange", renderApp);

supabase.auth.onAuthStateChange((_event, s) => {
  session = s;
});

(async function init() {
  const {
    data: { session: s },
  } = await supabase.auth.getSession();
  session = s;
  if (session) {
    await loadProfile();
  }
  if (!location.hash) location.hash = "#/dashboard";
  await renderApp();
})();
