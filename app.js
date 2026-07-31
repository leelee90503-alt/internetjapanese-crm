import { supabase } from "./supabaseClient.js";
import { renderAuthScreen } from "./auth.js";
import { renderDashboard } from "./screen-dashboard.js";
import { renderMasterData } from "./screen-masterData.js";
import { renderCustomerUpload } from "./screen-customerUpload.js";
import { renderOrders } from "./screen-orders.js";
import { renderFollowups } from "./screen-followups.js";
import { renderStaff } from "./screen-staff.js";
import { renderCommissionReports } from "./screen-commissionReports.js";
import { renderMissingCommission } from "./screen-missingCommission.js";
import { escapeHtml } from "./normalize.js";

// Sidebar navigation tree. A "link" item is a standalone top-level menu
// entry; a "group" item is a top-level heading that expands to show its
// child links (a sub-menu).
const NAV = [
  { type: "link", path: "#/dashboard", label: "Dashboard", render: renderDashboard, adminOnly: false },
  { type: "link", path: "#/orders", label: "Customer / Order", render: renderOrders, adminOnly: false },
  { type: "link", path: "#/followups", label: "Follow-ups", render: renderFollowups, adminOnly: false },
  {
    type: "group",
    key: "commission",
    label: "Commission",
    children: [
      { path: "#/upload", label: "Customer Excel Upload", render: renderCustomerUpload, adminOnly: false },
      { path: "#/commission-reports", label: "Commission Report", render: renderCommissionReports, adminOnly: true },
      { path: "#/missing-commission", label: "Missing Commission", render: renderMissingCommission, adminOnly: true },
    ],
  },
  {
    type: "group",
    key: "setting",
    label: "Setting",
    children: [
      { path: "#/master-data", label: "Master Data", render: renderMasterData, adminOnly: false },
      { path: "#/staff", label: "Staff Management", render: renderStaff, adminOnly: true },
    ],
  },
];

// Flat list of every actual route (for hash lookup / the admin-only guard),
// derived from NAV so there is a single source of truth for the menu.
const ROUTES = NAV.flatMap((item) => (item.type === "group" ? item.children : [item]));

function groupKeyForPath(path) {
  const group = NAV.find((item) => item.type === "group" && item.children.some((c) => c.path === path));
  return group ? group.key : null;
}

const root = document.getElementById("app");
let session = null;
let profile = null;
let expandedGroups = new Set();

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
    console.error("Failed to load profile", error);
    profile = null;
    return;
  }
  profile = data;
}

function currentRoute() {
  const hash = location.hash || "#/dashboard";
  return ROUTES.find((r) => r.path === hash) || ROUTES[0];
}

function renderNav(route, isAdmin) {
  return NAV.map((item) => {
    if (item.type === "link") {
      if (item.adminOnly && !isAdmin) return "";
      return `<a href="${item.path}" class="${item.path === route.path ? "nav-link active" : "nav-link"}">${item.label}</a>`;
    }
    // group (main menu with sub-menu items)
    const visibleChildren = item.children.filter((c) => !c.adminOnly || isAdmin);
    if (visibleChildren.length === 0) return "";
    const isExpanded = expandedGroups.has(item.key) || visibleChildren.some((c) => c.path === route.path);
    return `
      <div class="nav-group">
        <button type="button" class="nav-group-title" data-group="${item.key}">
          <span>${item.label}</span>
          <span class="nav-group-arrow">${isExpanded ? "&#9662;" : "&#9656;"}</span>
        </button>
        <div class="nav-group-children" ${isExpanded ? "" : 'style="display:none"'}>
          ${visibleChildren
            .map(
              (c) =>
                `<a href="${c.path}" class="${c.path === route.path ? "nav-link active" : "nav-link"}">${c.label}</a>`
            )
            .join("")}
        </div>
      </div>`;
  }).join("");
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
    root.innerHTML = `<div class="auth-wrap"><div class="auth-card"><p>Could not load your profile. Please refresh in a moment.</p></div></div>`;
    return;
  }

  const route = currentRoute();
  if (route.adminOnly && profile.role !== "admin") {
    location.hash = "#/dashboard";
    return;
  }

  const isAdmin = profile.role === "admin";
  const activeGroupKey = groupKeyForPath(route.path);
  if (activeGroupKey) expandedGroups.add(activeGroupKey);

  root.innerHTML = `
    <div class="layout">
      <aside class="sidebar">
        <div class="brand">CRM</div>
        <nav>
          ${renderNav(route, isAdmin)}
        </nav>
        <div class="sidebar-footer">
          <div class="who">${escapeHtml(profile.full_name || profile.email)}<br/><span class="muted">${
    isAdmin ? "Admin" : "Staff"
  }</span></div>
          <button class="btn small" id="logout-btn">Log Out</button>
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

  root.querySelectorAll(".nav-group-title").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.dataset.group;
      if (expandedGroups.has(key)) expandedGroups.delete(key);
      else expandedGroups.add(key);
      renderApp();
    });
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
