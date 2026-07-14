import type { User } from '../db/usersRepository.js';
import type { SavedDeclarationSummary, SavedArticleCost } from '../db/declarationsRepository.js';
import type { AppSettings } from '../db/appSettingsRepository.js';
import { renderBrandOverrideStyle, renderLogoImg, FONT_OPTIONS } from './brandingStyles.js';

export type SuperAdminPage = 'dashboard' | 'users' | 'costs' | 'settings';

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDate(iso: string): string {
  const date = new Date(iso);
  return Number.isNaN(date.getTime())
    ? iso
    : date.toLocaleDateString('fr-FR', { year: 'numeric', month: 'short', day: 'numeric' });
}

const NAV_ITEMS: Array<{ page: SuperAdminPage; href: string; label: string; icon: string }> = [
  {
    page: 'dashboard',
    href: '/superadmin/dashboard',
    label: 'Tableau de bord',
    icon: '<path d="M3 10.5l7-6 7 6M5 9v7.5A1.5 1.5 0 0 0 6.5 18h7a1.5 1.5 0 0 0 1.5-1.5V9" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
  },
  {
    page: 'users',
    href: '/superadmin/users',
    label: 'Utilisateurs',
    icon: '<circle cx="7.5" cy="7" r="2.75" stroke="currentColor" stroke-width="1.6"/><path d="M2.5 17c0-2.9 2.24-5 5-5s5 2.1 5 5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><circle cx="14.5" cy="7.5" r="2.1" stroke="currentColor" stroke-width="1.5"/><path d="M13 12.3c1.9.2 3.5 1.9 3.5 4.2" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>',
  },
  {
    page: 'costs',
    href: '/superadmin/costs',
    label: 'Coût de produit',
    icon: '<path d="M10 2.5v15M13.5 5.5c0-1.4-1.6-2-3.5-2s-3.5.9-3.5 2.3c0 3 7 1.4 7 4.3 0 1.4-1.6 2.4-3.5 2.4s-3.5-.9-3.5-2.3" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
  },
  {
    page: 'settings',
    href: '/superadmin/settings',
    label: 'Réglages',
    icon: '<circle cx="10" cy="10" r="2.6" stroke="currentColor" stroke-width="1.5"/><path d="M10 3v1.8M10 15.2V17M17 10h-1.8M4.8 10H3M14.9 5.1l-1.3 1.3M6.4 13.6l-1.3 1.3M14.9 14.9l-1.3-1.3M6.4 6.4L5.1 5.1" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>',
  },
];

function renderSidebar(activePage: SuperAdminPage): string {
  const items = NAV_ITEMS.map(
    (item) => `<a href="${item.href}" class="nav-item${item.page === activePage ? ' active' : ''}">
      <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">${item.icon}</svg>
      <span>${item.label}</span>
    </a>`
  ).join('');

  return `<nav class="sidebar">
    <div class="sidebar-brand">
      <div class="header-badge">
        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 3l7 3v5.2c0 4.6-3 8.7-7 9.8-4-1.1-7-5.2-7-9.8V6l7-3Z" stroke="white" stroke-width="1.6" stroke-linejoin="round"/>
          <path d="M9 12l2 2 4-4.5" stroke="white" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      <div>
        <div class="sidebar-title">Superadmin</div>
        <div class="sidebar-subtitle">Panneau d'administration</div>
      </div>
    </div>
    <div class="nav-items">${items}</div>
    <a class="sidebar-exit" href="/">&larr; Retour à l'outil</a>
  </nav>`;
}

function renderTopbar(title: string, settings: AppSettings): string {
  return `<div class="topbar">
    <h1>${escapeHtml(title)}</h1>
    <div class="topbar-actions">
      ${renderLogoImg(settings)}
      <button class="theme-toggle" id="themeToggle" type="button" aria-label="Changer de thème">
        <svg class="icon-sun" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <circle cx="10" cy="10" r="3.5" stroke="currentColor" stroke-width="1.5"/>
          <path d="M10 2v1.5M10 16.5V18M18 10h-1.5M3.5 10H2M15.5 4.5l-1 1M5.5 14.5l-1 1M15.5 15.5l-1-1M5.5 5.5l-1-1" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
        </svg>
        <svg class="icon-moon" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M17 11.3A7 7 0 1 1 8.7 3a5.5 5.5 0 0 0 8.3 8.3Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
        </svg>
      </button>
      <form method="post" action="/logout">
        <button class="logout-btn" type="submit">
          <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M8 17H5a1.5 1.5 0 0 1-1.5-1.5v-11A1.5 1.5 0 0 1 5 3h3M13.5 14l3.5-4-3.5-4M17 10H7.5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
          Déconnexion
        </button>
      </form>
    </div>
  </div>`;
}

function renderShell(
  activePage: SuperAdminPage,
  title: string,
  bodyHtml: string,
  settings: AppSettings
): string {
  return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Superadmin — ${escapeHtml(title)}</title>
<script>
  (function () {
    var saved = localStorage.getItem('theme');
    var theme = saved || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    document.documentElement.setAttribute('data-theme', theme);
  })();
</script>
<style>
  /* Deliberately violet-accented (vs. the indigo/blue admin tool) so a
     superadmin can tell at a glance which page they're on. */
  :root {
    --ink-900: #0f172a; --ink-700: #334155; --ink-500: #64748b; --ink-400: #94a3b8;
    --line: #e2e8f0; --line-soft: #eef2f7;
    --brand: #6d28d9; --brand-600: #7c3aed; --brand-700: #6d28d9; --brand-soft: #f3e8ff;
    --danger: #b91c1c; --danger-bg: #fef2f2; --danger-line: #fecaca;
    --success: #15803d; --success-bg: #f0fdf4; --success-line: #bbf7d0;
    --warn: #b45309; --warn-bg: #fffbeb; --warn-line: #fde68a;
    --card-bg: #ffffff; --input-bg: #f8fafc; --page-bg: #faf8fc;
    --sidebar-bg: #191325; --sidebar-ink: #e9e4f5; --sidebar-ink-muted: #9c93b5; --sidebar-hover: #2a2140;
  }
  :root[data-theme="dark"] {
    --ink-900: #f1f5f9; --ink-700: #cbd5e1; --ink-500: #94a3b8; --ink-400: #64748b;
    --line: #334155; --line-soft: #1e293b;
    --brand: #c084fc; --brand-600: #a855f7; --brand-700: #c084fc; --brand-soft: #3b0764;
    --danger: #f87171; --danger-bg: #3f1212; --danger-line: #7f1d1d;
    --success: #4ade80; --success-bg: #052e16; --success-line: #14532d;
    --warn: #fbbf24; --warn-bg: #422006; --warn-line: #92400e;
    --card-bg: #0f172a; --input-bg: #1e293b; --page-bg: #0a0414;
    --sidebar-bg: #0f0a1a; --sidebar-ink: #e9e4f5; --sidebar-ink-muted: #7c7295; --sidebar-hover: #1e1730;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: var(--ink-900); background: var(--page-bg); transition: background 0.2s, color 0.2s;
  }
  .app-shell { display: flex; min-height: 100vh; }

  .sidebar {
    width: 232px; flex: none; background: var(--sidebar-bg); color: var(--sidebar-ink);
    padding: 18px 14px; display: flex; flex-direction: column; gap: 4px;
  }
  .sidebar-brand { display: flex; align-items: center; gap: 10px; padding: 4px 8px 22px; }
  .header-badge {
    width: 34px; height: 34px; flex: none; border-radius: 9px;
    background: linear-gradient(135deg, var(--brand-600), var(--brand-700));
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 8px 16px -6px rgba(124, 58, 237, 0.5);
  }
  .header-badge svg { width: 18px; height: 18px; }
  .sidebar-title { font-size: 14px; font-weight: 700; color: #fff; }
  .sidebar-subtitle { font-size: 10.5px; color: var(--sidebar-ink-muted); margin-top: 1px; }
  .nav-items { display: flex; flex-direction: column; gap: 2px; flex: 1; }
  .nav-item {
    display: flex; align-items: center; gap: 10px; padding: 10px 12px; border-radius: 9px;
    color: var(--sidebar-ink-muted); text-decoration: none; font-size: 13.5px; font-weight: 600;
    transition: background 0.12s, color 0.12s;
  }
  .nav-item svg { width: 17px; height: 17px; flex: none; }
  .nav-item:hover { background: var(--sidebar-hover); color: #fff; }
  .nav-item.active { background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); color: #fff; }
  .sidebar-exit { font-size: 12.5px; color: var(--sidebar-ink-muted); text-decoration: none; padding: 10px 12px; }
  .sidebar-exit:hover { color: #fff; }

  .main { flex: 1; min-width: 0; display: flex; flex-direction: column; }
  .topbar {
    display: flex; align-items: center; justify-content: space-between; padding: 18px 28px;
    border-bottom: 1px solid var(--line); background: var(--card-bg); transition: background 0.2s, border-color 0.2s;
  }
  .topbar h1 { font-size: 18px; font-weight: 700; letter-spacing: -0.01em; margin: 0; }
  .topbar-actions { display: flex; align-items: center; gap: 8px; }
  .theme-toggle, .logout-btn {
    height: 36px; border-radius: 9px; border: 1px solid var(--line); background: var(--line-soft);
    display: flex; align-items: center; justify-content: center; cursor: pointer; color: var(--ink-700);
    font-family: inherit; font-size: 12.5px; font-weight: 600; transition: background 0.15s, border-color 0.15s, color 0.15s;
  }
  .theme-toggle { width: 36px; }
  .logout-btn { padding: 0 13px; gap: 6px; }
  .theme-toggle:hover { border-color: var(--brand-600); }
  .logout-btn:hover { border-color: var(--danger-line); color: var(--danger); }
  .theme-toggle svg { width: 17px; height: 17px; }
  .logout-btn svg { width: 15px; height: 15px; }
  .theme-toggle .icon-moon { display: none; }
  :root[data-theme="dark"] .theme-toggle .icon-sun { display: none; }
  :root[data-theme="dark"] .theme-toggle .icon-moon { display: block; }

  .content { padding: 28px; flex: 1; }
  .lede { font-size: 13.5px; color: var(--ink-500); margin: 0 0 22px; }

  .stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 16px; margin-bottom: 24px; }
  .stat-card { border-radius: 14px; padding: 18px 20px; color: #fff; box-shadow: 0 10px 24px -10px rgba(15, 23, 42, 0.35); }
  .stat-card .stat-value { font-size: 26px; font-weight: 800; letter-spacing: -0.02em; }
  .stat-card .stat-label { font-size: 12.5px; font-weight: 600; opacity: 0.92; margin-top: 2px; }
  .stat-card-brand { background: linear-gradient(135deg, #7c3aed, #5b21b6); }
  .stat-card-success { background: linear-gradient(135deg, #10b981, #047857); }
  .stat-card-warn { background: linear-gradient(135deg, #f59e0b, #b45309); }
  .stat-card-danger { background: linear-gradient(135deg, #f43f5e, #be123c); }

  .card {
    background: var(--card-bg); border: 1px solid var(--line); border-radius: 16px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 20px 50px -12px rgba(15, 23, 42, 0.1);
    padding: 24px; margin-bottom: 20px; transition: background 0.2s, border-color 0.2s;
  }
  .card h2 { font-size: 15px; margin: 0 0 16px; }
  table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  th { text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em;
    color: var(--ink-400); font-weight: 600; padding: 0 10px 10px; border-bottom: 1px solid var(--line); }
  td { padding: 12px 10px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
  tr:last-child td { border-bottom: none; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.cost { font-weight: 700; color: var(--brand-700); }
  .muted { color: var(--ink-400); }
  .status { display: inline-block; padding: 3px 9px; border-radius: 999px; font-size: 12px; font-weight: 600; }
  .status-active { background: var(--success-bg); color: var(--success); border: 1px solid var(--success-line); }
  .status-disabled { background: var(--danger-bg); color: var(--danger); border: 1px solid var(--danger-line); }
  .role { display: inline-block; padding: 3px 9px; border-radius: 999px; font-size: 12px; font-weight: 600; }
  .role-superadmin { background: var(--brand-soft); color: var(--brand-700); }
  .role-admin { background: var(--line-soft); color: var(--ink-700); }
  form { display: inline; }
  button {
    font-family: inherit; font-size: 12.5px; font-weight: 600; padding: 7px 12px;
    border-radius: 7px; cursor: pointer; border: 1px solid transparent; transition: filter 0.12s;
  }
  button:hover { filter: brightness(0.96); }
  .btn-disable { background: var(--danger-bg); color: var(--danger); border-color: var(--danger-line); }
  .btn-enable { background: var(--success-bg); color: var(--success); border-color: var(--success-line); }
  .field { margin-bottom: 14px; }
  label { display: block; font-size: 12.5px; font-weight: 600; color: var(--ink-700); margin: 0 0 6px; }
  input, select {
    width: 100%; padding: 10px 12px; font-size: 14px; font-family: inherit; color: var(--ink-900);
    background: var(--input-bg); border: 1.5px solid var(--line); border-radius: 9px;
  }
  input:focus, select:focus { outline: none; background: var(--card-bg); border-color: var(--brand-600); box-shadow: 0 0 0 3.5px rgba(124, 58, 237, 0.16); }
  .create-form { display: grid; grid-template-columns: 1.2fr 1.2fr 1fr auto; gap: 12px; align-items: end; }
  .create-form button {
    padding: 10px 16px; font-size: 13.5px; color: #fff;
    background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border: none;
    box-shadow: 0 6px 16px -4px rgba(124, 58, 237, 0.4);
  }
  .error {
    display: flex; gap: 8px; align-items: flex-start; background: var(--danger-bg); color: var(--danger);
    border: 1px solid var(--danger-line); border-radius: 10px; padding: 11px 13px; margin-bottom: 18px;
    font-size: 13px; line-height: 1.45;
  }
  .error svg { flex: none; margin-top: 1px; }

  .placeholder-card {
    text-align: center; padding: 56px 24px;
  }
  .placeholder-icon {
    width: 56px; height: 56px; border-radius: 50%; margin: 0 auto 16px;
    background: var(--brand-soft); color: var(--brand-700);
    display: flex; align-items: center; justify-content: center;
  }
  .placeholder-icon svg { width: 26px; height: 26px; }
  .placeholder-card h2 { margin: 0 0 6px; }
  .placeholder-card p { color: var(--ink-500); font-size: 13.5px; margin: 0; max-width: 420px; margin: 0 auto; line-height: 1.55; }

  @media (max-width: 900px) {
    .app-shell { flex-direction: column; }
    .sidebar { width: 100%; flex-direction: row; align-items: center; overflow-x: auto; padding: 10px 12px; }
    .sidebar-brand { padding: 0 10px 0 0; }
    .nav-items { flex-direction: row; }
    .sidebar-exit { display: none; }
  }
  @media (max-width: 640px) { .create-form { grid-template-columns: 1fr; } }
</style>
${renderBrandOverrideStyle(settings)}
</head>
<body>
  <div class="app-shell">
    ${renderSidebar(activePage)}
    <div class="main">
      ${renderTopbar(title, settings)}
      <div class="content">${bodyHtml}</div>
    </div>
  </div>
  <script>
    document.getElementById('themeToggle').addEventListener('click', function () {
      var next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('theme', next);
    });
  </script>
</body>
</html>`;
}

function statCard(
  variant: 'brand' | 'success' | 'warn' | 'danger',
  value: string | number,
  label: string
): string {
  return `<div class="stat-card stat-card-${variant}">
    <div class="stat-value">${escapeHtml(String(value))}</div>
    <div class="stat-label">${escapeHtml(label)}</div>
  </div>`;
}

export function renderSuperAdminOverview(
  users: User[],
  declarationCount: number,
  settings: AppSettings
): string {
  const total = users.length;
  const active = users.filter((u) => u.disabledAt === null).length;
  const disabled = users.filter((u) => u.disabledAt !== null).length;
  const superadmins = users.filter((u) => u.role === 'superadmin').length;

  const body = `
    <p class="lede">Vue d'ensemble des comptes et de l'activité de l'application.</p>
    <div class="stat-grid">
      ${statCard('brand', total, 'Comptes au total')}
      ${statCard('success', active, 'Comptes actifs')}
      ${statCard('warn', superadmins, 'Superadmins')}
      ${statCard('danger', disabled, 'Comptes désactivés')}
    </div>
    <div class="stat-grid">
      ${statCard('brand', declarationCount, 'Déclarations générées (total)')}
    </div>
    <div class="card">
      <h2>Accès rapide</h2>
      <p class="lede" style="margin-bottom:0;">Gérez les comptes admin dans <a href="/superadmin/users" style="color:var(--brand-600);font-weight:600;text-decoration:none;">Utilisateurs</a>, ou consultez <a href="/superadmin/costs" style="color:var(--brand-600);font-weight:600;text-decoration:none;">Coût de produit</a>.</p>
    </div>
  `;
  return renderShell('dashboard', 'Tableau de bord', body, settings);
}

function renderUserRow(user: User, currentUserId: number): string {
  const isSelf = user.id === currentUserId;
  const isDisabled = user.disabledAt !== null;
  const statusBadge = isDisabled
    ? '<span class="status status-disabled">Désactivé</span>'
    : '<span class="status status-active">Actif</span>';
  const roleBadge =
    user.role === 'superadmin'
      ? '<span class="role role-superadmin">Superadmin</span>'
      : '<span class="role role-admin">Admin</span>';

  let actionCell: string;
  if (isSelf) {
    actionCell = '<span class="muted">—</span>';
  } else if (isDisabled) {
    actionCell = `<form method="post" action="/superadmin/users/${user.id}/enable"><button type="submit" class="btn-enable">Réactiver</button></form>`;
  } else {
    actionCell = `<form method="post" action="/superadmin/users/${user.id}/disable"><button type="submit" class="btn-disable">Désactiver</button></form>`;
  }

  return `<tr>
    <td>${escapeHtml(user.username)}</td>
    <td>${roleBadge}</td>
    <td>${formatDate(user.createdAt)}</td>
    <td>${statusBadge}</td>
    <td>${actionCell}</td>
  </tr>`;
}

export function renderSuperAdminUsers(
  users: User[],
  currentUserId: number,
  settings: AppSettings,
  errorMessage?: string
): string {
  const rows = users.map((user) => renderUserRow(user, currentUserId)).join('');
  const errorBlock = errorMessage
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(errorMessage)}</span></div>`
    : '';

  const body = `
    <p class="lede">Créez des comptes admin, et activez/désactivez l'accès. L'historique d'un compte désactivé reste intact.</p>
    <div class="card">
      <h2>Nouveau compte</h2>
      ${errorBlock}
      <form class="create-form" method="post" action="/superadmin/users">
        <div class="field">
          <label for="username">Identifiant</label>
          <input type="text" id="username" name="username" placeholder="ex. karim" required />
        </div>
        <div class="field">
          <label for="password">Mot de passe</label>
          <input type="password" id="password" name="password" placeholder="••••••••" required />
        </div>
        <div class="field">
          <label for="role">Rôle</label>
          <select id="role" name="role">
            <option value="admin" selected>Admin</option>
            <option value="superadmin">Superadmin</option>
          </select>
        </div>
        <button type="submit">Créer</button>
      </form>
    </div>
    <div class="card">
      <h2>Comptes (${users.length})</h2>
      <table>
        <thead><tr><th>Identifiant</th><th>Rôle</th><th>Créé le</th><th>Statut</th><th></th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
  return renderShell('users', 'Utilisateurs', body, settings);
}

const PLACEHOLDER_ICON =
  '<path d="M10 6.5v5M10 14.5h.01" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="10" cy="10" r="8" stroke="currentColor" stroke-width="1.6"/>';

// Only used for the "costs" page's empty state now — Users, Dashboard, and
// Settings are all fully real pages; Services was removed from the menu.
export function renderSuperAdminPlaceholder(
  title: string,
  description: string,
  settings: AppSettings
): string {
  const body = `
    <div class="card placeholder-card">
      <div class="placeholder-icon"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">${PLACEHOLDER_ICON}</svg></div>
      <h2>Bientôt disponible</h2>
      <p>${escapeHtml(description)}</p>
    </div>
  `;
  return renderShell('costs', title, body, settings);
}

function formatMoney(value: number): string {
  return value.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Shows the cost breakdown of the most recently *persisted* declaration
// (across all admins — matches the superadmin's "sees everything" role),
// read from the database rather than the admin tool's own in-memory
// last-generated-declaration state. Persisted means this survives a
// redeploy/restart, unlike the earlier in-memory-only version.
export function renderSuperAdminCosts(
  summary: SavedDeclarationSummary,
  articles: SavedArticleCost[],
  settings: AppSettings
): string {
  const rows = articles
    .map(
      (article) => `<tr>
        <td>${escapeHtml(article.nomArticle)}</td>
        <td>${escapeHtml(article.hsCode)}</td>
        <td>${escapeHtml(article.pays)}</td>
        <td class="num">${article.quantite}</td>
        <td class="num cost">${formatMoney(article.costPerUnit)}</td>
      </tr>`
    )
    .join('');

  const partialNote = summary.costEstimatePartial
    ? `<div class="error" style="background:var(--warn-bg);color:var(--warn);border-color:var(--warn-line);">
        <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
        <span>Coût partiel — données d'expédition (fret, assurance, montant facturé) non détectées ou incomplètes ; seuls les droits et taxes sont inclus ci-dessous.</span>
      </div>`
    : '';

  const totalCard = statCard(
    summary.costEstimatePartial ? 'warn' : 'brand',
    formatMoney(summary.totalLandedCost),
    summary.costEstimatePartial ? 'Coût douanier total (partiel)' : 'Coût total estimé'
  );

  const body = `
    <p class="lede">Déclaration ${escapeHtml(summary.code)} — ${escapeHtml(summary.redevable)} (la plus récente générée sur l'application, tous admins confondus).</p>
    ${partialNote}
    <div class="stat-grid">${totalCard}</div>
    <div class="card">
      <h2>Coût par produit</h2>
      <table>
        <thead><tr><th>Produit</th><th>HSC</th><th>Pays</th><th>Qté</th><th>Coût / unité</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;

  return renderShell('costs', 'Coût de produit', body, settings);
}

export function renderSuperAdminSettings(
  settings: AppSettings,
  errorMessage?: string,
  successMessage?: string
): string {
  const errorBlock = errorMessage
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(errorMessage)}</span></div>`
    : '';
  const successBlock = successMessage
    ? `<div class="error" style="background:var(--success-bg);color:var(--success);border-color:var(--success-line);"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(successMessage)}</span></div>`
    : '';

  const currentLogo = settings.logoDataUri
    ? `<div style="margin-bottom:10px;"><img src="${settings.logoDataUri}" alt="Logo actuel" style="height:48px;max-width:220px;object-fit:contain;border-radius:8px;border:1px solid var(--line);padding:6px;" /></div>`
    : '';

  const selectedFont = settings.fontFamily ?? 'system';
  const fontOptions = FONT_OPTIONS.map(
    (opt) =>
      `<option value="${opt.value}"${opt.value === selectedFont ? ' selected' : ''}>${escapeHtml(opt.label)}</option>`
  ).join('');

  const body = `
    <p class="lede">Personnalisez l'identité visuelle de l'application : nom de la société, logo, couleur principale et police.</p>
    ${errorBlock}
    ${successBlock}
    <div class="card">
      <h2>Identité</h2>
      <form method="post" action="/superadmin/settings" enctype="multipart/form-data">
        <div class="field">
          <label for="companyName">Nom de la société</label>
          <input type="text" id="companyName" name="companyName" value="${escapeHtml(settings.companyName ?? '')}" placeholder="ex. Global Trade Logistics SARL" />
        </div>
        <div class="field">
          <label for="logo">Logo</label>
          ${currentLogo}
          <input type="file" id="logo" name="logo" accept="image/png,image/jpeg,image/webp,image/svg+xml" />
        </div>
        <div class="field">
          <label for="brandColor">Couleur principale</label>
          <input type="color" id="brandColor" name="brandColor" value="${escapeHtml(settings.brandColor ?? '#4f46e5')}" style="height:44px;padding:4px;cursor:pointer;" />
        </div>
        <div class="field">
          <label for="fontFamily">Police</label>
          <select id="fontFamily" name="fontFamily">${fontOptions}</select>
        </div>
        <button type="submit">Enregistrer</button>
      </form>
    </div>
  `;

  return renderShell('settings', 'Réglages', body, settings);
}
