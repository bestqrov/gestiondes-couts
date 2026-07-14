import type { User } from '../db/usersRepository.js';
import type { TransactionDocument, CountryProductCount } from '../db/transactionsRepository.js';
import type { AppSettings } from '../db/appSettingsRepository.js';
import { renderBrandOverrideStyle, renderLogoImg, renderFaviconLink, FONT_OPTIONS } from './brandingStyles.js';
import { renderWorldMapPanel, WORLD_MAP_STYLE } from './worldMap.js';

export type SuperAdminPage = 'dashboard' | 'generate' | 'users' | 'costs' | 'settings';

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
    page: 'generate',
    href: '/superadmin/generate',
    label: 'Générer une déclaration',
    icon: '<path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>',
  },
  {
    page: 'users',
    href: '/superadmin/users',
    label: 'Utilisateurs',
    icon: '<circle cx="7.5" cy="7" r="2.75" stroke="currentColor" stroke-width="1.6"/><path d="M2.5 17c0-2.9 2.24-5 5-5s5 2.1 5 5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><circle cx="14.5" cy="7.5" r="2.1" stroke="currentColor" stroke-width="1.5"/><path d="M13 12.3c1.9.2 3.5 1.9 3.5 4.2" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>',
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
    <form method="post" action="/logout">
      <button class="sidebar-logout" type="submit">
        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M8 17H5a1.5 1.5 0 0 1-1.5-1.5v-11A1.5 1.5 0 0 1 5 3h3M13.5 14l3.5-4-3.5-4M17 10H7.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
        Déconnexion
      </button>
    </form>
  </nav>`;
}

function renderTopbar(title: string, settings: AppSettings): string {
  const companyName = settings.companyName
    ? `<div class="topbar-company">${escapeHtml(settings.companyName)}</div>`
    : '';
  return `<div class="topbar">
    <h1>${escapeHtml(title)}</h1>
    ${companyName}
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
    </div>
  </div>`;
}

function renderShell(
  activePage: SuperAdminPage,
  title: string,
  bodyHtml: string,
  settings: AppSettings,
  extraStyle = ''
): string {
  return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Superadmin — ${escapeHtml(title)}</title>
${renderFaviconLink(settings)}
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
  html, body { margin: 0; padding: 0; height: 100%; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: var(--ink-900); background: var(--page-bg); transition: background 0.2s, color 0.2s;
  }
  .app-shell { display: flex; height: 100vh; }

  /* Sticky within the full-height flex row above, so it stays in view
     while .content (below) scrolls independently — the "fixed sidebar"
     look without the margin-offset bookkeeping position:fixed would need. */
  .sidebar {
    position: sticky; top: 0; height: 100vh; overflow-y: auto;
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

  .sidebar-logout {
    width: 100%; display: flex; align-items: center; gap: 10px; padding: 12px 14px;
    margin-top: 8px; border-radius: 10px; border: 1px solid rgba(248, 113, 113, 0.25);
    background: rgba(248, 113, 113, 0.1); color: #fca5a5;
    font-family: inherit; font-size: 14px; font-weight: 700; cursor: pointer;
    transition: background 0.15s, border-color 0.15s, color 0.15s;
  }
  .sidebar-logout svg { width: 19px; height: 19px; flex: none; }
  .sidebar-logout:hover { background: rgba(248, 113, 113, 0.2); border-color: rgba(248, 113, 113, 0.45); color: #fff; }

  .main { flex: 1; min-width: 0; height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
  .topbar {
    position: relative; flex: none;
    display: flex; align-items: center; justify-content: space-between; padding: 18px 28px;
    border-bottom: 1px solid var(--line); background: var(--header-bg, var(--card-bg));
    transition: background 0.2s, border-color 0.2s;
  }
  .topbar h1 { font-size: 18px; font-weight: 700; letter-spacing: -0.01em; margin: 0; color: var(--header-ink, var(--ink-900)); }
  .topbar-company {
    position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%);
    font-size: 14.5px; font-weight: 700; color: var(--header-ink, var(--ink-900)); white-space: nowrap;
    max-width: 40%; overflow: hidden; text-overflow: ellipsis; pointer-events: none;
  }
  .topbar-actions { display: flex; align-items: center; gap: 8px; }
  .theme-toggle {
    height: 36px; width: 36px; border-radius: 9px; border: 1px solid var(--line); background: var(--line-soft);
    display: flex; align-items: center; justify-content: center; cursor: pointer; color: var(--ink-700);
    font-family: inherit; font-size: 12.5px; font-weight: 600; transition: background 0.15s, border-color 0.15s, color 0.15s;
  }
  .theme-toggle:hover { border-color: var(--brand-600); }
  .theme-toggle svg { width: 17px; height: 17px; }
  .theme-toggle .icon-moon { display: none; }
  :root[data-theme="dark"] .theme-toggle .icon-sun { display: none; }
  :root[data-theme="dark"] .theme-toggle .icon-moon { display: block; }

  .content { padding: 28px; flex: 1; overflow-y: auto; }
  .lede { font-size: 13.5px; color: var(--ink-500); margin: 0 0 22px; }

  .stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 16px; margin-bottom: 24px; }
  .stat-card { border-radius: 14px; padding: 18px 20px; color: #fff; box-shadow: 0 10px 24px -10px rgba(15, 23, 42, 0.35); }
  .stat-card .stat-value { font-size: 26px; font-weight: 800; letter-spacing: -0.02em; }
  .stat-card .stat-label { font-size: 12.5px; font-weight: 600; opacity: 0.92; margin-top: 2px; }
  .stat-card-brand { background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); }
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
    .app-shell { flex-direction: column; height: auto; }
    .sidebar {
      position: static; width: 100%; height: auto; flex-direction: row; align-items: center;
      overflow-x: auto; overflow-y: visible; padding: 10px 12px;
    }
    .sidebar-brand { padding: 0 10px 0 0; }
    .nav-items { flex-direction: row; }
    .sidebar-logout { width: auto; margin-top: 0; white-space: nowrap; }
    .main { height: auto; overflow: visible; }
    .content { overflow-y: visible; }
  }
</style>
<style>${extraStyle}</style>
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
  settings: AppSettings,
  countryCounts: CountryProductCount[]
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
    ${renderWorldMapPanel(countryCounts)}
    <div class="card">
      <h2>Accès rapide</h2>
      <p class="lede" style="margin-bottom:0;">Gérez les comptes admin dans <a href="/superadmin/users" style="color:var(--brand-600);font-weight:600;text-decoration:none;">Utilisateurs</a>.</p>
    </div>
  `;
  return renderShell('dashboard', 'Tableau de bord', body, settings, WORLD_MAP_STYLE);
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
    <p class="lede">Comptes de l'application. L'historique d'un compte désactivé reste intact.</p>
    ${errorBlock}
    <div class="card settings-section">
      <div class="settings-section-head">
        <div class="settings-icon settings-icon-violet"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="10" cy="7" r="3" stroke="currentColor" stroke-width="1.4"/><path d="M4 17c0-3.3 2.7-6 6-6s6 2.7 6 6" stroke="currentColor" stroke-width="1.4"/></svg></div>
        <div>
          <h2>Nouveau compte</h2>
          <p class="settings-section-hint">Accès limité à « Générer une déclaration » — pas de tableau de bord ni de réglages.</p>
        </div>
      </div>
      <form method="post" action="/superadmin/users">
        <input type="hidden" name="role" value="admin" />
        <div class="settings-field-grid">
          <div class="field">
            <label for="newUserUsername">Nom d'utilisateur</label>
            <input type="text" id="newUserUsername" name="username" required />
          </div>
          <div class="field">
            <label for="newUserPassword">Mot de passe</label>
            <input type="password" id="newUserPassword" name="password" required autocomplete="new-password" />
          </div>
        </div>
        <button type="submit" class="settings-save">
          <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 4.5v11M4.5 10h11" stroke="white" stroke-width="1.8" stroke-linecap="round"/></svg>
          Créer le compte
        </button>
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
  return renderShell('users', 'Utilisateurs', body, settings, SETTINGS_PAGE_STYLE);
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

function renderCostsSearchForm(searchQuery: string): string {
  const clearLink = searchQuery
    ? `<a href="/superadmin/costs" class="search-clear">Effacer</a>`
    : '';
  return `
    <form method="get" action="/superadmin/costs" class="search-form">
      <input type="text" name="q" placeholder="Rechercher par nom / société (redevable)..." value="${escapeHtml(searchQuery)}" />
      <button type="submit" class="search-submit">Rechercher</button>
      ${clearLink}
    </form>
  `;
}

function renderSearchResultCard(result: TransactionDocument): string {
  const valeurDeclaree =
    result.valeurTotaleDeclaree !== null ? formatMoney(result.valeurTotaleDeclaree) : '—';
  return `
    <div class="card search-result-card">
      <div class="search-result-header">
        <strong>${escapeHtml(result.redevable)}</strong>
        <span class="muted">${escapeHtml(result.code)} · ${formatDate(result.createdAt)}</span>
      </div>
      <div class="search-result-footer">
        <div class="search-stat"><span class="label">Coût total</span><span class="value">${formatMoney(result.totalLandedCost)}</span></div>
        <div class="search-stat"><span class="label">Total des taxes</span><span class="value">${formatMoney(result.totalTaxes)}</span></div>
        <div class="search-stat"><span class="label">Valeur déclarée</span><span class="value">${valeurDeclaree}</span></div>
      </div>
    </div>
  `;
}

// Shows the cost breakdown of the most recently *persisted* transaction
// (across all admins — matches the superadmin's "sees everything" role),
// read from MongoDB rather than the admin tool's own in-memory
// last-generated-declaration state. Persisted means this survives a
// redeploy/restart, unlike the earlier in-memory-only version.
//
// A search box (by redevable/company name, across every persisted
// transaction) sits above this — when a query is active, its results
// (each a compact header+footer card with the three totals) replace the
// single most-recent detail view below.
export function renderSuperAdminCosts(
  mostRecent: TransactionDocument,
  settings: AppSettings,
  searchQuery = '',
  searchResults?: TransactionDocument[]
): string {
  const searchForm = renderCostsSearchForm(searchQuery);

  if (searchQuery && searchResults) {
    const resultsHtml =
      searchResults.length > 0
        ? searchResults.map(renderSearchResultCard).join('')
        : `<div class="card placeholder-card"><p>Aucune déclaration ne correspond à « ${escapeHtml(searchQuery)} ».</p></div>`;
    const body = `
      <p class="lede">Recherche de déclarations par nom / société.</p>
      ${searchForm}
      ${resultsHtml}
    `;
    return renderShell('costs', 'Coût de produit', body, settings, COSTS_SEARCH_STYLE);
  }

  const rows = mostRecent.articles
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

  const partialNote = mostRecent.costEstimatePartial
    ? `<div class="error" style="background:var(--warn-bg);color:var(--warn);border-color:var(--warn-line);">
        <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
        <span>Coût partiel — données d'expédition (fret, assurance, montant facturé) non détectées ou incomplètes ; seuls les droits et taxes sont inclus ci-dessous.</span>
      </div>`
    : '';

  const totalCard = statCard(
    mostRecent.costEstimatePartial ? 'warn' : 'brand',
    formatMoney(mostRecent.totalLandedCost),
    mostRecent.costEstimatePartial ? 'Coût douanier total (partiel)' : 'Coût total estimé'
  );

  const body = `
    <p class="lede">Déclaration ${escapeHtml(mostRecent.code)} — ${escapeHtml(mostRecent.redevable)} (la plus récente générée sur l'application, tous admins confondus).</p>
    ${searchForm}
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

  return renderShell('costs', 'Coût de produit', body, settings, COSTS_SEARCH_STYLE);
}

const COSTS_SEARCH_STYLE = `
  .search-form { display: flex; gap: 10px; margin-bottom: 20px; }
  .search-form input {
    flex: 1; padding: 10px 12px; font-size: 14px; font-family: inherit; color: var(--ink-900);
    background: var(--input-bg); border: 1.5px solid var(--line); border-radius: 9px;
  }
  .search-form input:focus { outline: none; border-color: var(--brand-600); background: var(--card-bg); }
  .search-submit {
    width: auto; margin-top: 0; padding: 0 18px; font-size: 13.5px; color: #fff;
    background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border: none; border-radius: 9px;
  }
  .search-clear {
    display: flex; align-items: center; padding: 0 14px; font-size: 13px; color: var(--ink-500);
    text-decoration: none; border: 1.5px solid var(--line); border-radius: 9px;
  }
  .search-clear:hover { color: var(--danger); border-color: var(--danger-line); }
  .search-result-card { padding: 0; overflow: hidden; }
  .search-result-header {
    padding: 16px 20px; display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap;
    font-size: 14.5px; border-bottom: 1px solid var(--line-soft);
  }
  .search-result-footer { display: flex; }
  .search-stat {
    flex: 1; padding: 14px 20px; display: flex; flex-direction: column; gap: 3px;
  }
  .search-stat:not(:last-child) { border-right: 1px solid var(--line-soft); }
  .search-stat .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.03em; color: var(--ink-400); font-weight: 600; }
  .search-stat .value { font-size: 15px; font-weight: 700; color: var(--brand-700); }
  @media (max-width: 560px) {
    .search-form { flex-wrap: wrap; }
    .search-result-footer { flex-direction: column; }
    .search-stat:not(:last-child) { border-right: none; border-bottom: 1px solid var(--line-soft); }
  }
`;

export function renderSuperAdminSettings(
  settings: AppSettings,
  errorMessage?: string,
  successMessage?: string,
  currentUsername?: string,
  credentialsError?: string,
  credentialsSuccess?: string
): string {
  const errorBlock = errorMessage
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(errorMessage)}</span></div>`
    : '';
  const successBlock = successMessage
    ? `<div class="error" style="background:var(--success-bg);color:var(--success);border-color:var(--success-line);"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(successMessage)}</span></div>`
    : '';
  const credentialsErrorBlock = credentialsError
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(credentialsError)}</span></div>`
    : '';
  const credentialsSuccessBlock = credentialsSuccess
    ? `<div class="error" style="background:var(--success-bg);color:var(--success);border-color:var(--success-line);"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(credentialsSuccess)}</span></div>`
    : '';

  const currentLogo = settings.logoDataUri
    ? `<img src="${settings.logoDataUri}" alt="Logo actuel" class="logo-preview" />`
    : `<div class="logo-preview logo-preview-empty"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 6h12v9a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6Z" stroke="currentColor" stroke-width="1.4"/><path d="M4 6l2.5-2.5h7L16 6" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><circle cx="10" cy="11" r="2" stroke="currentColor" stroke-width="1.3"/></svg></div>`;

  const selectedFont = settings.fontFamily ?? 'system';
  const fontOptions = FONT_OPTIONS.map(
    (opt) =>
      `<option value="${opt.value}"${opt.value === selectedFont ? ' selected' : ''}>${escapeHtml(opt.label)}</option>`
  ).join('');

  const defaultTab = credentialsError || credentialsSuccess ? 'identifiants' : 'profil';

  const body = `
    <p class="lede">Personnalisez l'identité visuelle de l'application : nom de la société, logo, couleur principale, police, coordonnées et identifiants de connexion.</p>

    <div class="settings-tabs" role="tablist">
      <button type="button" class="settings-tab" data-tab="profil">
        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 17V5.5A1.5 1.5 0 0 1 5.5 4h9A1.5 1.5 0 0 1 16 5.5V17" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M7.5 8h1M11.5 8h1M7.5 11h1M11.5 11h1M8.5 17v-3h3v3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
        Profil Société
      </button>
      <button type="button" class="settings-tab" data-tab="apparence">
        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="10" cy="10" r="6.5" stroke="currentColor" stroke-width="1.5"/><path d="M10 3.5a6.5 6.5 0 0 1 0 13" stroke="currentColor" stroke-width="1.5"/></svg>
        Thème
      </button>
      <button type="button" class="settings-tab" data-tab="contact">
        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 5.5h14v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-9Z" stroke="currentColor" stroke-width="1.4"/><path d="M3.5 5.8l6.5 5 6.5-5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg>
        Contact
      </button>
      <button type="button" class="settings-tab" data-tab="identifiants">
        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="10" cy="7" r="3" stroke="currentColor" stroke-width="1.4"/><path d="M4 17c0-3.3 2.7-6 6-6s6 2.7 6 6" stroke="currentColor" stroke-width="1.4"/></svg>
        Identifiants
      </button>
    </div>

    <form method="post" action="/superadmin/settings" enctype="multipart/form-data">
      <div class="tab-panel" data-panel="profil">
        ${errorBlock}
        ${successBlock}
        <div class="card settings-section logo-card">
          <div class="settings-section-head">
            <div class="settings-icon settings-icon-violet"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 6h12v9a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6Z" stroke="currentColor" stroke-width="1.3"/></svg></div>
            <div>
              <h2>Logo de la société</h2>
              <p class="settings-section-hint">Téléchargez le logo de votre établissement. Il sera affiché dans la barre latérale et sur tous les documents.</p>
            </div>
          </div>
          <div class="logo-field">
            <span id="logoPreviewWrap">${currentLogo}</span>
            <div class="logo-field-input">
              <div class="logo-field-buttons">
                <label for="logo" class="btn-upload">
                  <svg width="14" height="14" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 13V4M6.5 7.5 10 4l3.5 3.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/><path d="M4 14v1.5A1.5 1.5 0 0 0 5.5 17h9a1.5 1.5 0 0 0 1.5-1.5V14" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
                  Télécharger un logo
                </label>
                <input type="file" id="logo" name="logo" accept="image/png,image/jpeg,image/webp,image/svg+xml" style="display:none;" />
                <button type="button" id="removeLogoBtn" class="btn-remove-logo"${settings.logoDataUri ? '' : ' style="display:none;"'}>Supprimer</button>
                <input type="hidden" id="removeLogo" name="removeLogo" value="" />
              </div>
              <span class="field-hint">Formats acceptés : PNG, JPEG, WEBP ou SVG — 2 Mo maximum.</span>
            </div>
          </div>
        </div>

        <div class="card settings-section">
          <div class="settings-section-head">
            <div class="settings-icon settings-icon-violet"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 17V5.5A1.5 1.5 0 0 1 5.5 4h9A1.5 1.5 0 0 1 16 5.5V17" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M7.5 8h1M11.5 8h1M7.5 11h1M11.5 11h1M8.5 17v-3h3v3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg></div>
            <div>
              <h2>Société</h2>
              <p class="settings-section-hint">Nom affiché dans l'application et sur la page de connexion.</p>
            </div>
          </div>
          <div class="settings-field-grid">
            <div class="field">
              <label for="companyName">Nom de la société</label>
              <input type="text" id="companyName" name="companyName" value="${escapeHtml(settings.companyName ?? '')}" placeholder="ex. Global Trade Logistics SARL" />
            </div>
          </div>
        </div>
      </div>

      <div class="tab-panel" data-panel="apparence" hidden>
        <div class="card settings-section">
          <div class="settings-section-head">
            <div class="settings-icon settings-icon-amber"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="10" cy="10" r="6.5" stroke="currentColor" stroke-width="1.5"/><path d="M10 3.5a6.5 6.5 0 0 1 0 13" stroke="currentColor" stroke-width="1.5"/></svg></div>
            <div>
              <h2>Apparence</h2>
              <p class="settings-section-hint">Couleur d'accent et police utilisées dans toute l'application.</p>
            </div>
          </div>
          <div class="settings-field-grid">
            <div class="field">
              <label for="brandColor">Couleur principale</label>
              <div class="color-field">
                <input type="color" id="brandColor" name="brandColor" value="${escapeHtml(settings.brandColor ?? '#4f46e5')}" />
                <span class="color-field-hex" id="brandColorHex">${escapeHtml(settings.brandColor ?? '#4f46e5')}</span>
              </div>
            </div>
            <div class="field">
              <label for="fontFamily">Police</label>
              <select id="fontFamily" name="fontFamily">${fontOptions}</select>
            </div>
          </div>
        </div>
      </div>

      <div class="tab-panel" data-panel="contact" hidden>
        <div class="card settings-section">
          <div class="settings-section-head">
            <div class="settings-icon settings-icon-green"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 5.5h14v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-9Z" stroke="currentColor" stroke-width="1.4"/><path d="M3.5 5.8l6.5 5 6.5-5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/></svg></div>
            <div>
              <h2>Contact</h2>
              <p class="settings-section-hint">Affichés dans le pied de page de l'écran de connexion.</p>
            </div>
          </div>
          <div class="settings-field-grid">
            <div class="field">
              <label for="contactEmail">Email de contact</label>
              <input type="email" id="contactEmail" name="contactEmail" value="${escapeHtml(settings.contactEmail ?? '')}" placeholder="ex. contact@societe.com" />
            </div>
            <div class="field">
              <label for="contactWhatsapp">Téléphone / WhatsApp</label>
              <input type="text" id="contactWhatsapp" name="contactWhatsapp" value="${escapeHtml(settings.contactWhatsapp ?? '')}" placeholder="ex. +212 6 00 00 00 00" />
            </div>
          </div>
        </div>
      </div>

      <button type="submit" class="settings-save">
        <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
        Enregistrer
      </button>
    </form>

    <div class="tab-panel" data-panel="identifiants" hidden>
      <div class="card settings-section">
        <div class="settings-section-head">
          <div class="settings-icon settings-icon-violet"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="4.5" y="9" width="11" height="7.5" rx="1.3" stroke="currentColor" stroke-width="1.4"/><path d="M6.5 9V6.5a3.5 3.5 0 0 1 7 0V9" stroke="currentColor" stroke-width="1.4"/></svg></div>
          <div>
            <h2>Identifiants de connexion</h2>
            <p class="settings-section-hint">Nom d'utilisateur et mot de passe de votre propre compte superadmin.</p>
          </div>
        </div>
        ${credentialsErrorBlock}
        ${credentialsSuccessBlock}
        <form method="post" action="/superadmin/settings/credentials">
          <div class="settings-field-grid">
            <div class="field">
              <label for="username">Nom d'utilisateur</label>
              <input type="text" id="username" name="username" value="${escapeHtml(currentUsername ?? '')}" required />
            </div>
            <div class="field">
              <label for="newPassword">Nouveau mot de passe</label>
              <input type="password" id="newPassword" name="newPassword" placeholder="Laisser vide pour ne pas changer" autocomplete="new-password" />
            </div>
          </div>
          <button type="submit" class="settings-save">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Mettre à jour les identifiants
          </button>
        </form>
      </div>
    </div>

    <script>
      var brandColorInput = document.getElementById('brandColor');
      var brandColorHex = document.getElementById('brandColorHex');
      if (brandColorInput && brandColorHex) {
        brandColorInput.addEventListener('input', function () {
          brandColorHex.textContent = brandColorInput.value;
        });
      }

      var removeLogoBtn = document.getElementById('removeLogoBtn');
      var removeLogoInput = document.getElementById('removeLogo');
      var logoPreviewWrap = document.getElementById('logoPreviewWrap');
      if (removeLogoBtn && removeLogoInput && logoPreviewWrap) {
        removeLogoBtn.addEventListener('click', function () {
          removeLogoInput.value = '1';
          logoPreviewWrap.innerHTML = '<div class="logo-preview logo-preview-empty"><svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 6h12v9a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6Z" stroke="currentColor" stroke-width="1.4"/><path d="M4 6l2.5-2.5h7L16 6" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><circle cx="10" cy="11" r="2" stroke="currentColor" stroke-width="1.3"/></svg></div>';
          removeLogoBtn.style.display = 'none';
        });
      }

      var tabs = Array.prototype.slice.call(document.querySelectorAll('.settings-tab'));
      var panels = Array.prototype.slice.call(document.querySelectorAll('.tab-panel'));
      function activateTab(name) {
        tabs.forEach(function (tab) {
          tab.classList.toggle('active', tab.getAttribute('data-tab') === name);
        });
        panels.forEach(function (panel) {
          panel.hidden = panel.getAttribute('data-panel') !== name;
        });
      }
      tabs.forEach(function (tab) {
        tab.addEventListener('click', function () {
          activateTab(tab.getAttribute('data-tab'));
        });
      });
      activateTab('${defaultTab}');
    </script>
  `;

  return renderShell('settings', 'Réglages', body, settings, SETTINGS_PAGE_STYLE);
}

const SETTINGS_PAGE_STYLE = `
  .settings-tabs {
    display: flex; gap: 4px; border-bottom: 1px solid var(--line); margin-bottom: 24px; overflow-x: auto;
  }
  .settings-tab {
    display: flex; align-items: center; gap: 8px; padding: 12px 18px; font-family: inherit;
    font-size: 14px; font-weight: 700; color: var(--ink-500); background: none; border: none;
    border-bottom: 2.5px solid transparent; cursor: pointer; white-space: nowrap;
    transition: color 0.15s, border-color 0.15s, background 0.15s;
  }
  .settings-tab svg { width: 17px; height: 17px; flex: none; }
  .settings-tab:hover { color: var(--ink-700); }
  .settings-tab.active { color: var(--brand-600); border-bottom-color: var(--brand-600); background: var(--brand-soft); border-radius: 8px 8px 0 0; }
  .tab-panel[hidden] { display: none; }

  .logo-card {
    background: linear-gradient(135deg, var(--brand-soft), var(--card-bg) 70%);
  }
  .logo-field { display: flex; align-items: flex-start; gap: 20px; }
  .logo-field-buttons { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 6px; }
  .btn-upload {
    display: inline-flex; align-items: center; gap: 7px; padding: 10px 16px; font-size: 13.5px; font-weight: 700;
    color: #fff; background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border-radius: 9px;
    cursor: pointer; border: none;
  }
  .btn-upload:hover { filter: brightness(1.05); }
  .btn-remove-logo {
    padding: 10px 16px; font-size: 13.5px; font-weight: 700; color: var(--danger);
    background: var(--danger-bg); border: 1px solid var(--danger-line); border-radius: 9px; cursor: pointer;
  }
  .btn-remove-logo:hover { filter: brightness(0.97); }

  .settings-field-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px 20px; }

  .settings-section-head { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 18px; }
  .settings-section-head h2 { margin: 0 0 3px; }
  .settings-section-hint { margin: 0; font-size: 12px; color: var(--ink-500); line-height: 1.4; }
  .settings-icon {
    width: 34px; height: 34px; flex: none; border-radius: 9px;
    display: flex; align-items: center; justify-content: center;
  }
  .settings-icon svg { width: 18px; height: 18px; }
  .settings-icon-violet, .settings-icon-amber, .settings-icon-green {
    background: var(--brand-soft); color: var(--brand-700);
  }

  label svg { width: 13px; height: 13px; vertical-align: -1px; margin-right: 4px; color: var(--ink-400); }

  .logo-field { display: flex; align-items: center; gap: 14px; }
  .logo-field-input { flex: 1; display: flex; flex-direction: column; gap: 6px; }
  .field-hint { font-size: 11.5px; color: var(--ink-400); }
  .logo-preview {
    width: 96px; height: 96px; flex: none; border-radius: 14px; object-fit: contain;
    border: 1.5px dashed var(--line); padding: 8px; background: var(--card-bg);
  }
  .logo-preview-empty { display: flex; align-items: center; justify-content: center; color: var(--ink-400); }
  .logo-preview-empty svg { width: 24px; height: 24px; }

  .color-field { display: flex; align-items: center; gap: 10px; }
  .color-field input[type="color"] { width: 52px; height: 44px; padding: 4px; cursor: pointer; flex: none; }
  .color-field-hex {
    font-family: 'Courier New', monospace; font-size: 13px; font-weight: 600; color: var(--ink-700);
    background: var(--input-bg); border: 1px solid var(--line); border-radius: 8px; padding: 9px 12px;
  }

  .settings-save {
    width: 100%; padding: 13px; font-size: 14px; font-weight: 700; color: #fff;
    background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border: none; border-radius: 10px;
    cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 8px;
    box-shadow: 0 8px 18px -6px rgba(124, 58, 237, 0.4);
  }
  .settings-save:hover { filter: brightness(1.05); }
`;

// The visual language here (big drop zones, full-width submit button, the
// success/cost/results panels) deliberately overrides the shell's compact
// admin-form styling — those rules live in renderShell's shared <style>
// too, but this extraStyle block is appended after it in the same
// document, so it wins for this page only without affecting the others.
const GENERATE_PAGE_STYLE = `
  .drop-zone {
    position: relative; border: 1.5px dashed var(--line); border-radius: 13px;
    padding: 22px 20px; margin-bottom: 14px; cursor: pointer; background: var(--input-bg);
    transition: border-color 0.15s, background 0.15s, box-shadow 0.15s;
    display: flex; gap: 14px; align-items: flex-start;
  }
  .drop-zone:hover { border-color: var(--brand-600); background: var(--brand-soft); }
  .drop-zone.dragover { border-color: var(--brand-600); background: var(--brand-soft); box-shadow: 0 0 0 3.5px rgba(124, 58, 237, 0.14); }
  .drop-zone.filled { border-style: solid; border-color: var(--success); background: var(--success-bg); }
  .drop-zone input { display: none; }
  .dz-icon {
    flex: none; width: 38px; height: 38px; border-radius: 10px; background: var(--brand-soft);
    display: flex; align-items: center; justify-content: center; color: var(--brand-600);
    transition: background 0.15s, color 0.15s;
  }
  .drop-zone.filled .dz-icon { background: var(--success-bg); color: var(--success); }
  .dz-icon svg { width: 19px; height: 19px; }
  .dz-body { min-width: 0; flex: 1; }
  .drop-zone .label { font-weight: 600; font-size: 14.5px; color: var(--ink-900); }
  .drop-zone .hint { font-size: 12.5px; color: var(--ink-500); margin-top: 2px; }
  .drop-zone .filename {
    margin-top: 7px; font-size: 12.5px; font-weight: 600; color: var(--success);
    display: flex; align-items: center; gap: 5px; overflow-wrap: anywhere;
  }
  #generateForm button[type="submit"] {
    width: 100%; margin-top: 8px; padding: 13px; font-size: 14.5px; font-weight: 600;
    color: #fff; background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border: none;
    box-shadow: 0 6px 16px -4px rgba(124, 58, 237, 0.45);
    transition: transform 0.12s, box-shadow 0.12s, filter 0.12s; display: flex; align-items: center; justify-content: center; gap: 8px;
  }
  #generateForm button[type="submit"]:hover:not(:disabled) { filter: brightness(1.06); }
  #generateForm button[type="submit"]:disabled { background: var(--line); box-shadow: none; cursor: not-allowed; }
  #status { margin-top: 14px; font-size: 13px; color: var(--ink-500); text-align: center; min-height: 18px; }
  #status.busy { color: var(--brand-600); font-weight: 600; }

  #successPanel { display: none; }
  #successPanel.visible { display: block; }
  .success-icon {
    width: 56px; height: 56px; border-radius: 50%; background: var(--success-bg); border: 1.5px solid var(--success-line);
    display: flex; align-items: center; justify-content: center; margin: 4px auto 16px;
  }
  .success-icon svg { width: 26px; height: 26px; }
  .success-title { text-align: center; font-size: 17px; font-weight: 700; color: var(--ink-900); }
  .success-subtitle { text-align: center; font-size: 13px; color: var(--ink-500); margin-top: 4px; margin-bottom: 22px; }
  .success-actions { display: flex; gap: 10px; flex-wrap: wrap; }
  .success-actions button { flex: 1; min-width: 130px; margin-top: 0; padding: 10px 14px; }
  button.secondary {
    background: var(--card-bg); color: var(--ink-700); border: 1.5px solid var(--line); box-shadow: none;
  }
  button.secondary:hover:not(:disabled) { background: var(--input-bg); border-color: var(--brand-600); }
  button.ghost { background: transparent; color: var(--ink-500); box-shadow: none; font-weight: 500; font-size: 13px; padding: 9px; width: auto; }
  button.ghost:hover { background: var(--line-soft); }

  .cost-panel { margin-top: 18px; border-top: 1px solid var(--line-soft); padding-top: 18px; display: none; }
  .cost-panel.visible { display: block; }
  .cost-panel .note {
    font-size: 12px; color: var(--warn); background: var(--warn-bg); border: 1px solid var(--warn-line);
    border-radius: 8px; padding: 9px 11px; margin-bottom: 12px; line-height: 1.5;
  }

  #successPanel.cost-exclusive .success-icon,
  #successPanel.cost-exclusive .success-title,
  #successPanel.cost-exclusive .success-subtitle,
  #successPanel.cost-exclusive .results-section,
  #successPanel.cost-exclusive #newDeclarationBtn,
  #successPanel.cost-exclusive #downloadAgainBtn,
  #successPanel.cost-exclusive #showResultsBtn { display: none; }
  #successPanel.cost-exclusive .cost-panel { border-top: none; margin-top: 0; }
  table.cost-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  table.cost-table th, table.cost-table td { text-align: center; padding: 7px 9px; border-bottom: 1px solid var(--line-soft); }
  table.cost-table td.num { text-align: center; font-variant-numeric: tabular-nums; }
  table.cost-table td.cost { font-weight: 700; color: var(--brand-700); }

  .results-section { margin-top: 18px; border-top: 1px solid var(--line-soft); padding-top: 18px; display: none; }
  .results-section.visible { display: block; }
  .results-toolbar { display: flex; gap: 10px; margin-bottom: 16px; }
  .results-toolbar button { width: auto; margin-top: 0; padding: 9px 14px; font-size: 12.5px; }
  .results-heading { font-size: 14px; margin: 0 0 14px; }
  .results-columns { display: flex; flex-direction: column; gap: 20px; }
  .results-column {
    border: 1px solid var(--line); border-radius: 12px; padding: 14px; background: var(--input-bg);
    border-left-width: 3px;
  }
  .results-column-a { border-left-color: var(--brand-600); }
  .results-column-b { border-left-color: var(--success); }
  .results-column h3 {
    display: flex; align-items: center; gap: 6px;
    font-size: 12.5px; text-transform: uppercase; letter-spacing: 0.03em; margin: 0 0 10px;
  }
  .results-column h3 svg { width: 15px; height: 15px; flex: none; }
  .results-column-a h3 { color: var(--brand-700); }
  .results-column-a h3 svg { color: var(--brand-600); }
  .results-column-b h3 { color: var(--success); }
  .results-column-b h3 svg { color: var(--success); }
  .results-table-scroll { max-height: 360px; overflow: auto; border: 1px solid var(--line); border-radius: 10px; background: var(--card-bg); }
  .results-table-scroll table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .results-table-scroll th, .results-table-scroll td {
    border-bottom: 1px solid var(--line-soft); padding: 6px 9px; text-align: left; white-space: nowrap; color: var(--ink-900);
  }
  .results-column-a .results-table-scroll th { background: var(--brand-soft); position: sticky; top: 0; color: var(--brand-700); font-weight: 600; }
  .results-column-b .results-table-scroll th { background: var(--success-bg); position: sticky; top: 0; color: var(--success); font-weight: 600; }
  .results-section .note {
    font-size: 12px; color: var(--warn); background: var(--warn-bg); border: 1px solid var(--warn-line);
    border-radius: 8px; padding: 9px 11px; margin-bottom: 10px; line-height: 1.5;
  }

  .modal-overlay {
    display: none; position: fixed; inset: 0; background: rgba(15, 23, 42, 0.45);
    align-items: center; justify-content: center; padding: 20px; z-index: 80;
  }
  .modal-overlay.visible { display: flex; }
  .modal-card {
    background: var(--card-bg); border-radius: 16px; max-width: 360px; width: 100%;
    padding: 26px; box-shadow: 0 24px 60px -12px rgba(15, 23, 42, 0.35); text-align: center;
  }
  .modal-icon {
    width: 48px; height: 48px; border-radius: 50%; margin: 0 auto 14px;
    background: var(--warn-bg); border: 1.5px solid var(--warn-line);
    display: flex; align-items: center; justify-content: center; color: var(--warn);
  }
  .modal-card h2 { font-size: 16px; margin: 0 0 8px; color: var(--ink-900); }
  .modal-card p { font-size: 13.5px; color: var(--ink-500); margin: 0 0 20px; line-height: 1.5; }
  .modal-card button { margin-top: 0; width: auto; padding: 10px 18px; }
`;

export function renderSuperAdminGenerate(settings: AppSettings, errorMessage?: string): string {
  const errorBlock = errorMessage
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(errorMessage)}</span></div>`
    : '';

  const body = `
    <p class="lede">Déposez les deux documents (Liquidation et DUM) — l'ordre n'a pas d'importance, ils sont identifiés automatiquement.</p>
    <div class="card">
      ${errorBlock}
      <form id="generateForm" method="post" action="/generate" enctype="multipart/form-data">
        <div class="drop-zone" id="zone-liquidation">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M6 2.75h8l4 4V19.5A1.75 1.75 0 0 1 16.25 21h-9A1.75 1.75 0 0 1 5.5 19.25V4.5A1.75 1.75 0 0 1 6 2.75Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
              <path d="M13.5 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
            </svg>
          </div>
          <div class="dz-body">
            <div class="label">Liquidation Douanière (BE)</div>
            <div class="hint">PDF ou image — cliquez ou glissez-déposez</div>
            <div class="filename" id="name-liquidation"></div>
          </div>
          <input type="file" name="liquidation" id="input-liquidation" accept=".pdf,.png,.jpg,.jpeg,.tif,.tiff,.bmp" required />
        </div>

        <div class="drop-zone" id="zone-dum">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M6 2.75h8l4 4V19.5A1.75 1.75 0 0 1 16.25 21h-9A1.75 1.75 0 0 1 5.5 19.25V4.5A1.75 1.75 0 0 1 6 2.75Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
              <path d="M13.5 2.75V7a1 1 0 0 0 1 1h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/>
            </svg>
          </div>
          <div class="dz-body">
            <div class="label">DUM</div>
            <div class="hint">PDF ou image — cliquez ou glissez-déposez</div>
            <div class="filename" id="name-dum"></div>
          </div>
          <input type="file" name="dum" id="input-dum" accept=".pdf,.png,.jpg,.jpeg,.tif,.tiff,.bmp" required />
        </div>

        <button type="submit" id="submitBtn">
          <svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="white" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
          Générer les fichiers Excel
        </button>
        <div id="status"></div>
      </form>

      <div id="successPanel">
        <div class="success-icon">
          <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M5 12.5l4.5 4.5L19 7" stroke="#059669" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </div>
        <div class="success-title">Génération réussie</div>
        <div class="success-subtitle" id="successSubtitle">Le fichier Excel est prêt.</div>

        <div class="success-actions">
          <button type="button" class="secondary" id="downloadAgainBtn">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Télécharger
          </button>
          <button type="button" id="showCostsBtn">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 16.5V9M8 16.5V4M13 16.5v-6M17.5 16.5V7" stroke="white" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Coûts des produits
          </button>
          <button type="button" class="secondary" id="showResultsBtn">
            <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M3 5h14M3 10h14M3 15h9" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
            Afficher résultats
          </button>
        </div>

        <div class="cost-panel" id="costPanel">
          <div class="note">Coût douanier uniquement (droits &amp; taxes ÷ quantité) — le coût total avec achat, fret et assurance arrive dans une prochaine mise à jour.</div>
          <div id="costTableContainer"></div>
        </div>

        <div class="results-section" id="resultsSection">
          <div class="results-toolbar">
            <button type="button" class="secondary" id="exportExcelBtn">
              <svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 3v10.5M10 13.5l-4-4M10 13.5l4-4M4 16.5h12" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
              Exporter Excel
            </button>
          </div>
          <div id="resultsContent"></div>
        </div>

        <button type="button" class="ghost" id="newDeclarationBtn">← Nouvelle déclaration</button>
      </div>
    </div>

    <div class="modal-overlay" id="validationModal">
      <div class="modal-card">
        <div class="modal-icon">
          <svg width="22" height="22" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
        </div>
        <h2>Fichiers manquants</h2>
        <p id="validationModalMessage">Veuillez sélectionner les 2 fichiers (Liquidation et DUM) avant de générer.</p>
        <button type="button" id="validationModalOk">Compris</button>
      </div>
    </div>

    <script>
      function wireDropZone(zoneId, inputId, nameId) {
        const zone = document.getElementById(zoneId);
        const input = document.getElementById(inputId);
        const nameEl = document.getElementById(nameId);
        const checkIcon = '<svg width="12" height="12" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4 10.5l3.5 3.5L16 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

        function setFile(name) {
          nameEl.innerHTML = checkIcon + '<span>' + name + '</span>';
          zone.classList.add('filled');
        }

        zone.addEventListener('click', () => input.click());
        input.addEventListener('change', () => {
          if (input.files[0]) setFile(input.files[0].name);
        });
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
          e.preventDefault();
          zone.classList.remove('dragover');
          if (e.dataTransfer.files.length > 0) {
            input.files = e.dataTransfer.files;
            setFile(input.files[0].name);
          }
        });
      }

      wireDropZone('zone-liquidation', 'input-liquidation', 'name-liquidation');
      wireDropZone('zone-dum', 'input-dum', 'name-dum');

      const form = document.getElementById('generateForm');
      const submitBtn = document.getElementById('submitBtn');
      const statusEl = document.getElementById('status');
      const successPanel = document.getElementById('successPanel');
      const successSubtitle = document.getElementById('successSubtitle');
      const costPanel = document.getElementById('costPanel');
      const costTableContainer = document.getElementById('costTableContainer');
      const showCostsBtn = document.getElementById('showCostsBtn');
      const downloadAgainBtn = document.getElementById('downloadAgainBtn');
      const newDeclarationBtn = document.getElementById('newDeclarationBtn');
      const validationModal = document.getElementById('validationModal');
      const validationModalOk = document.getElementById('validationModalOk');
      const resultsSection = document.getElementById('resultsSection');
      const resultsContent = document.getElementById('resultsContent');
      const showResultsBtn = document.getElementById('showResultsBtn');
      const exportExcelBtn = document.getElementById('exportExcelBtn');
      let cachedCostSummary = null;
      let resultsLoaded = false;

      async function fetchCostSummary() {
        if (cachedCostSummary) return cachedCostSummary;
        const response = await fetch('/last-declaration-cost-summary');
        const result = await response.json();
        if (!response.ok || !result.success) {
          throw new Error(result.error || 'Impossible de charger les coûts.');
        }
        cachedCostSummary = result;
        return result;
      }

      function showValidationModal(message) {
        document.getElementById('validationModalMessage').textContent = message;
        validationModal.classList.add('visible');
      }
      validationModalOk.addEventListener('click', () => validationModal.classList.remove('visible'));
      validationModal.addEventListener('click', (e) => {
        if (e.target === validationModal) validationModal.classList.remove('visible');
      });

      function triggerBlobDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename || 'Declaration.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
      }

      function showFormError(message) {
        const existingError = document.querySelector('#generateForm .error, .card > .error');
        if (existingError) existingError.remove();
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error';
        errorDiv.innerHTML = '<svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span></span>';
        errorDiv.querySelector('span').textContent = 'Échec : ' + message;
        form.parentNode.insertBefore(errorDiv, form);
      }

      form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const liquidationFile = document.getElementById('input-liquidation').files[0];
        const dumFile = document.getElementById('input-dum').files[0];
        if (!liquidationFile || !dumFile) {
          const missing = [];
          if (!liquidationFile) missing.push('Liquidation Douanière');
          if (!dumFile) missing.push('DUM');
          showValidationModal('Fichier(s) manquant(s) : ' + missing.join(', ') + '. Sélectionnez-les avant de générer.');
          return;
        }

        submitBtn.disabled = true;
        statusEl.className = 'busy';
        statusEl.textContent = 'Traitement en cours (OCR + parsing + génération)...';

        const formData = new FormData();
        formData.append('liquidation', liquidationFile);
        formData.append('dum', dumFile);

        try {
          const response = await fetch('/generate', { method: 'POST', body: formData });

          if (!response.ok) {
            const contentType = response.headers.get('content-type') || '';
            if (contentType.includes('application/json')) {
              const result = await response.json();
              if (response.status === 401) {
                window.location.href = '/login';
                return;
              }
              throw new Error(result.error || 'Erreur serveur (' + response.status + ')');
            }
            throw new Error('Erreur serveur (' + response.status + ')');
          }

          await response.blob();

          statusEl.className = '';
          statusEl.textContent = '';
          submitBtn.disabled = false;

          form.style.display = 'none';
          successPanel.classList.add('visible');
          costPanel.classList.remove('visible');
          cachedCostSummary = null;
          resultsSection.classList.remove('visible');
          resultsContent.innerHTML = '';
          resultsLoaded = false;

          fetchCostSummary()
            .then((data) => {
              successSubtitle.textContent = 'Déclaration ' + data.code + ' — ' + data.redevable;
            })
            .catch(() => {});
        } catch (err) {
          statusEl.className = '';
          statusEl.textContent = '';
          submitBtn.disabled = false;
          showFormError(err.message);
        }
      });

      downloadAgainBtn.addEventListener('click', async () => {
        downloadAgainBtn.disabled = true;
        try {
          const response = await fetch('/download');
          if (!response.ok) throw new Error('Impossible de retélécharger le fichier.');
          const blob = await response.blob();
          triggerBlobDownload(blob);
        } catch (err) {
          showFormError(err.message);
        } finally {
          downloadAgainBtn.disabled = false;
        }
      });

      function escapeHtmlClient(value) {
        return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }

      function renderCostTable(data) {
        successSubtitle.textContent = 'Déclaration ' + data.code + ' — ' + data.redevable;

        const rows = data.articles.map((article) => (
          '<tr>' +
          '<td>' + escapeHtmlClient(article.nomArticle) + '</td>' +
          '<td>' + escapeHtmlClient(article.hsCode) + '</td>' +
          '<td>' + escapeHtmlClient(article.pays) + '</td>' +
          '<td class="num">' + article.quantite + '</td>' +
          '<td class="num">' + article.totalTaxes.toFixed(2) + '</td>' +
          '<td class="num cost">' + article.dutyCostPerUnit.toFixed(2) + '</td>' +
          '</tr>'
        )).join('');

        costTableContainer.innerHTML =
          '<table class="cost-table"><thead><tr>' +
          '<th>Produit</th><th>HSC</th><th>Pays</th><th>Qté</th><th>Droits &amp; taxes</th><th>Coût / unité</th>' +
          '</tr></thead><tbody>' + rows + '</tbody></table>';
      }

      const showCostsBtnOriginalHtml = showCostsBtn.innerHTML;
      const backIcon = '<svg width="15" height="15" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12.5 15.5L7 10l5.5-5.5" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';

      showCostsBtn.addEventListener('click', async () => {
        const isVisible = costPanel.classList.contains('visible');
        if (isVisible) {
          costPanel.classList.remove('visible');
          successPanel.classList.remove('cost-exclusive');
          showCostsBtn.innerHTML = showCostsBtnOriginalHtml;
          return;
        }

        showCostsBtn.disabled = true;
        try {
          const result = await fetchCostSummary();
          renderCostTable(result);
          costPanel.classList.add('visible');
          successPanel.classList.add('cost-exclusive');
          showCostsBtn.innerHTML = backIcon + 'Retour';
        } catch (err) {
          showFormError(err.message);
        } finally {
          showCostsBtn.disabled = false;
        }
      });

      showResultsBtn.addEventListener('click', async () => {
        const isVisible = resultsSection.classList.contains('visible');
        if (isVisible) {
          resultsSection.classList.remove('visible');
          return;
        }

        showResultsBtn.disabled = true;
        try {
          if (!resultsLoaded) {
            const response = await fetch('/last-declaration-results');
            if (!response.ok) throw new Error('Impossible de charger les résultats.');
            resultsContent.innerHTML = await response.text();
            resultsLoaded = true;
          }
          resultsSection.classList.add('visible');
        } catch (err) {
          showFormError(err.message);
        } finally {
          showResultsBtn.disabled = false;
        }
      });

      exportExcelBtn.addEventListener('click', () => downloadAgainBtn.click());

      newDeclarationBtn.addEventListener('click', () => {
        window.location.reload();
      });
    </script>
  `;

  return renderShell('generate', 'Générer une déclaration', body, settings, GENERATE_PAGE_STYLE);
}
