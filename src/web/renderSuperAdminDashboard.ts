import type { User } from '../db/usersRepository.js';

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

export function renderSuperAdminDashboard(
  users: User[],
  currentUserId: number,
  errorMessage?: string
): string {
  const rows = users.map((user) => renderUserRow(user, currentUserId)).join('');
  const errorBlock = errorMessage
    ? `<div class="error"><svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10 6.5v4M10 13.2h.01M10 2.5l7.5 13H2.5l7.5-13Z" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg><span>${escapeHtml(errorMessage)}</span></div>`
    : '';

  return `<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Superadmin — Gestion des comptes</title>
<style>
  :root {
    --ink-900: #0f172a; --ink-700: #334155; --ink-500: #64748b; --ink-400: #94a3b8;
    --line: #e2e8f0; --line-soft: #eef2f7;
    --brand: #4338ca; --brand-600: #4f46e5; --brand-700: #3730a3; --brand-soft: #eef2ff;
    --danger: #b91c1c; --danger-bg: #fef2f2; --danger-line: #fecaca;
    --success: #15803d; --success-bg: #f0fdf4; --success-line: #bbf7d0;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; min-height: 100vh; padding: 32px 24px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: var(--ink-900);
    background: radial-gradient(1100px 600px at 12% -10%, #eef2ff 0%, transparent 55%),
      radial-gradient(900px 500px at 100% 110%, #f0fdfa 0%, transparent 55%), #f8fafc;
  }
  .wrap { max-width: 820px; margin: 0 auto; }
  .top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 4px; }
  h1 { font-size: 20px; font-weight: 700; letter-spacing: -0.01em; margin: 0; }
  a.back { font-size: 13px; color: var(--ink-500); text-decoration: none; }
  a.back:hover { color: var(--brand-600); }
  .lede { font-size: 13.5px; color: var(--ink-500); margin: 4px 0 24px; }
  .card {
    background: #ffffff; border: 1px solid var(--line); border-radius: 16px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 20px 50px -12px rgba(15, 23, 42, 0.1);
    padding: 24px; margin-bottom: 20px;
  }
  .card h2 { font-size: 15px; margin: 0 0 16px; }
  table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  th { text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em;
    color: var(--ink-400); font-weight: 600; padding: 0 10px 10px; border-bottom: 1px solid var(--line); }
  td { padding: 12px 10px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
  tr:last-child td { border-bottom: none; }
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
    background: #f8fafc; border: 1.5px solid var(--line); border-radius: 9px;
  }
  input:focus, select:focus { outline: none; background: #fff; border-color: var(--brand-600); box-shadow: 0 0 0 3.5px rgba(79, 70, 229, 0.14); }
  .create-form { display: grid; grid-template-columns: 1.2fr 1.2fr 1fr auto; gap: 12px; align-items: end; }
  .create-form button {
    padding: 10px 16px; font-size: 13.5px; color: #fff;
    background: linear-gradient(135deg, var(--brand-600), var(--brand-700)); border: none;
    box-shadow: 0 6px 16px -4px rgba(67, 56, 202, 0.4);
  }
  .error {
    display: flex; gap: 8px; align-items: flex-start; background: var(--danger-bg); color: var(--danger);
    border: 1px solid var(--danger-line); border-radius: 10px; padding: 11px 13px; margin-bottom: 18px;
    font-size: 13px; line-height: 1.45;
  }
  .error svg { flex: none; margin-top: 1px; }
  @media (max-width: 640px) { .create-form { grid-template-columns: 1fr; } }
</style>
</head>
<body>
  <div class="wrap">
    <div class="top">
      <h1>Gestion des comptes</h1>
      <a class="back" href="/">&larr; Retour à l'outil</a>
    </div>
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
  </div>
</body>
</html>`;
}
