# Récupération de mot de passe par email — Design

## Context

Today, only the superadmin can reset a user's password (Utilisateurs >
Modifier), and a user can only change their own password while already
logged in (Réglages > Identifiants). There's no self-service recovery if
someone is locked out. This adds a standard "forgot password" email flow.

## Scope

- Add a mandatory `email` field to user accounts (going forward — existing
  accounts without one keep working, just can't use the reset flow until
  an email is added).
- "Mot de passe oublié ?" link on the login page → email entry → reset
  email sent via Resend → reset link → new password form.
- Out of scope: email verification/confirmation flow, changing a user's own
  email requiring re-verification — an email typo just means recovery mail
  goes nowhere; the superadmin can still fix it via Utilisateurs > Modifier.

## Data layer (`src/db/usersRepository.ts`)

Extend `UserDocument`/`User` with:
```ts
email: string | null; // null only for accounts created before this feature
resetToken: string | null; // sha256 hash of the token, never the raw token
resetTokenExpiresAt: string | null; // ISO timestamp
```

New functions:
- `createUser(..., email: string)` — email becomes a required parameter.
- `updateEmail(collection, userId, email)`.
- `setResetToken(collection, userId, tokenHash, expiresAt)`.
- `findUserByValidResetToken(collection, tokenHash)` — returns the user only
  if `resetTokenExpiresAt` is still in the future.
- `clearResetToken(collection, userId)` — called after a successful reset,
  and the token is also single-use (cleared immediately on use, not just
  on expiry).
- `findUserByEmail(collection, email)`.

A unique index on `email` (like the existing one on `username`) is NOT
added, since existing accounts have `email: null` and Mongo's unique index
would reject multiple nulls under some configurations — instead, the
"forgot password" lookup just takes the first match (acceptable at this
app's scale; duplicate emails aren't otherwise prevented).

## Email sending (`src/email/resendClient.ts`, new)

A small wrapper around Resend's HTTP API (no SDK dependency needed — a
single `fetch` call), reading `RESEND_API_KEY` and `RESEND_FROM_EMAIL` from
the environment (same pattern as `MONGODB_URI`). If either is unset, the
forgot-password route logs a clear error and still shows the generic
"a reset link was sent" message (never reveals config problems to the
end user) — mirrors how MongoDB-unreachable states are already handled
elsewhere in this app.

```ts
export async function sendPasswordResetEmail(
  to: string,
  companyName: string,
  resetUrl: string
): Promise<void>
```

## Routes (`src/web/server.ts`)

- `GET /forgot-password` — the email-entry form.
- `POST /forgot-password` — looks up the user by email; if found, generates
  a random 32-byte token (`crypto.randomBytes`), stores its SHA-256 hash +
  a 1-hour expiry, emails the raw token as part of a link
  (`/reset-password?token=...`). Always redirects to a "si un compte
  existe..." confirmation page, whether or not the email matched.
- `GET /reset-password?token=...` — the new-password form (token passed
  through as a hidden field). If the token doesn't resolve to a valid,
  unexpired reset request, shows an error instead of the form.
- `POST /reset-password` — validates the token again server-side, sets the
  new password, clears the token, logs the user in (same as a normal
  login) and redirects to their landing page.

## UI

- Login page: small "Mot de passe oublié ?" link under the password field.
- New `forgot-password.html` / `reset-password.html` view templates,
  reusing the login page's split-panel branding treatment (same
  `{{BRAND_OVERRIDE}}`/`{{FAVICON_LINK}}`/logo placeholders).
- "Nouveau compte" (Utilisateurs) and the "Modifier" inline edit form gain
  a required email input.
- Réglages > Identifiants gains an email input (required, prefilled with
  the current value if set).

## Testing

- Repository unit tests (fake Mongo collection) for the new
  `usersRepository` functions: email storage, token set/validate/expire/
  single-use, findUserByEmail.
- A test for `sendPasswordResetEmail`'s request-building logic against a
  mocked `fetch` (verifies the API call shape, not a real send).
- Manual smoke test against the real disposable MongoDB test database,
  and — only once `RESEND_API_KEY` is actually configured — one real send
  to a test inbox to confirm the email arrives and the link works
  end-to-end.
