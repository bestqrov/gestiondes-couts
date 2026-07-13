import { randomUUID } from 'node:crypto';
import type { NextFunction, Request, Response } from 'express';
import type { UserRole } from '../db/usersRepository.js';

// Per design spec §2: sessions are still held in memory (not persisted to
// SQLite) — restarting the server logs everyone out, same tradeoff as
// before. What's new is that sessions now carry a role, and the
// credentials they were created from live in the `users` table (bcrypt
// password hashes) instead of a single hardcoded pair.
const SESSION_COOKIE_NAME = 'session';

export interface SessionInfo {
  userId: number;
  username: string;
  role: UserRole;
}

const activeSessions = new Map<string, SessionInfo>();

export function createSession(user: SessionInfo): string {
  const sessionId = randomUUID();
  activeSessions.set(sessionId, user);
  return sessionId;
}

function getSessionIdFromCookie(cookieHeader: string | undefined): string | undefined {
  if (!cookieHeader) return undefined;
  const match = cookieHeader
    .split(';')
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${SESSION_COOKIE_NAME}=`));
  return match?.slice(SESSION_COOKIE_NAME.length + 1);
}

export function setSessionCookie(res: Response, sessionId: string): void {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE_NAME}=${sessionId}; HttpOnly; Path=/; SameSite=Lax`);
}

function getSession(req: Request): SessionInfo | undefined {
  const sessionId = getSessionIdFromCookie(req.headers.cookie);
  if (!sessionId) return undefined;
  return activeSessions.get(sessionId);
}

export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const session = getSession(req);
  if (session) {
    (req as Request & { session?: SessionInfo }).session = session;
    next();
    return;
  }
  // The client's fetch()-based upload flow expects JSON back from POST
  // /generate; redirecting it to the (HTML) /login page here made the
  // client's response.json() throw a confusing "Unexpected token '<'"
  // instead of the real problem. Only redirect for normal page navigation
  // (GET); respond with JSON for API-style POST requests.
  if (req.method === 'POST') {
    res.status(401).json({ success: false, error: 'Session expirée, veuillez vous reconnecter.' });
    return;
  }
  res.redirect('/login');
}

export function requireSuperAdmin(req: Request, res: Response, next: NextFunction): void {
  const existing = (req as Request & { session?: SessionInfo }).session;
  const session = existing ?? getSession(req);
  if (session?.role === 'superadmin') {
    (req as Request & { session?: SessionInfo }).session = session;
    next();
    return;
  }
  res.status(403).send('Accès refusé — réservé au superadmin.');
}
