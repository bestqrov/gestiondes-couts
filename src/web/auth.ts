import { randomUUID } from 'node:crypto';
import type { NextFunction, Request, Response } from 'express';

// Single hardcoded credential pair for an internal, single-user local tool —
// this is an access gate, not a security boundary (no password hashing, no
// rate limiting, no HTTPS enforcement). Per design spec §3.6.1: session
// persists in memory for the server's lifetime; restarting the server logs
// everyone out. Do not reuse this pattern if this app is ever exposed beyond
// localhost.
const VALID_USERNAME = process.env.APP_USERNAME ?? 'redwan';
const VALID_PASSWORD = process.env.APP_PASSWORD ?? 'redwan2026';

const SESSION_COOKIE_NAME = 'session';
const activeSessions = new Set<string>();

export function checkCredentials(username: string, password: string): boolean {
  return username === VALID_USERNAME && password === VALID_PASSWORD;
}

export function createSession(): string {
  const sessionId = randomUUID();
  activeSessions.add(sessionId);
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

export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const sessionId = getSessionIdFromCookie(req.headers.cookie);
  if (sessionId && activeSessions.has(sessionId)) {
    next();
    return;
  }
  res.redirect('/login');
}
