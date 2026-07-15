import { describe, it, expect } from 'vitest';
import {
  createSession,
  requireAuth,
  requireSuperAdmin,
  destroySession,
  type SessionInfo,
} from '../../src/web/auth.js';
import type { NextFunction, Request, Response } from 'express';

function makeReqRes(cookieHeader: string | undefined, method = 'GET') {
  const req = { headers: { cookie: cookieHeader }, method } as unknown as Request;
  const redirectCalls: string[] = [];
  const jsonCalls: unknown[] = [];
  const sendCalls: unknown[] = [];
  const headers: Record<string, string> = {};
  let statusCode: number | undefined;
  const res = {
    redirect: (url: string) => redirectCalls.push(url),
    status: (code: number) => {
      statusCode = code;
      return res;
    },
    json: (body: unknown) => jsonCalls.push(body),
    send: (body: unknown) => sendCalls.push(body),
    setHeader: (name: string, value: string) => {
      headers[name] = value;
    },
  } as unknown as Response;
  let nextCalled = false;
  const next: NextFunction = () => {
    nextCalled = true;
  };
  return {
    req,
    res,
    next,
    redirectCalls,
    jsonCalls,
    sendCalls,
    headers,
    wasNextCalled: () => nextCalled,
    getStatusCode: () => statusCode,
  };
}

describe('requireAuth', () => {
  it('calls next() and attaches the session for a valid session cookie', () => {
    const sessionId = createSession({ userId: '1', username: 'alice', role: 'admin' });
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(`session=${sessionId}`);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(true);
    expect(redirectCalls).toHaveLength(0);
    expect((req as Request & { session?: SessionInfo }).session).toEqual({
      userId: '1',
      username: 'alice',
      role: 'admin',
    });
  });

  it('redirects to /login for a GET request with no session cookie', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(undefined);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('redirects to /login for a GET request with an unknown/expired session id', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes('session=not-a-real-session');

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('responds with JSON 401 instead of redirecting for an unauthenticated POST request', () => {
    const { req, res, next, wasNextCalled, redirectCalls, jsonCalls, getStatusCode } = makeReqRes(
      undefined,
      'POST'
    );

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toHaveLength(0);
    expect(getStatusCode()).toBe(401);
    expect(jsonCalls).toEqual([{ success: false, error: 'Session expirée, veuillez vous reconnecter.' }]);
  });
});

describe('requireSuperAdmin', () => {
  it('calls next() when the session role is superadmin', () => {
    const sessionId = createSession({ userId: '2', username: 'root', role: 'superadmin' });
    const { req, res, next, wasNextCalled, sendCalls } = makeReqRes(`session=${sessionId}`);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(true);
    expect(sendCalls).toHaveLength(0);
  });

  it('responds 403 when the session role is admin (not superadmin)', () => {
    const sessionId = createSession({ userId: '3', username: 'alice', role: 'admin' });
    const { req, res, next, wasNextCalled, getStatusCode, sendCalls } = makeReqRes(`session=${sessionId}`);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(getStatusCode()).toBe(403);
    expect(sendCalls).toHaveLength(1);
  });

  it('responds 403 when there is no session at all', () => {
    const { req, res, next, wasNextCalled, getStatusCode } = makeReqRes(undefined);

    requireSuperAdmin(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(getStatusCode()).toBe(403);
  });
});

describe('destroySession', () => {
  it('invalidates the session so a subsequent requireAuth call fails, and clears the cookie', () => {
    const sessionId = createSession({ userId: '4', username: 'bob', role: 'admin' });
    const { req, res, headers } = makeReqRes(`session=${sessionId}`);

    destroySession(req, res);

    expect(headers['Set-Cookie']).toContain('Max-Age=0');

    const afterLogout = makeReqRes(`session=${sessionId}`);
    requireAuth(afterLogout.req, afterLogout.res, afterLogout.next);
    expect(afterLogout.wasNextCalled()).toBe(false);
    expect(afterLogout.redirectCalls).toEqual(['/login']);
  });

  it('is a no-op (does not throw) when there is no session cookie', () => {
    const { req, res, headers } = makeReqRes(undefined);

    expect(() => destroySession(req, res)).not.toThrow();
    expect(headers['Set-Cookie']).toContain('Max-Age=0');
  });
});
