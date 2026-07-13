import { describe, it, expect } from 'vitest';
import { checkCredentials, createSession, requireAuth } from '../../src/web/auth.js';
import type { NextFunction, Request, Response } from 'express';

describe('checkCredentials', () => {
  it('accepts the configured username/password', () => {
    expect(checkCredentials('redwan', 'redwan2026')).toBe(true);
  });

  it('rejects a wrong password', () => {
    expect(checkCredentials('redwan', 'wrong')).toBe(false);
  });

  it('rejects a wrong username', () => {
    expect(checkCredentials('someone-else', 'redwan2026')).toBe(false);
  });
});

describe('requireAuth', () => {
  function makeReqRes(cookieHeader: string | undefined, method = 'GET') {
    const req = { headers: { cookie: cookieHeader }, method } as unknown as Request;
    const redirectCalls: string[] = [];
    const jsonCalls: unknown[] = [];
    let statusCode: number | undefined;
    const res = {
      redirect: (url: string) => redirectCalls.push(url),
      status: (code: number) => {
        statusCode = code;
        return res;
      },
      json: (body: unknown) => jsonCalls.push(body),
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
      wasNextCalled: () => nextCalled,
      getStatusCode: () => statusCode,
    };
  }

  it('calls next() for a valid session cookie', () => {
    const sessionId = createSession();
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(`session=${sessionId}`);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(true);
    expect(redirectCalls).toHaveLength(0);
  });

  it('redirects to /login when there is no session cookie', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes(undefined);

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('redirects to /login for an unknown/expired session id', () => {
    const { req, res, next, wasNextCalled, redirectCalls } = makeReqRes('session=not-a-real-session');

    requireAuth(req, res, next);

    expect(wasNextCalled()).toBe(false);
    expect(redirectCalls).toEqual(['/login']);
  });

  it('responds with JSON 401 instead of redirecting for an unauthenticated POST request', () => {
    // POST requests come from the fetch()-based upload flow, which expects
    // JSON back — redirecting to the (HTML) /login page there made the
    // client's response.json() throw a confusing parse error instead of a
    // clear "session expired" message.
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
