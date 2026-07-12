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
  function makeReqRes(cookieHeader: string | undefined) {
    const req = { headers: { cookie: cookieHeader } } as unknown as Request;
    const redirectCalls: string[] = [];
    const res = { redirect: (url: string) => redirectCalls.push(url) } as unknown as Response;
    let nextCalled = false;
    const next: NextFunction = () => {
      nextCalled = true;
    };
    return { req, res, next, redirectCalls, wasNextCalled: () => nextCalled };
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
});
