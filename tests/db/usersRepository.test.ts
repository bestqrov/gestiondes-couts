import { describe, it, expect } from 'vitest';
import { createDatabase } from '../../src/db/database.js';
import {
  createUser,
  findUserByUsername,
  verifyPassword,
  listUsers,
  setUserDisabled,
  seedSuperAdminIfEmpty,
  updateUsername,
  updatePassword,
} from '../../src/db/usersRepository.js';

describe('usersRepository', () => {
  it('creates a user with a hashed password and finds it back by username', () => {
    const db = createDatabase(':memory:');

    const created = createUser(db, 'alice', 'hunter2', 'admin');
    expect(created.username).toBe('alice');
    expect(created.role).toBe('admin');
    expect(created.disabledAt).toBeNull();

    const found = findUserByUsername(db, 'alice');
    expect(found).toBeDefined();
    expect(found!.passwordHash).not.toBe('hunter2'); // must be hashed, not plaintext
    expect(verifyPassword(found!.passwordHash, 'hunter2')).toBe(true);
    expect(verifyPassword(found!.passwordHash, 'wrong-password')).toBe(false);

    db.close();
  });

  it('returns undefined for an unknown username', () => {
    const db = createDatabase(':memory:');
    expect(findUserByUsername(db, 'nobody')).toBeUndefined();
    db.close();
  });

  it('lists users ordered by creation time, oldest first', () => {
    const db = createDatabase(':memory:');
    createUser(db, 'first', 'pw', 'admin');
    createUser(db, 'second', 'pw', 'admin');

    const users = listUsers(db);
    expect(users.map((u) => u.username)).toEqual(['first', 'second']);

    db.close();
  });

  it('disables and re-enables a user', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'bob', 'pw', 'admin');

    setUserDisabled(db, user.id, true);
    expect(listUsers(db).find((u) => u.id === user.id)!.disabledAt).not.toBeNull();

    setUserDisabled(db, user.id, false);
    expect(listUsers(db).find((u) => u.id === user.id)!.disabledAt).toBeNull();

    db.close();
  });

  it('updates a user’s username', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'redwan', 'pw', 'superadmin');

    updateUsername(db, user.id, 'redouane');

    expect(findUserByUsername(db, 'redwan')).toBeUndefined();
    expect(findUserByUsername(db, 'redouane')).toBeDefined();

    db.close();
  });

  it('updates a user’s password, re-hashing it', () => {
    const db = createDatabase(':memory:');
    const user = createUser(db, 'redouane', 'oldpass', 'superadmin');

    updatePassword(db, user.id, 'newpass*2026');

    const found = findUserByUsername(db, 'redouane')!;
    expect(verifyPassword(found.passwordHash, 'oldpass')).toBe(false);
    expect(verifyPassword(found.passwordHash, 'newpass*2026')).toBe(true);

    db.close();
  });

  it('seeds a superadmin only when the users table is empty', () => {
    const db = createDatabase(':memory:');

    seedSuperAdminIfEmpty(db, 'root', 'rootpass');
    expect(listUsers(db)).toHaveLength(1);
    expect(listUsers(db)[0].role).toBe('superadmin');

    seedSuperAdminIfEmpty(db, 'root2', 'otherpass');
    expect(listUsers(db)).toHaveLength(1); // unchanged — table wasn't empty
    expect(listUsers(db)[0].username).toBe('root');

    db.close();
  });
});
