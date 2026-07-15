import { describe, it, expect } from 'vitest';
import { createFakeCollection } from '../helpers/fakeMongoCollection.js';
import {
  createUser,
  findUserByUsername,
  verifyPassword,
  listUsers,
  setUserDisabled,
  seedSuperAdminIfEmpty,
  updateUsername,
  updatePassword,
  ensureUsersIndexes,
  type UserDocument,
} from '../../src/db/usersRepository.js';

function makeCollection() {
  return createFakeCollection<UserDocument>();
}

describe('usersRepository', () => {
  it('creates a user with a hashed password and finds it back by username', async () => {
    const collection = makeCollection();

    const created = await createUser(collection, 'alice', 'hunter2', 'admin');
    expect(created.username).toBe('alice');
    expect(created.role).toBe('admin');
    expect(created.disabledAt).toBeNull();
    expect(typeof created.id).toBe('string');

    const found = await findUserByUsername(collection, 'alice');
    expect(found).toBeDefined();
    expect(found!.passwordHash).not.toBe('hunter2'); // must be hashed, not plaintext
    expect(verifyPassword(found!.passwordHash, 'hunter2')).toBe(true);
    expect(verifyPassword(found!.passwordHash, 'wrong-password')).toBe(false);
  });

  it('returns undefined for an unknown username', async () => {
    const collection = makeCollection();
    expect(await findUserByUsername(collection, 'nobody')).toBeUndefined();
  });

  it('lists users ordered by creation time, oldest first', async () => {
    const collection = makeCollection();
    await createUser(collection, 'first', 'pw', 'admin');
    await createUser(collection, 'second', 'pw', 'admin');

    const users = await listUsers(collection);
    expect(users.map((u) => u.username)).toEqual(['first', 'second']);
  });

  it('disables and re-enables a user', async () => {
    const collection = makeCollection();
    const user = await createUser(collection, 'bob', 'pw', 'admin');

    await setUserDisabled(collection, user.id, true);
    expect((await listUsers(collection)).find((u) => u.id === user.id)!.disabledAt).not.toBeNull();

    await setUserDisabled(collection, user.id, false);
    expect((await listUsers(collection)).find((u) => u.id === user.id)!.disabledAt).toBeNull();
  });

  it('updates a user’s username', async () => {
    const collection = makeCollection();
    const user = await createUser(collection, 'redwan', 'pw', 'superadmin');

    await updateUsername(collection, user.id, 'redouane');

    expect(await findUserByUsername(collection, 'redwan')).toBeUndefined();
    expect(await findUserByUsername(collection, 'redouane')).toBeDefined();
  });

  it('updates a user’s password, re-hashing it', async () => {
    const collection = makeCollection();
    const user = await createUser(collection, 'redouane', 'oldpass', 'superadmin');

    await updatePassword(collection, user.id, 'newpass*2026');

    const found = (await findUserByUsername(collection, 'redouane'))!;
    expect(verifyPassword(found.passwordHash, 'oldpass')).toBe(false);
    expect(verifyPassword(found.passwordHash, 'newpass*2026')).toBe(true);
  });

  it('seeds a superadmin only when the users collection is empty', async () => {
    const collection = makeCollection();

    await seedSuperAdminIfEmpty(collection, 'root', 'rootpass');
    expect(await listUsers(collection)).toHaveLength(1);
    expect((await listUsers(collection))[0].role).toBe('superadmin');

    await seedSuperAdminIfEmpty(collection, 'root2', 'otherpass');
    expect(await listUsers(collection)).toHaveLength(1); // unchanged — collection wasn't empty
    expect((await listUsers(collection))[0].username).toBe('root');
  });

  it('rejects a duplicate username once the unique index is set up', async () => {
    const collection = makeCollection();
    await ensureUsersIndexes(collection);
    await createUser(collection, 'taken', 'pw', 'admin');

    await expect(createUser(collection, 'taken', 'pw2', 'admin')).rejects.toMatchObject({
      code: 11000,
    });
  });
});
