import bcrypt from 'bcryptjs';
import type Database from 'better-sqlite3';

export type UserRole = 'admin' | 'superadmin';

export interface User {
  id: number;
  username: string;
  role: UserRole;
  createdAt: string;
  disabledAt: string | null;
}

export interface UserWithPasswordHash extends User {
  passwordHash: string;
}

interface UserRow {
  id: number;
  username: string;
  password_hash: string;
  role: UserRole;
  created_at: string;
  disabled_at: string | null;
}

function rowToUser(row: UserRow): User {
  return {
    id: row.id,
    username: row.username,
    role: row.role,
    createdAt: row.created_at,
    disabledAt: row.disabled_at,
  };
}

const SALT_ROUNDS = 10;

export function createUser(
  db: Database.Database,
  username: string,
  password: string,
  role: UserRole
): User {
  const passwordHash = bcrypt.hashSync(password, SALT_ROUNDS);
  const createdAt = new Date().toISOString();
  const result = db
    .prepare('INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)')
    .run(username, passwordHash, role, createdAt);
  return { id: Number(result.lastInsertRowid), username, role, createdAt, disabledAt: null };
}

export function findUserByUsername(
  db: Database.Database,
  username: string
): UserWithPasswordHash | undefined {
  const row = db.prepare('SELECT * FROM users WHERE username = ?').get(username) as
    | UserRow
    | undefined;
  if (!row) return undefined;
  return { ...rowToUser(row), passwordHash: row.password_hash };
}

export function verifyPassword(passwordHash: string, password: string): boolean {
  return bcrypt.compareSync(password, passwordHash);
}

export function listUsers(db: Database.Database): User[] {
  const rows = db.prepare('SELECT * FROM users ORDER BY created_at ASC').all() as UserRow[];
  return rows.map(rowToUser);
}

export function countUsers(db: Database.Database): number {
  const row = db.prepare('SELECT COUNT(*) as count FROM users').get() as { count: number };
  return row.count;
}

export function updateUsername(db: Database.Database, userId: number, username: string): void {
  db.prepare('UPDATE users SET username = ? WHERE id = ?').run(username, userId);
}

export function updatePassword(db: Database.Database, userId: number, password: string): void {
  const passwordHash = bcrypt.hashSync(password, SALT_ROUNDS);
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(passwordHash, userId);
}

export function setUserDisabled(db: Database.Database, userId: number, disabled: boolean): void {
  db.prepare('UPDATE users SET disabled_at = ? WHERE id = ?').run(
    disabled ? new Date().toISOString() : null,
    userId
  );
}

export function seedSuperAdminIfEmpty(
  db: Database.Database,
  username: string,
  password: string
): void {
  if (countUsers(db) > 0) return;
  createUser(db, username, password, 'superadmin');
}
