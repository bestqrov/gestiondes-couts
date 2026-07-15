import bcrypt from 'bcryptjs';
import { ObjectId, type Collection } from 'mongodb';

export type UserRole = 'admin' | 'superadmin';

export const USERS_COLLECTION = 'users';

export interface User {
  id: string;
  username: string;
  role: UserRole;
  createdAt: string;
  disabledAt: string | null;
}

export interface UserWithPasswordHash extends User {
  passwordHash: string;
}

// The MongoDB document shape — _id is optional here because insertOne is
// called with a plain object before Mongo assigns one.
export interface UserDocument {
  _id?: ObjectId;
  username: string;
  passwordHash: string;
  role: UserRole;
  createdAt: string;
  disabledAt: string | null;
}

function docToUser(doc: UserDocument & { _id: ObjectId }): User {
  return {
    id: doc._id.toHexString(),
    username: doc.username,
    role: doc.role,
    createdAt: doc.createdAt,
    disabledAt: doc.disabledAt,
  };
}

const SALT_ROUNDS = 10;

// Called once at boot (see server.ts) — a unique index on username makes
// MongoDB itself reject duplicates (surfaced as an E11000 error, code
// 11000), the same role SQLite's UNIQUE constraint used to play.
export async function ensureUsersIndexes(collection: Collection<UserDocument>): Promise<void> {
  await collection.createIndex({ username: 1 }, { unique: true });
}

export async function createUser(
  collection: Collection<UserDocument>,
  username: string,
  password: string,
  role: UserRole
): Promise<User> {
  const passwordHash = bcrypt.hashSync(password, SALT_ROUNDS);
  const createdAt = new Date().toISOString();
  const doc: UserDocument = { username, passwordHash, role, createdAt, disabledAt: null };
  const result = await collection.insertOne(doc);
  return { id: result.insertedId.toHexString(), username, role, createdAt, disabledAt: null };
}

export async function findUserByUsername(
  collection: Collection<UserDocument>,
  username: string
): Promise<UserWithPasswordHash | undefined> {
  const doc = await collection.findOne({ username });
  if (!doc) return undefined;
  return { ...docToUser(doc as UserDocument & { _id: ObjectId }), passwordHash: doc.passwordHash };
}

export function verifyPassword(passwordHash: string, password: string): boolean {
  return bcrypt.compareSync(password, passwordHash);
}

export async function listUsers(collection: Collection<UserDocument>): Promise<User[]> {
  const docs = await collection.find().sort({ createdAt: 1 }).toArray();
  return docs.map((doc) => docToUser(doc as UserDocument & { _id: ObjectId }));
}

export async function countUsers(collection: Collection<UserDocument>): Promise<number> {
  return collection.countDocuments();
}

export async function setUserDisabled(
  collection: Collection<UserDocument>,
  userId: string,
  disabled: boolean
): Promise<void> {
  await collection.updateOne(
    { _id: new ObjectId(userId) },
    { $set: { disabledAt: disabled ? new Date().toISOString() : null } }
  );
}

export async function updateUsername(
  collection: Collection<UserDocument>,
  userId: string,
  username: string
): Promise<void> {
  await collection.updateOne({ _id: new ObjectId(userId) }, { $set: { username } });
}

export async function updatePassword(
  collection: Collection<UserDocument>,
  userId: string,
  password: string
): Promise<void> {
  const passwordHash = bcrypt.hashSync(password, SALT_ROUNDS);
  await collection.updateOne({ _id: new ObjectId(userId) }, { $set: { passwordHash } });
}

export async function deleteUser(collection: Collection<UserDocument>, userId: string): Promise<void> {
  await collection.deleteOne({ _id: new ObjectId(userId) });
}

export async function seedSuperAdminIfEmpty(
  collection: Collection<UserDocument>,
  username: string,
  password: string
): Promise<void> {
  if ((await countUsers(collection)) > 0) return;
  await createUser(collection, username, password, 'superadmin');
}
