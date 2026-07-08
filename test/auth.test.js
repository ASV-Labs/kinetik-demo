import assert from "node:assert/strict";
import test from "node:test";
import { readSession, requireReleaseApprover } from "../src/auth.js";

test("reads bearer sessions from authorization headers", () => {
  const session = readSession({
    headers: { authorization: "Bearer demo-owner-token" }
  });

  assert.equal(session.userId, "user-owner");
  assert.equal(session.role, "owner");
});

test("rejects missing release approval sessions", () => {
  const result = requireReleaseApprover({ headers: {} });

  assert.equal(result.allowed, false);
  assert.equal(result.status, 401);
});

