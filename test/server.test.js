import assert from "node:assert/strict";
import test from "node:test";
import { createServer } from "../src/server.js";

test("serves readiness summaries", async () => {
  const server = createServer();
  await new Promise((resolve) => server.listen(0, resolve));

  try {
    const { port } = server.address();
    const response = await fetch(`http://127.0.0.1:${port}/api/readiness`);
    const body = await response.json();

    assert.equal(response.status, 200);
    assert.equal(body.readiness[0].id, "rel-2026-07");
    assert.equal(typeof body.readiness[0].readinessScore, "number");
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});

test("requires auth before approving releases", async () => {
  const server = createServer();
  await new Promise((resolve) => server.listen(0, resolve));

  try {
    const { port } = server.address();
    const rejected = await fetch(`http://127.0.0.1:${port}/api/releases/approve`, {
      method: "POST"
    });
    const accepted = await fetch(`http://127.0.0.1:${port}/api/releases/approve`, {
      method: "POST",
      headers: { authorization: "Bearer demo-reviewer-token" }
    });
    const body = await accepted.json();

    assert.equal(rejected.status, 401);
    assert.equal(accepted.status, 202);
    assert.equal(body.auditEvent.action, "release.approve");
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
});
