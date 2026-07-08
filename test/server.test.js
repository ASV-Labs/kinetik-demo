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

