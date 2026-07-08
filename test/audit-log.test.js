import assert from "node:assert/strict";
import test from "node:test";
import { recordReleaseAudit } from "../src/audit-log.js";

test("records release audit events with actor and target context", () => {
  const event = recordReleaseAudit(
    {
      id: "rel-1",
      owner: "support",
      status: "ready",
      risk: "low"
    },
    "release.mark_ready",
    "user-123"
  );

  assert.equal(event.actorId, "user-123");
  assert.equal(event.targetType, "release");
  assert.equal(event.targetId, "rel-1");
  assert.equal(event.metadata.risk, "low");
});

