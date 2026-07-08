import assert from "node:assert/strict";
import test from "node:test";
import { summarizeReadiness } from "../src/readiness.js";

test("summarizes release readiness with a bounded score", () => {
  const [summary] = summarizeReadiness([
    {
      id: "rel-1",
      title: "Billing cleanup",
      owner: "product",
      status: "ready",
      risk: "medium",
      changes: ["copy", "export", "support docs"]
    }
  ]);

  assert.equal(summary.readinessScore, 50);
  assert.equal(summary.changeCount, 3);
});

