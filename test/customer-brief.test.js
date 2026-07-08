import assert from "node:assert/strict";
import test from "node:test";
import { buildCustomerBrief } from "../src/customer-brief.js";

test("builds a customer-facing brief with change citations", () => {
  const brief = buildCustomerBrief({
    title: "July Readiness Report",
    owner: "operations",
    risk: "medium",
    changes: ["customer export cleanup", "status summary endpoint"]
  });

  assert.equal(brief.audience, "customer_success");
  assert.equal(brief.citations.length, 2);
  assert.match(brief.summary, /2 changes/);
});

