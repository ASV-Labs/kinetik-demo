import assert from "node:assert/strict";
import test from "node:test";
import { listReleases } from "../src/releases.js";

test("lists release records without mutating source data", () => {
  const [release] = listReleases();
  release.status = "changed";

  assert.equal(listReleases()[0].status, "draft");
});

