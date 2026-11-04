// agent-office#557 closed as not observable from the suite; ensureIdentity is kept
// as hardening, and this is the part of it that can be tested on its own: whether
// dispatch should check in before writing a pointer.
import assert from "node:assert/strict";
import { test } from "node:test";

import { needsCheckin } from "../lib/office.ts";

test("no agent and no override means check in", () => {
  assert.equal(needsCheckin(false, undefined), true, "a dispatch with no identity must fetch one");
  assert.equal(needsCheckin(false, ""), true, "an empty slug is no identity either");
});

test("an existing agent is not re-fetched", () => {
  assert.equal(needsCheckin(false, "opencode"), false);
});

test("a credentials override is an identity of its own", () => {
  assert.equal(needsCheckin(true, undefined), false, "the override token is the identity; there is no pointer to write");
});
