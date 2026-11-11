import assert from "node:assert/strict"
import test from "node:test"

import { identityHeaders, normaliseChildName } from "../index.ts"
// The suite's own env: the browser opener stays off in this process too.
import "./suite.mts"

// The rules #512 answers 400 to, tested as rules: a launcher's name is normalised
// before it reaches the office, because a name outside [a-z0-9._-], with leading
// punctuation, or over 64 characters is a refusal rather than a nit. These live
// apart from test/identity.test.mts, which covers worktree identity end to end
// and needs a home of its own; these need nothing.
test("a launcher's child name is normalised to what the office accepts", () => {
  assert.equal(normaliseChildName("  Office Inbox!  ", undefined), "office-inbox")
  assert.equal(normaliseChildName("--Lead__Agent.2", undefined), "lead__agent.2")
  assert.equal(normaliseChildName("MiXeD Case  Name", undefined), "mixed-case-name")
})

test("a long child name is cut to 64 characters", () => {
  const long = "a".repeat(70)
  assert.equal(normaliseChildName(long, undefined)?.length, 64)
  assert.equal(normaliseChildName(long, undefined), "a".repeat(64))
})

test("a cut that lands on a separator does not leave one trailing", () => {
  // 63 characters and then a space: the cut would otherwise end on the separator
  // that the space became.
  assert.equal(normaliseChildName("a".repeat(63) + " b", undefined), "a".repeat(63))
  assert.equal(normaliseChildName("a".repeat(63) + "-b", undefined), "a".repeat(63))
})

test("with no name the child is s- plus eight hex of its session id", () => {
  assert.equal(normaliseChildName(undefined, "01a18ad8-a92f-756b-8833-c9845c4e3d5b"), "s-01a18ad8")
  // Too short to give eight hex characters, so there is no honest name to send.
  assert.equal(normaliseChildName(undefined, "abc"), undefined)
  assert.equal(normaliseChildName("   ", "abc"), undefined)
})

test("a top-level session names its agent and sends no via", () => {
  assert.deepEqual(identityHeaders({ approved: "opencode", sessionID: "s1", runtime: "opencode" }), {
    "X-Office-Session-Agent": "opencode",
  })
})

test("a child names its parent and runtime, and sends nothing without a parent", () => {
  assert.deepEqual(identityHeaders({ approved: "opencode", child: true, childName: "Office Inbox", runtime: "opencode" }), {
    "X-Office-Session-Agent": "opencode",
    "X-Office-Via": "child=office-inbox; parent=opencode; runtime=opencode",
  })
  assert.deepEqual(identityHeaders({ child: true, childName: "Office Inbox", runtime: "opencode" }), {})
  assert.deepEqual(identityHeaders({ approved: "opencode", child: true, runtime: "opencode" }), {
    "X-Office-Session-Agent": "opencode",
  })
})
