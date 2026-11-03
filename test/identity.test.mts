import assert from "node:assert/strict"
import test from "node:test"

import { identityHeaders, normaliseChildName } from "../index.ts"

// The rules #512 answers 400 to, tested as rules: the plugin has to normalise a
// launcher's name before it reaches the office, and a name outside [a-z0-9._-]
// with a leading punctuation mark or over 64 characters is a refusal, not a nit.
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

test("with no name the child is s- plus eight hex of its session id", () => {
  assert.equal(normaliseChildName(undefined, "01a18ad8-a92f-756b-8833-c9845c4e3d5b"), "s-01a18ad8")
  // Not a session id at all, so there is no honest name to send.
  assert.equal(normaliseChildName(undefined, undefined), undefined)
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
  // No parent to name: no via at all, rather than a header with an empty parent.
  assert.deepEqual(identityHeaders({ child: true, childName: "Office Inbox", runtime: "opencode" }), {})
  assert.deepEqual(identityHeaders({ approved: "opencode", child: true, runtime: "opencode" }), {
    "X-Office-Session-Agent": "opencode",
  })
})

test("a cut that lands on a separator does not leave one trailing", () => {
  // 63 characters and then a space: the strip would otherwise cut at position 64
  // and leave the separator that the space became.
  const name = "a".repeat(63) + " b"
  assert.equal(normaliseChildName(name, undefined), "a".repeat(63))
  assert.equal(normaliseChildName("a".repeat(63) + "-b", undefined), "a".repeat(63))
})
