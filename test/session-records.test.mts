// agent-office#560: session approvals live in one host-wide directory, with the
// runtime recorded, so two agents' sessions on one machine resolve their own
// identity and three plugins agree on who a session is.
import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { test } from "node:test";

// A HOME of its own: the records are host-wide now, so an unrestored real HOME
// would make this file depend on what another test wrote there (#560).
process.env.HOME = mkdtempSync(join(tmpdir(), "ao-sessions-"));
const { legacySessionPointerPath, saveSessionAgent, sessionAgentOf, sessionPointerPath } = await import("../lib/office.ts");

test("the session record is host-wide and names the runtime", () => {
  const path = sessionPointerPath("session-ada");
  assert.ok(path.includes(`${join("agent-office", "sessions")}`), path);
  assert.ok(!path.includes("/opencode/sessions/"), `not buried inside a runtime: ${path}`);
});

test("an approval is written atomically and carries the runtime", () => {
  saveSessionAgent("session-ada", "opencode");
  const record = JSON.parse(readFileSync(sessionPointerPath("session-ada"), "utf8"));
  assert.deepEqual({ agent: record.agent, runtime: record.runtime }, { agent: "opencode", runtime: "opencode" });
  assert.equal(typeof record.at, "number");
});

test("a record at the previous path is read once and copied forward", () => {
  const id = "session-legacy";
  mkdirSync(dirname(legacySessionPointerPath(id)), { recursive: true });
  writeFileSync(legacySessionPointerPath(id), JSON.stringify({ agent: "opencode" }), { mode: 0o600 });
  assert.equal(sessionAgentOf(id), "opencode");
  assert.equal(JSON.parse(readFileSync(sessionPointerPath(id), "utf8")).agent, "opencode");
});

test("two sessions on one host read their own agent, not each other's", () => {
  saveSessionAgent("session-one", "opencode");
  saveSessionAgent("session-two", "other-agent");
  assert.equal(sessionAgentOf("session-one"), "opencode");
  assert.equal(sessionAgentOf("session-two"), "other-agent");
  assert.equal(sessionAgentOf("session-three"), undefined, "a session that never approved has no identity");
});
