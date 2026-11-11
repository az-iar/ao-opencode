// agent-office#560: session approvals live in one host-wide directory, with the
// runtime recorded, so two agents' sessions on one machine resolve their own
// identity and three plugins agree on who a session is.
import assert from "node:assert/strict";
import { chmodSync, existsSync, mkdirSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { after, test } from "node:test";
import { removeTempRoot, tempDir } from "./suite.mts";

// A HOME of its own: the records are host-wide now, so an unrestored real HOME
// would make this file depend on what another test wrote there (#560).
process.env.HOME = tempDir("ao-sessions-");
after(removeTempRoot);
const { legacySessionPointerPath, saveSessionAgent, sessionAgentOf, sessionPointerPath } = await import("../lib/office.ts");

test("the session record is host-wide and names the runtime", () => {
  const path = sessionPointerPath("session-ada");
  assert.ok(path.includes(`${join("agent-office", "sessions")}`), path);
  assert.ok(!path.includes("/opencode/sessions/"), `not buried inside a runtime: ${path}`);
});

// ao-pi#40: the record names its own session. The file is sha256(id)[:16], so the
// name cannot be read back to the id it approves.
test("an approval is written atomically and carries the runtime", () => {
  saveSessionAgent("session-ada", "opencode");
  const record = JSON.parse(readFileSync(sessionPointerPath("session-ada"), "utf8"));
  assert.deepEqual(
    { agent: record.agent, session: record.session, runtime: record.runtime },
    { agent: "opencode", session: "session-ada", runtime: "opencode" },
  );
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

test("revoking the session record is not undone by the legacy file", () => {
  const id = "session-revoke";
  mkdirSync(dirname(legacySessionPointerPath(id)), { recursive: true });
  writeFileSync(legacySessionPointerPath(id), JSON.stringify({ agent: "opencode" }), { mode: 0o600 });
  assert.equal(sessionAgentOf(id), "opencode");
  assert.equal(existsSync(legacySessionPointerPath(id)), false, "the legacy file is retired once copied");
  unlinkSync(sessionPointerPath(id));
  assert.equal(sessionAgentOf(id), undefined, "a deleted approval stays deleted");
});

test("a read-only sessions directory still returns the agent, with nothing copied", () => {
  const id = "session-readonly";
  mkdirSync(dirname(legacySessionPointerPath(id)), { recursive: true });
  writeFileSync(legacySessionPointerPath(id), JSON.stringify({ agent: "opencode" }), { mode: 0o600 });
  const sessions = dirname(sessionPointerPath(id));
  mkdirSync(sessions, { recursive: true });
  chmodSync(sessions, 0o500);
  try {
    assert.equal(sessionAgentOf(id), "opencode", "the identity survives a read-only directory");
    assert.equal(existsSync(sessionPointerPath(id)), false, "and nothing is copied into it");
  } finally {
    chmodSync(sessions, 0o700);
  }
});
