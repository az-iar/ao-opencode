// Azri's decision on card 53828e66, same rule as ao-pi: a task with repo "" gets
// no worktree, so it cannot land in the session's checkout (agent-office#556).
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { dispatchTarget, scratchTaskDir } from "../lib/office.ts";

test("a repo-less task gets a scratch directory, never the session's repo", () => {
  assert.deepEqual(dispatchTarget("/home/azri/projects/personal/ao-opencode", ""), { kind: "scratch" });
});

test("whitespace is not a repo either", () => {
  assert.deepEqual(dispatchTarget("/repo", "   "), { kind: "scratch" });
});

test("a named repo keeps the worktree path", () => {
  assert.deepEqual(dispatchTarget("/repo", "az-iar/ao-pi"), { kind: "worktree", repo: "/repo" });
});

test("the scratch directory is under the agent's state, not under any repo", () => {
  const dir = scratchTaskDir("1efe615b-b1ad-4689-9596-85225573406b");
  assert.ok(dir.includes("/agent-office/opencode/scratch/"), dir);
  assert.ok(!dir.includes("/projects/"), dir);
});

test("both dispatch sites read the task's repo", () => {
  // Source-level, because the bug was a dispatch site that never looked at
  // task.repo at all (agent-office#556).
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8");
  assert.ok(source.includes("dispatchTarget("), "the hc dispatch does not call dispatchTarget");
  assert.ok(/task\?\.repo|task\.repo/.test(source), "nothing reads task.repo");
});
