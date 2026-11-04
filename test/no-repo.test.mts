// Azri's decision on card 53828e66, same rule as ao-pi: a task with repo "" gets
// no worktree, so it cannot land in the session's checkout (agent-office#556).
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { dispatchTarget, scratchTaskDir } from "../lib/office.ts";

test("a repo-less task gets a scratch directory, never the session's repo", () => {
  assert.deepEqual(dispatchTarget(""), { kind: "scratch" });
});

test("whitespace is not a repo either", () => {
  assert.deepEqual(dispatchTarget("   "), { kind: "scratch" });
});

test("a named repo takes the worktree of the checkout it resolved to", () => {
  // Not the session's checkout: dispatchTarget is handed the resolution, and the
  // resolver is the only thing that decides which checkout that is (#558).
  assert.deepEqual(dispatchTarget("az-iar/ao-pi", { path: "/work/ao-pi" }), { kind: "worktree", repo: "/work/ao-pi" });
});

test("a named repo that resolved to nothing is refused with the resolver's reason", () => {
  assert.deepEqual(dispatchTarget("az-iar/elsewhere", { error: "no local checkout of az-iar/elsewhere beside the session's checkout" }), {
    kind: "refused",
    reason: "no local checkout of az-iar/elsewhere beside the session's checkout",
  });
});

test("the scratch directory is under the agent's state, not under any repo", () => {
  const dir = scratchTaskDir("1efe615b-b1ad-4689-9596-85225573406b");
  assert.ok(dir.includes("/agent-office/opencode/scratch/"), dir);
  assert.ok(!dir.includes("/projects/"), dir);
});

test("both dispatch sites resolve the task's repo and call dispatchTarget", () => {
  // Source-level, because the bug was a dispatch site that never looked at
  // task.repo at all (agent-office#556), and ao-pi#18 found the second site still
  // dispatching in the poller's checkout after the first was fixed.
  for (const file of ["../index.ts", "../poll.ts"]) {
    const source = readFileSync(new URL(file, import.meta.url), "utf8");
    assert.ok(source.includes("dispatchTarget("), `${file} does not call dispatchTarget`);
    assert.ok(source.includes("resolveDispatchRepo("), `${file} does not resolve the task's repo`);
    assert.ok(/task\?\.repo|task\.repo/.test(source), `${file} never reads task.repo`);
    assert.ok(!/addWorktree\(repo,/.test(source), `${file} still dispatches in the session's checkout`);
  }
});
