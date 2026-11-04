// Azri's decision on card 53828e66, same rule as ao-pi: a task with repo "" gets
// no worktree, so it cannot land in the session's checkout (agent-office#556).
import assert from "node:assert/strict";
import { spawnSync as spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

import { dispatchTarget, repoTitle, sameRepo, scratchTaskDir } from "../lib/office.ts";

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

test("the hc dispatch reads the task's repo", () => {
  // Source-level, because the bug was a dispatch site that never looked at
  // task.repo at all (agent-office#556).
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8");
  assert.ok(source.includes("dispatchTarget("), "the hc dispatch does not call dispatchTarget");
  // Source-level only: the plugin-level proof is the scratch dispatch test in
  // plugin.test.mts, which asserts the launch and the pointer.
  assert.ok(/task\?\.repo|task\.repo/.test(source), "nothing reads task.repo");
});

test("a task naming a different repository is refused, not dispatched here", () => {
  // vinn's second point: falling back to the session's checkout is how a worker
  // ends up in a repo nobody asked for.
  const root = mkdtempSync(join(tmpdir(), "ao-repo-"));
  const repo = join(root, "mine");
  mkdirSync(repo);
  spawnSync("git", ["-C", repo, "init", "-q"]);
  spawnSync("git", ["-C", repo, "remote", "add", "origin", "git@github.com:az-iar/ao-opencode.git"]);
  assert.equal(sameRepo(repo, "az-iar/ao-opencode"), true);
  assert.equal(sameRepo(repo, "Az-Iar/AO-OpenCode.git"), true, "case and .git do not matter");
  assert.equal(sameRepo(repo, "az-iar/agent-office"), false, "a different repo is not this one");
  const bare = join(root, "no-origin");
  mkdirSync(bare);
  spawnSync("git", ["-C", bare, "init", "-q"]);
  assert.equal(sameRepo(bare, "az-iar/anything"), true, "a checkout with no origin cannot disagree");
  assert.equal(repoTitle(repo), "az-iar/ao-opencode");
});
