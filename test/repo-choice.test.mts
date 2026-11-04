// A task naming a repo runs in THAT repository, never in the session's
// (agent-office#558). The match is the checkout's origin, and a zero or two-plus
// match is refused rather than guessed: the same rule ao-pi carries (ao-pi#17-#19).
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { test } from "node:test";

import { chooseRepo, dedupeCandidates, dispatchTarget, repoFromRemote, resolveDispatchRepo, siblingCheckouts } from "../lib/office.ts";

test("repoFromRemote reads owner/name out of a remote and normalises it", () => {
  assert.equal(repoFromRemote("https://github.com/az-iar/agent-office.git"), "az-iar/agent-office");
  assert.equal(repoFromRemote("git@github.com:az-iar/ao-pi.git"), "az-iar/ao-pi");
  assert.equal(repoFromRemote("ssh://git@github.com/az-iar/ao-pi"), "az-iar/ao-pi");
  assert.equal(repoFromRemote("  https://GitHub.com/Az-Iar/Agent-Office/  "), "az-iar/agent-office");
  assert.equal(repoFromRemote("/srv/git/agent-office"), undefined, "a bare path names no owner");
});

test("chooseRepo takes the one checkout whose remote is the task's repo", () => {
  const candidates = [
    { path: "/work/ao-pi", remote: "git@github.com:az-iar/ao-pi.git" },
    { path: "/work/agent-office", remote: "https://github.com/az-iar/agent-office.git" },
  ];
  assert.deepEqual(chooseRepo("az-iar/agent-office", candidates), { path: "/work/agent-office" });
  // A task field may carry the URL, a different case or a trailing .git.
  assert.deepEqual(chooseRepo("https://github.com/Az-Iar/Agent-Office.git", candidates), { path: "/work/agent-office" });
  assert.deepEqual(chooseRepo("Az-Iar/AO-Pi", candidates), { path: "/work/ao-pi" });
});

test("chooseRepo refuses on no match, naming the repo", () => {
  const chosen = chooseRepo("az-iar/elsewhere", [{ path: "/work/ao-pi", remote: "git@github.com:az-iar/ao-pi.git" }]);
  assert.ok("error" in chosen);
  assert.match(chosen.error, /no local checkout of az-iar\/elsewhere/);
});

test("chooseRepo refuses on two matches, naming both checkouts", () => {
  const chosen = chooseRepo("az-iar/agent-office", [
    { path: "/work/one", remote: "https://github.com/az-iar/agent-office.git" },
    { path: "/work/two", remote: "git@github.com:az-iar/agent-office.git" },
  ]);
  assert.ok("error" in chosen);
  assert.match(chosen.error, /^2 checkouts of az-iar\/agent-office/);
  assert.match(chosen.error, /\bone\b.*\btwo\b|two.*one/);
  // Directory names, not paths: a blocked task's reason scrubs every absolute path
  // to <path>, so paths there would name neither candidate.
  assert.doesNotMatch(chosen.error, /\/work\//);
});

test("resolveDispatchRepo looks at the session's own checkout first", () => {
  const session = { path: "/work/ao-opencode", remote: "git@github.com:az-iar/ao-opencode.git" };
  const candidates = [session, { path: "/work/agent-office", remote: "https://github.com/az-iar/agent-office.git" }];
  assert.deepEqual(resolveDispatchRepo("/work/ao-opencode", "az-iar/ao-opencode", () => candidates), { path: "/work/ao-opencode" });
  assert.deepEqual(resolveDispatchRepo("/work/ao-opencode", "az-iar/agent-office", () => candidates), { path: "/work/agent-office" });
  // A repo-less task keeps the session's checkout here; ao-opencode's dispatch
  // turns that into a scratch directory before this is ever called.
  assert.deepEqual(resolveDispatchRepo("/work/ao-opencode", "", () => candidates), { path: "/work/ao-opencode" });
});

// vinn's rule in ao-opencode#32, kept by Azri on #558: a checkout with no origin
// cannot disagree with the task, so it still runs a task naming another repo -
// but says so, and a matching sibling still wins.
test("a named repo in an origin-less checkout runs there, with a note saying why", () => {
  const chosen = resolveDispatchRepo("/work/repo", "az-iar/elsewhere", () => [{ path: "/work/repo" }]);
  assert.ok("path" in chosen);
  assert.equal(chosen.path, "/work/repo");
  assert.match(String(chosen.note), /No checkout matched az-iar\/elsewhere and this checkout has no origin/);
  // dispatchTarget carries the note to the dispatch notice, so it is not silent.
  assert.deepEqual(dispatchTarget("az-iar/elsewhere", chosen), { kind: "worktree", repo: "/work/repo", note: chosen.note });
});

test("a matching sibling wins over the origin-less checkout", () => {
  const candidates = [{ path: "/work/repo" }, { path: "/work/agent-office", remote: "git@github.com:az-iar/agent-office.git" }];
  assert.deepEqual(resolveDispatchRepo("/work/repo", "az-iar/agent-office", () => candidates), { path: "/work/agent-office" });
});

test("two matches are still refused when the session checkout has no origin", () => {
  const candidates = [
    { path: "/work/repo" },
    { path: "/work/a", remote: "git@github.com:az-iar/agent-office.git" },
    { path: "/work/b", remote: "https://github.com/az-iar/agent-office.git" },
  ];
  const chosen = resolveDispatchRepo("/work/repo", "az-iar/agent-office", () => candidates);
  assert.ok("error" in chosen);
  assert.match(chosen.error, /^2 checkouts of az-iar\/agent-office/);
});

test("an origin-less checkout does not absorb a task whose checkout is elsewhere with an origin of its own", () => {
  // The fallback is only for zero matches; a session WITH an origin refuses instead.
  const candidates = [{ path: "/work/repo", remote: "git@github.com:az-iar/ao-opencode.git" }];
  const chosen = resolveDispatchRepo("/work/repo", "az-iar/elsewhere", () => candidates);
  assert.ok("error" in chosen);
  assert.match(chosen.error, /no local checkout of az-iar\/elsewhere/);
});

test("siblingCheckouts finds the session's checkout and the ones beside it, and counts a worktree once", () => {
  const root = mkdtempSync(join(tmpdir(), "ao-siblings-"));
  const session = join(root, "ao-opencode");
  mkdirSync(session);
  const git = (path: string, ...args: string[]) => execFileSync("git", ["-C", path, ...args], { encoding: "utf8" });
  git(session, "init", "-q");
  git(session, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init");
  git(session, "remote", "add", "origin", "git@github.com:az-iar/ao-opencode.git");
  // A linked worktree of the same repository sits beside it in this fixture: it
  // shares the origin, so only the common dir tells it from a second clone.
  const linked = join(root, "ao-opencode-task");
  git(session, "worktree", "add", "-q", "-b", "task", linked);
  const other = join(root, "agent-office");
  mkdirSync(other);
  git(other, "init", "-q");
  git(other, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init");
  git(other, "remote", "add", "origin", "https://github.com/az-iar/agent-office.git");

  const scanned = siblingCheckouts(session);
  assert.ok(scanned.some((candidate) => candidate.path === session));
  assert.ok(scanned.some((candidate) => candidate.path === other));
  const deduped = dedupeCandidates(scanned);
  assert.equal(deduped.filter((candidate) => repoFromRemote(candidate.remote ?? "") === "az-iar/ao-opencode").length, 1);
  // The main checkout wins over its worktree, so the worktree is not rooted there.
  assert.equal(deduped.find((candidate) => repoFromRemote(candidate.remote ?? "") === "az-iar/ao-opencode")?.path, session);
  // The dispatch for another repository lands in that checkout, not this one.
  assert.deepEqual(resolveDispatchRepo(session, "az-iar/agent-office", siblingCheckouts), { path: other });
  assert.deepEqual(resolveDispatchRepo(session, "az-iar/ao-opencode", siblingCheckouts), { path: session });
  // And the dispatch target is the resolved checkout, under the AO_HOME layout.
  assert.deepEqual(dispatchTarget("az-iar/agent-office", resolveDispatchRepo(session, "az-iar/agent-office", siblingCheckouts)), {
    kind: "worktree",
    repo: join(root, basename(other)),
  });
});
