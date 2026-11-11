// One worker per Office task (agent-office#399): worktrees under AO_HOME, a
// task re-read before its claim, and the shared ao.task hc label.
import assert from "node:assert/strict"
import { execFileSync } from "node:child_process"
import { createHash } from "node:crypto"
import { mkdirSync, writeFileSync } from "node:fs"
import { homedir, tmpdir } from "node:os"
import { join, resolve } from "node:path"
import { after, test } from "node:test"
import { addWorktree, aoHome, dispatchRefusal, hcTaskTerminals, hcTaskWorkers, taskLabel, taskWorktreePath } from "../lib/office.ts"
import { removeTempRoot, tempDir } from "./suite.mts"

process.env.AO_HOME = tempDir("ao-home-")
after(removeTempRoot)
const short = (id: string) => createHash("sha256").update(id).digest("hex").slice(0, 8)

function repo() {
  const path = join(tempDir("ao-dispatch-"), "my-repo")
  execFileSync("git", ["init", "-q", path])
  execFileSync("git", ["-C", path, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init"])
  return path
}

// A fake hc: `ls` prints `out` and exits `code`, `get` prints `get` (one
// answer for every terminal) and exits `getCode`.
function fakeHc(out: string, code = 0, get = '{"state":"live","status":{"state":"working"}}', getCode = 0) {
  const bin = join(tempDir("ao-hc-"), "hc")
  writeFileSync(bin, `#!/bin/sh\ncase "$1" in\n  ls) printf '%s' '${out}'; exit ${code} ;;\n  get) printf '%s' '${get}'; exit ${getCode} ;;\nesac\nexit 0\n`, { mode: 0o755 })
  return bin
}
const old = Date.now() - 60 * 60_000
function withHc<T>(bin: string, fn: () => T): T {
  process.env.AGENT_OFFICE_HC_BIN = bin
  try {
    return fn()
  } finally {
    delete process.env.AGENT_OFFICE_HC_BIN
  }
}

test("AO_HOME: unset or empty is ~/.ao, a quoted ~ is the home folder, relative resolves", () => {
  assert.equal(aoHome({}), join(homedir(), ".ao"))
  assert.equal(aoHome({ AO_HOME: "  " }), join(homedir(), ".ao"))
  assert.equal(aoHome({ AO_HOME: "~/.ao" }), join(homedir(), ".ao"))
  assert.equal(aoHome({ AO_HOME: "~" }), homedir())
  assert.equal(aoHome({ AO_HOME: "rel/ao" }), resolve("rel/ao"))
  assert.equal(aoHome({ AO_HOME: "/srv/ao" }), "/srv/ao")
})

test("the task worktree path is ${AO_HOME}/worktrees/<repo>/<short>-<slug>, as ao-claude computes it", () => {
  assert.equal(taskWorktreePath("/x/agent-office", "task-1", "Fix the Add-member link!", { AO_HOME: "/srv/ao" }), `/srv/ao/worktrees/agent-office/${short("task-1")}-fix-the-add-member-link`)
  assert.equal(taskWorktreePath("/x/agent-office", "task-1", "", { AO_HOME: "/srv/ao" }), `/srv/ao/worktrees/agent-office/${short("task-1")}`)
})

test("addWorktree puts the task under AO_HOME on office/<short>, and a retry reuses it", async () => {
  const r = repo()
  const first = await addWorktree(r, "task-a", "Fix it")
  assert.equal(first.path, join(process.env.AO_HOME!, "worktrees", "my-repo", `${short("task-a")}-fix-it`))
  assert.equal(first.branch, `office/${short("task-a")}`)
  assert.equal(execFileSync("git", ["-C", first.path, "branch", "--show-current"], { encoding: "utf8" }).trim(), first.branch)
  // A retry with another title finds the branch's worktree, not a new path.
  assert.deepEqual(await addWorktree(r, "task-a", "Renamed"), first)
})

test("addWorktree never takes over a folder that is not the task's worktree", async () => {
  const r = repo()
  mkdirSync(taskWorktreePath(r, "task-b", "Fix it"), { recursive: true })
  await assert.rejects(addWorktree(r, "task-b", "Fix it"), /exists and is not the office\/[0-9a-f]{8} worktree/)
})

test("hc ls lists the task's live terminals; an unreadable or failing hc cannot say", () => {
  const live = fakeHc('{"terminals":[{"address":"t42","state":"live"}]}')
  assert.deepEqual(withHc(live, () => hcTaskTerminals("task-c")), ["t42"])
  assert.deepEqual(withHc(fakeHc('{"terminals":[]}'), () => hcTaskTerminals("task-c")), [])
  assert.equal(withHc(fakeHc("", 1), () => hcTaskTerminals("task-c")), undefined)
  assert.equal(withHc(fakeHc("not json"), () => hcTaskTerminals("task-c")), undefined)
  assert.equal(withHc(join(tmpdir(), "no-such-hc"), () => hcTaskTerminals("task-c")), undefined)
  assert.equal(taskLabel("task-c"), "ao.task=task-c")
})

test("a labelled terminal holds its task only while working or blocked, or while it is booting", () => {
  const ls = '{"terminals":[{"address":"t42","state":"live"}]}'
  const answer = (state: string, created: number) => `{"state":"live","status":{"app":"opencode","state":"${state}"},"created_ms":${created}}`
  assert.deepEqual(withHc(fakeHc(ls, 0, answer("working", old)), () => hcTaskWorkers("task-w")), ["t42"])
  assert.deepEqual(withHc(fakeHc(ls, 0, answer("blocked", old)), () => hcTaskWorkers("task-w")), ["t42"])
  // A finished session left open frees its task, so a retry can run.
  assert.deepEqual(withHc(fakeHc(ls, 0, answer("done", old)), () => hcTaskWorkers("task-w")), [])
  assert.deepEqual(withHc(fakeHc(ls, 0, answer("idle", old)), () => hcTaskWorkers("task-w")), [])
  // A just-started agent reads idle or done at turn 0: it still holds the task.
  assert.deepEqual(withHc(fakeHc(ls, 0, answer("idle", Date.now())), () => hcTaskWorkers("task-w")), ["t42"])
  assert.deepEqual(withHc(fakeHc(ls, 0, '{"state":"finished"}'), () => hcTaskWorkers("task-w")), [])
  assert.equal(withHc(fakeHc(ls, 0, "garbage", 1), () => hcTaskWorkers("task-w")), undefined)
})

test("a task is dispatched only while todo and with no live ao.task terminal", () => {
  const none = fakeHc('{"terminals":[]}')
  const busy = fakeHc('{"terminals":[{"address":"t7"}]}', 0, `{"state":"live","status":{"state":"working"},"created_ms":${old}}`)
  const doneAndOpen = fakeHc('{"terminals":[{"address":"t7"}]}', 0, `{"state":"live","status":{"state":"done"},"created_ms":${old}}`)
  const broken = fakeHc("", 2)
  assert.equal(withHc(none, () => dispatchRefusal("task-d", { status: "todo" }, true)), undefined)
  assert.match(withHc(none, () => dispatchRefusal("task-d", { status: "in_progress" }, true))!, /in_progress, not todo/)
  assert.match(withHc(none, () => dispatchRefusal("task-d", undefined, true))!, /unreadable, not todo/)
  assert.match(withHc(busy, () => dispatchRefusal("task-d", { status: "todo" }, true))!, /live worker terminal \(t7\)/)
  assert.match(withHc(busy, () => dispatchRefusal("task-d", { status: "todo" }, false))!, /live worker terminal/)
  assert.equal(withHc(doneAndOpen, () => dispatchRefusal("task-d", { status: "todo" }, true)), undefined)
  // hc dispatch waits when hc cannot say; headless goes ahead without hc.
  assert.match(withHc(broken, () => dispatchRefusal("task-d", { status: "todo" }, true))!, /hc could not list/)
  assert.equal(withHc(broken, () => dispatchRefusal("task-d", { status: "todo" }, false)), undefined)
})
