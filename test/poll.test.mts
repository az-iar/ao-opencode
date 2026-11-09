import assert from "node:assert/strict"
import { execFileSync, spawn } from "node:child_process"
import { createHash } from "node:crypto"
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join, dirname } from "node:path"
import { after, before, test } from "node:test"
import { fileURLToPath } from "node:url"
import { fakeOffice } from "./fake-office.mts"

const office = fakeOffice()
const pollScript = fileURLToPath(new URL("../poll.ts", import.meta.url))
let url = ""

before(async () => {
  url = await office.listen()
})
after(() => office.close())

// A git workspace with its own credentials, claim and state directories, and
// an `opencode` stub that records each run's cwd and argv as JSON.
function workspace() {
  const root = mkdtempSync(join(tmpdir(), "ao-poll-"))
  const dir = join(root, "repo")
  const runs = join(root, "runs")
  mkdirSync(dir)
  mkdirSync(runs)
  execFileSync("git", ["-C", dir, "init", "-q"])
  execFileSync("git", ["-C", dir, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init"])
  // A named repo is resolved by origin, and a checkout with none keeps #32's
  // fallback (it runs the task), so the poller's own checkout needs one for the
  // refusal cases to be refusals (agent-office#558).
  execFileSync("git", ["-C", dir, "remote", "add", "origin", "git@github.com:az-iar/ao-opencode.git"])
  const credentials = join(root, "store", "opencode.json")
  mkdirSync(join(root, "store"))
  writeFileSync(credentials, JSON.stringify({ office: url, client_id: "client-1", access_token: "access-poll", refresh_token: "refresh-poll", expires_at: Date.now() + 3_600_000 }))
  const stub = join(root, "opencode")
  writeFileSync(stub, `#!${process.execPath}\nrequire("fs").writeFileSync(require("path").join(${JSON.stringify(runs)}, String(Date.now()) + "-" + process.pid + ".json"), JSON.stringify({ cwd: process.cwd(), argv: process.argv.slice(2) }))\nprocess.exit(Number(process.env.STUB_EXIT ?? 0))\n`)
  chmodSync(stub, 0o755)
  const env: Record<string, string> = { ...(process.env as Record<string, string>), HOME: root, AGENT_OFFICE_URL: url, AGENT_OFFICE_OPENCODE_CREDENTIALS: credentials, AGENT_OFFICE_OPENCODE_BIN: stub }
  delete env.AGENT_OFFICE_TOKEN
  delete env.AGENT_OFFICE_AUTODISPATCH
  // Worktrees land under this HOME's ~/.ao; no real hc is ever asked.
  delete env.AO_HOME
  env.AGENT_OFFICE_HC_BIN = join(root, "no-hc")
  const poll = (args: string[] = [], extra: Record<string, string> = {}) =>
    new Promise<{ code: number | null; stderr: string }>((resolve) => {
      const child = spawn(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", pollScript, "--dir", dir, ...args], { env: { ...env, ...extra }, stdio: ["ignore", "ignore", "pipe"] })
      let stderr = ""
      child.stderr.on("data", (chunk) => (stderr += String(chunk)))
      child.on("close", (code) => resolve({ code, stderr }))
    })
  const recorded = () => readdirSync(runs).sort().map((file) => JSON.parse(readFileSync(join(runs, file), "utf8")) as { cwd: string; argv: string[] })
  const state = (kind: string) => join(root, "store", kind, "opencode")
  // Where a task's worktree goes: ${AO_HOME:-~/.ao}/worktrees/<repo>/<short>-<slug>.
  const worktreeFor = (id: string, slug: string) => join(root, ".ao", "worktrees", "repo", `${createHash("sha256").update(id).digest("hex").slice(0, 8)}-${slug}`)
  return { dir, poll, recorded, state, worktreeFor }
}

const reset = () => {
  office.state.openTasks = [{ id: "task-open-1", title: "Open task" }]
  office.state.unclaimedTasks = []
  office.state.unreadMessages = [
    { id: "message-old-1", from: "azri", subject: "Old one" },
    { id: "message-old-2", from: "azri", subject: "Old two" },
  ]
  office.calls.length = 0
}
const toolCalls = () => office.calls.filter((c) => c.path === "/mcp").map((c) => c.body.params.name)

test("not connected: a hint and a non-zero exit", async () => {
  const ws = workspace()
  const result = await ws.poll([], { AGENT_OFFICE_OPENCODE_CREDENTIALS: join(ws.dir, "missing.json") })
  assert.notEqual(result.code, 0)
  assert.match(result.stderr, /run \/office in OpenCode first/)
  assert.deepEqual(ws.recorded(), [])
})

test("first run handles unread mail but baselines tasks; a new message runs exactly once", async () => {
  reset()
  const ws = workspace()
  const first = await ws.poll()
  assert.equal(first.code, 0, first.stderr)
  assert.equal(ws.recorded().length, 1)
  assert.ok(!toolCalls().includes("update_task"), "a pre-existing task is baseline")
  const [inbox] = ws.recorded()
  assert.equal(inbox.cwd, ws.dir)
  assert.deepEqual(inbox.argv.slice(0, 5), ["run", "--dir", ws.dir, "--title", "Office inbox"])
  assert.match(inbox.argv[5], /^Agent Office inbox: 2 new message\(s\)/)

  assert.equal((await ws.poll()).code, 0)
  assert.equal(ws.recorded().length, 1, "nothing new, nothing run")

  office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-self", from: "opencode", subject: "mine" }, { id: "message-new", from: "azri", subject: "New work" }]
  const third = await ws.poll()
  assert.equal(third.code, 0, third.stderr)
  assert.equal(ws.recorded().length, 2)
  const prompt = ws.recorded()[1].argv[5]
  assert.match(prompt, /^Agent Office inbox: 1 new message\(s\)/)
  assert.match(prompt, /message:message-new/)
  assert.doesNotMatch(prompt, /message-old-1|message-self/)
  assert.match(prompt, /Never read, print, copy or send credentials/)
  assert.ok(!third.stderr.includes("access-poll"), "no token in the log")
  assert.ok(!third.stderr.includes("New work"), "no Office text in the log")
})

test("a failed run releases the claim and retries next time", async () => {
  reset()
  const ws = workspace()
  const failed = await ws.poll([], { STUB_EXIT: "1" })
  assert.match(failed.stderr, /inbox run failed/)
  assert.deepEqual(readdirSync(ws.state("notices")), [])
  await ws.poll()
  assert.equal(ws.recorded().length, 2)
})

test("a notice claim held by the plugin prevents a run", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  const { claimNotice } = await import("../lib/office.ts")
  assert.ok(claimNotice(ws.state("notices"), ["message:message-held"]))
  office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-held", from: "azri", subject: "Held" }]
  const held = await ws.poll()
  assert.match(held.stderr, /another OpenCode is handling these messages/)
  assert.equal(ws.recorded().length, 1)
  await ws.poll()
  assert.equal(ws.recorded().length, 1, "a claimed batch is seen")
})

test("a new todo task is claimed, run as one worker in its worktree, and its dispatch lock released", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  office.state.openTasks = [...office.state.openTasks, { id: "task-off", title: "Not now", status: "todo" }]
  const off = await ws.poll([], { AGENT_OFFICE_AUTODISPATCH: "off" })
  assert.match(off.stderr, /auto-dispatch is off/)
  assert.equal(ws.recorded().length, 1)

  office.state.openTasks = [...office.state.openTasks, { id: "task-new", title: "Fix it", status: "todo" }]
  const result = await ws.poll()
  assert.equal(result.code, 0, result.stderr)
  assert.equal(ws.recorded().length, 2)
  const run = ws.recorded()[1]
  const short = createHash("sha256").update("task-new").digest("hex").slice(0, 8)
  const worktree = ws.worktreeFor("task-new", "fix-it")
  assert.equal(execFileSync("git", ["-C", worktree, "branch", "--show-current"], { encoding: "utf8" }).trim(), `office/${short}`)
  assert.equal(run.cwd, worktree)
  assert.deepEqual(run.argv.slice(0, 5), ["run", "--dir", worktree, "--title", "Office task task-new"])
  assert.match(run.argv[5], /^Agent Office task task-new\. You are its worker/)
  assert.ok(run.argv[5].includes(`git worktree ${worktree} on branch office/${short}`))
  assert.match(run.argv[5], /Before you start any work, post one short comment on the task with post_comment/)
  assert.doesNotMatch(run.argv[5], /post_message to azri/)
  const update = office.calls.find((c) => c.body?.params?.name === "update_task")
  assert.equal(update?.body.params.arguments.task_id, "task-new")
  assert.equal(update?.body.params.arguments.status, "in_progress")
  assert.deepEqual(readdirSync(ws.state("dispatch")), [])
})

// agent-office#558: the cron path dispatched every task in the poller's checkout,
// named repo or not, so a task in another checkout landed in the wrong one. Both
// tests here are red on main: there the worktree is of `repo` either way.
test("a task naming another local checkout runs in a worktree of THAT checkout", async () => {
  reset()
  const ws = workspace()
  // A checkout beside the poller's, named by the task.
  const other = join(dirname(ws.dir), "agent-office")
  mkdirSync(other)
  execFileSync("git", ["-C", other, "init", "-q"])
  execFileSync("git", ["-C", other, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init"])
  execFileSync("git", ["-C", other, "remote", "add", "origin", "https://github.com/az-iar/agent-office.git"])
  await ws.poll()
  office.calls.length = 0
  const id = "task-other-repo"
  const task = { id, title: "Fix it", status: "todo", repo: "az-iar/agent-office" }
  office.state.taskDetails.set(id, task)
  office.state.openTasks = [...office.state.openTasks, task]
  try {
    const result = await ws.poll()
    assert.equal(result.code, 0, result.stderr)
    const short = createHash("sha256").update(id).digest("hex").slice(0, 8)
    const run = ws.recorded().at(-1)!
    assert.equal(run.cwd, join(dirname(ws.dir), ".ao", "worktrees", "agent-office", `${short}-fix-it`))
    assert.equal(
      execFileSync("git", ["-C", run.cwd, "rev-parse", "--path-format=absolute", "--git-common-dir"], { encoding: "utf8" }).trim(),
      join(other, ".git"),
      "the worktree belongs to the checkout the task names",
    )
  } finally {
    office.state.taskDetails.delete(id)
  }
})

test("a task naming a repo with no local checkout is refused, not dispatched here", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  const id = "task-nowhere-repo"
  const task = { id, title: "Fix it", status: "todo", repo: "az-iar/elsewhere" }
  office.state.taskDetails.set(id, task)
  office.state.openTasks = [...office.state.openTasks, task]
  try {
    const before = ws.recorded().length
    const result = await ws.poll()
    assert.equal(result.code, 0, result.stderr)
    assert.equal(ws.recorded().length, before, "no worker is started in the poller's checkout instead")
    const blocked = office.calls.find((c) => c.body?.params?.name === "update_task" && c.body.params.arguments?.task_id === id && c.body.params.arguments?.status === "blocked")
    assert.match(String(blocked?.body.params.arguments?.reason ?? ""), /no local checkout of az-iar\/elsewhere/)
  } finally {
    office.state.taskDetails.delete(id)
  }
})

test("--dry-run reports ids only, writes nothing and starts nothing", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-dry", from: "azri", subject: "Secret subject" }]
  office.state.openTasks = [...office.state.openTasks, { id: "task-dry", title: "Secret title", status: "todo" }]
  const dry = await ws.poll(["--dry-run"])
  assert.equal(dry.code, 0, dry.stderr)
  assert.match(dry.stderr, /would handle message\(s\) message:message-dry/)
  assert.match(dry.stderr, /would dispatch task task-dry/)
  assert.doesNotMatch(dry.stderr, /Secret/)
  assert.deepEqual(toolCalls(), ["checkin"])
  assert.equal(ws.recorded().length, 1)
  assert.equal(existsSync(ws.state("dispatch")), false)

  await ws.poll()
  assert.equal(ws.recorded().length, 3, "the dry run left the items new")
})

test("an overlapping run exits at once", async () => {
  reset()
  const ws = workspace()
  mkdirSync(ws.state("poll"), { recursive: true })
  writeFileSync(join(ws.state("poll"), "run.lock"), JSON.stringify({ pid: process.pid, created: Date.now() }))
  const result = await ws.poll()
  assert.equal(result.code, 0)
  assert.match(result.stderr, /another poll is still running/)
  assert.deepEqual(ws.recorded(), [])
  assert.equal(existsSync(join(ws.state("poll"), "seen.json")), false)
})

const taskUpdates = (id: string) => office.calls.filter((c) => c.body?.params?.name === "update_task" && c.body.params.arguments?.task_id === id).map((c) => c.body.params.arguments)
const short = (id: string) => createHash("sha256").update(id).digest("hex").slice(0, 8)

test("a claimed task whose worktree fails is blocked on its creating agent and not claimed again", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  // A folder that is not the task's worktree sits at its path: never taken over.
  mkdirSync(ws.worktreeFor("task-wt", "fix-it"), { recursive: true })
  office.state.taskDetails.set("task-wt", { id: "task-wt", title: "Fix it", status: "todo", created_by: "agent:planner" })
  office.state.openTasks = [...office.state.openTasks, { id: "task-wt", title: "Fix it", status: "todo" }]
  const result = await ws.poll()
  assert.equal(result.code, 0, result.stderr)
  assert.equal(ws.recorded().length, 1, "no worker run")
  const [claim, block] = taskUpdates("task-wt")
  assert.equal(claim.status, "in_progress")
  assert.equal(block.status, "blocked")
  assert.equal(block.waiting_on, "planner")
  assert.match(block.reason, /^OpenCode could not start this task: Error: .* exists and is not the office\/[0-9a-f]{8} worktree/)
  assert.deepEqual(readdirSync(ws.state("dispatch")), [])
  await ws.poll()
  assert.equal(taskUpdates("task-wt").length, 2, "never set back to todo or claimed again")
})

test("a failed opencode start is blocked with a masked reason, on Azri when the Office refuses the creator", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  office.state.taskDetails.set("task-ghost", { id: "task-ghost", title: "From a gone agent", status: "todo", created_by: "agent:ghost" })
  office.state.openTasks = [...office.state.openTasks, { id: "task-ghost", title: "From a gone agent", status: "todo" }]
  office.state.refuse = (name, args) => name === "update_task" && args?.waiting_on === "ghost"
  try {
    const result = await ws.poll([], { AGENT_OFFICE_OPENCODE_BIN: join(ws.dir, "missing", "opencode") })
    assert.equal(result.code, 0, result.stderr)
  } finally {
    office.state.refuse = undefined
  }
  const blocked = taskUpdates("task-ghost").filter((update) => update.status === "blocked")
  assert.deepEqual(blocked.map((update) => update.waiting_on), ["ghost", "azri"])
  assert.equal(blocked[1].reason, "OpenCode could not start this task: Error: spawn <path> ENOENT")
  assert.deepEqual(readdirSync(ws.state("dispatch")), [])
})

test("an HTTP error on the blocked update is logged, not retried on Azri, and a refused claim is not blocked", async () => {
  reset()
  const ws = workspace()
  await ws.poll()
  office.calls.length = 0
  for (const id of ["task-5xx", "task-taken"]) {
    mkdirSync(ws.worktreeFor(id, id), { recursive: true })
    office.state.taskDetails.set(id, { id, title: id, status: "todo", created_by: "agent:planner" })
  }
  office.state.openTasks = [...office.state.openTasks, ...["task-5xx", "task-taken"].map((id) => ({ id, title: id, status: "todo" }))]
  office.state.httpError = (name, args) => (name === "update_task" && args?.status === "blocked" ? 502 : undefined)
  office.state.refuse = (name, args) => name === "update_task" && args?.task_id === "task-taken"
  try {
    const result = await ws.poll()
    assert.equal(result.code, 0, result.stderr)
    assert.match(result.stderr, /could not mark task task-5xx blocked: Error: Office update_task: HTTP 502/)
  } finally {
    office.state.httpError = undefined
    office.state.refuse = undefined
  }
  assert.deepEqual(taskUpdates("task-5xx").map((update) => update.waiting_on ?? update.status), ["in_progress", "planner"])
  assert.deepEqual(taskUpdates("task-taken").map((update) => update.status), ["in_progress"])
  assert.deepEqual(readdirSync(ws.state("dispatch")), [])
})

test("a bridge agent starts no inbox run, records the mail as handled and marks nothing read", async () => {
  reset()
  office.state.agent = { ...office.state.agent, delivery_mode: "bridge" }
  const ws = workspace()
  try {
    const first = await ws.poll()
    assert.equal(first.code, 0, first.stderr)
    assert.match(first.stderr, /stands down/)
    assert.deepEqual(ws.recorded(), [], "no headless inbox run")
    assert.ok(!toolCalls().includes("mark_read"), "nothing is marked read on the server")
    const seen = JSON.parse(readFileSync(join(ws.state("poll"), "seen.json"), "utf8"))
    assert.deepEqual(seen.messages.sort(), ["message:message-old-1", "message:message-old-2"].sort())
    // Once per change of mode, not on every run.
    const second = await ws.poll()
    assert.equal(second.code, 0, second.stderr)
    assert.doesNotMatch(second.stderr, /stands down/)
    assert.deepEqual(ws.recorded(), [])
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
  }
})

test("a bridge agent still dispatches an assigned todo task", async () => {
  reset()
  office.state.agent = { ...office.state.agent, delivery_mode: "bridge" }
  const ws = workspace()
  try {
    await ws.poll() // the task baseline
    office.state.openTasks = [...office.state.openTasks, { id: "task-bridge-poll", title: "Fix it", status: "todo" }]
    const result = await ws.poll()
    assert.equal(result.code, 0, result.stderr)
    const runs = ws.recorded()
    assert.equal(runs.length, 1, "one task run and no inbox run")
    assert.deepEqual(runs[0].argv.slice(0, 5), ["run", "--dir", ws.worktreeFor("task-bridge-poll", "fix-it"), "--title", "Office task task-bridge-poll"])
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
  }
})

test("flipping a bridge agent back to poll delivers a message arriving after the flip", async () => {
  reset()
  office.state.agent = { ...office.state.agent, delivery_mode: "bridge" }
  const ws = workspace()
  try {
    await ws.poll() // bridge records the unread mail as handled
    assert.deepEqual(ws.recorded(), [])
    office.state.agent = { ...office.state.agent, delivery_mode: "poll" }
    office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-after-flip", from: "azri", subject: "After" }]
    const result = await ws.poll()
    assert.equal(result.code, 0, result.stderr)
    assert.match(result.stderr, /handles mail again/)
    assert.equal(ws.recorded().length, 1)
    const prompt = ws.recorded()[0].argv[5]
    assert.match(prompt, /message-after-flip/)
    assert.doesNotMatch(prompt, /message-old-1/, "mail the bridge handled is not delivered again")
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
  }
})

for (const mode of [undefined, "poll", "webhook", "agent37"] as const) {
  test(`delivery mode ${mode ?? "(missing)"} keeps today's inbox run`, async () => {
    reset()
    office.state.agent = { slug: "opencode", name: "OpenCode", delivery_mode: mode }
    const ws = workspace()
    try {
      const result = await ws.poll()
      assert.equal(result.code, 0, result.stderr)
      assert.equal(ws.recorded().length, 1, "the inbox run happens as before")
      assert.match(ws.recorded()[0].argv[5], /^Agent Office inbox: 2 new message\(s\)/)
    } finally {
      office.state.agent = { slug: "opencode", name: "OpenCode" }
    }
  })
}
