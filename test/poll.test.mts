import assert from "node:assert/strict"
import { execFileSync, spawn } from "node:child_process"
import { createHash } from "node:crypto"
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
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
  const credentials = join(root, "store", "opencode.json")
  mkdirSync(join(root, "store"))
  writeFileSync(credentials, JSON.stringify({ office: url, client_id: "client-1", access_token: "access-poll", refresh_token: "refresh-poll", expires_at: Date.now() + 3_600_000 }))
  const stub = join(root, "opencode")
  writeFileSync(stub, `#!${process.execPath}\nrequire("fs").writeFileSync(require("path").join(${JSON.stringify(runs)}, String(Date.now()) + "-" + process.pid + ".json"), JSON.stringify({ cwd: process.cwd(), argv: process.argv.slice(2) }))\nprocess.exit(Number(process.env.STUB_EXIT ?? 0))\n`)
  chmodSync(stub, 0o755)
  const env: Record<string, string> = { ...(process.env as Record<string, string>), HOME: root, AGENT_OFFICE_URL: url, AGENT_OFFICE_OPENCODE_CREDENTIALS: credentials, AGENT_OFFICE_OPENCODE_BIN: stub }
  delete env.AGENT_OFFICE_TOKEN
  delete env.AGENT_OFFICE_AUTODISPATCH
  const poll = (args: string[] = [], extra: Record<string, string> = {}) =>
    new Promise<{ code: number | null; stderr: string }>((resolve) => {
      const child = spawn(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", pollScript, "--dir", dir, ...args], { env: { ...env, ...extra }, stdio: ["ignore", "ignore", "pipe"] })
      let stderr = ""
      child.stderr.on("data", (chunk) => (stderr += String(chunk)))
      child.on("close", (code) => resolve({ code, stderr }))
    })
  const recorded = () => readdirSync(runs).sort().map((file) => JSON.parse(readFileSync(join(runs, file), "utf8")) as { cwd: string; argv: string[] })
  const state = (kind: string) => join(root, "store", kind, "opencode")
  return { dir, poll, recorded, state }
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

test("a new todo task is claimed, run as a supervisor in its worktree, and its dispatch lock released", async () => {
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
  const worktree = `${ws.dir}-office-${short}`
  assert.equal(execFileSync("git", ["-C", worktree, "branch", "--show-current"], { encoding: "utf8" }).trim(), `office/${short}`)
  assert.equal(run.cwd, worktree)
  assert.deepEqual(run.argv.slice(0, 5), ["run", "--dir", worktree, "--title", "Office task task-new"])
  assert.match(run.argv[5], /^Agent Office task task-new\. You are its supervisor/)
  assert.ok(run.argv[5].includes(`git worktree ${worktree} on branch office/${short}`))
  assert.match(run.argv[5], /Before you start any work, post one short comment on the task with post_comment/)
  assert.doesNotMatch(run.argv[5], /post_message to azri/)
  const update = office.calls.find((c) => c.body?.params?.name === "update_task")
  assert.equal(update?.body.params.arguments.task_id, "task-new")
  assert.equal(update?.body.params.arguments.status, "in_progress")
  assert.deepEqual(readdirSync(ws.state("dispatch")), [])
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
