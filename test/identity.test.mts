// A linked worktree nobody connected must not take on the main checkout's
// Office identity: no repo-pointer or legacy per-repo token fallback, no
// office MCP server, no polling, no status reports and no dispatch. A
// `/office` run there, or a task worktree its dispatcher adopted, connects it.
import assert from "node:assert/strict"
import { execFileSync, spawn } from "node:child_process"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { after, test } from "node:test"
import { fileURLToPath } from "node:url"
import { fakeOffice } from "./fake-office.mts"

const pollScript = fileURLToPath(new URL("../poll.ts", import.meta.url))

const office = fakeOffice()
const tick = () => new Promise((r) => setTimeout(r, 20))
const until = async (check: () => boolean, what: string) => {
  for (let i = 0; i < 250; i++) {
    if (check()) return
    await tick()
  }
  assert.fail(`timed out waiting for ${what}`)
}

// Credentials, agent pointers and dispatch state live under $HOME, so this
// suite gets a home of its own; $HOME must be set before lib/office.ts loads.
const home = mkdtempSync(join(tmpdir(), "ao-identity-"))
process.env.HOME = home
process.env.AO_HOME = join(home, "ao-home")
process.env.AGENT_OFFICE_CALLBACK_PORT = "45321"
process.env.AGENT_OFFICE_URL = await office.listen()
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_OPENCODE_CREDENTIALS
delete process.env.AGENT_OFFICE_TOKEN_SCOPE
delete process.env.AGENT_OFFICE_HC_BIN
delete process.env.HERDR_ENV

const lib = await import("../lib/office.ts")
const {
  adoptWorktree,
  agentCredentialsPath,
  agentPointerPath,
  agentPointerVia,
  credentialsPath,
  isUnconnectedWorktree,
  loadAgentPointer,
  loadCredentials,
  resolveCredentials,
  saveAgentPointer,
  saveCredentials,
} = lib
const { default: plugin } = await import("../index.ts")

const mcpCalls = () => office.calls.filter((c) => c.path === "/mcp").map((c) => `${c.body?.params?.name} ${JSON.stringify(c.body?.params?.arguments)}`)

function repo() {
  const path = join(mkdtempSync(join(tmpdir(), "ao-identity-")), "repo")
  mkdirSync(path)
  execFileSync("git", ["init", "-q", path])
  execFileSync("git", ["-C", path, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init"])
  return path
}

// A linked worktree of `r`: its git dir lives inside the common one.
function worktree(r: string, branch = "feature") {
  const path = `${r}-${branch}`
  execFileSync("git", ["-C", r, "worktree", "add", "-q", "-b", branch, path])
  return path
}

const credentials = (agent: string, access: string, refresh = `refresh-${agent}`) => ({
  office: process.env.AGENT_OFFICE_URL!,
  client_id: `client-${agent}`,
  access_token: access,
  refresh_token: refresh,
  expires_at: Date.now() + 3_600_000,
})

// What `/office` leaves behind in a workspace: the agent's token under its
// name, and the workspace's pointer to that agent. `via` is the mark a linked
// worktree's pointer needs; the checkout's is written unmarked, as it always
// was.
function connect(directory: string, agent: string, access: string, via?: "office" | "dispatch") {
  saveCredentials(agentCredentialsPath(agent), credentials(agent, access))
  saveAgentPointer(agentPointerPath(directory), agent, via)
}

const agentToken = (agent: string) => JSON.parse(readFileSync(agentCredentialsPath(agent), "utf8")).access_token

// The V2 plugin context, reduced to the domains this plugin uses. Transforms
// run immediately and again on reload, like the real registries.
function makeContext(directory: string, canonical: string, options: Record<string, unknown> = {}) {
  const servers = new Map<string, any>()
  const tools: any[] = []
  const commands: any[] = []
  const synthetics: any[] = []
  const prompts: any[] = []
  const created: any[] = []
  const transforms: ((editor: any) => void)[] = []
  const queue: any[] = []
  let wake: (() => void) | undefined
  let child = 0
  const editor = {
    list: () => [...servers],
    get: (name: string) => servers.get(name),
    set: (name: string, config: any) => servers.set(name, config),
    update: (name: string, update: (config: any) => void) => update(servers.get(name)),
    remove: (name: string) => servers.delete(name),
  }
  const ctx: any = {
    app: { name: "opencode", version: "2.0.11", channel: "latest" },
    location: { directory, project: { id: "project", directory: canonical, canonical } },
    options: { autoDispatch: "session", ...options },
    storage: { get: async () => undefined, set: async () => {}, remove: async () => {} },
    tool: {
      transform: async (callback: (editor: any) => void) => {
        callback({ list: () => tools, get: (id: string) => tools.find((tool) => tool.name === id), add: (tool: any) => tools.push(tool), update: () => {}, remove: () => {}, namespace: () => {} })
        return { dispose: async () => {} }
      },
    },
    command: {
      transform: async (callback: (editor: any) => void) => {
        callback({ add: (command: any) => commands.push(command) })
        return { dispose: async () => {} }
      },
    },
    mcp: {
      transform: async (callback: (editor: any) => void) => {
        transforms.push(callback)
        callback(editor)
        return { dispose: async () => {} }
      },
      reload: async () => {
        for (const callback of transforms) callback(editor)
      },
    },
    session: {
      get: async ({ sessionID }: { sessionID: string }) => ({ id: sessionID, location: { directory } }),
      remove: async () => {},
      create: async (input: any) => {
        const session = { id: `child-${++child}`, ...input }
        created.push(session)
        return session
      },
      prompt: async (prompt: any) => void prompts.push(prompt),
      synthetic: async (message: any) => void synthetics.push(message),
    },
    event: {
      subscribe: (options?: { signal?: AbortSignal }) => {
        const signal = options?.signal
        return (async function* () {
          while (!signal?.aborted) {
            if (queue.length) {
              yield queue.shift()
              continue
            }
            await new Promise<void>((resolve) => {
              wake = resolve
              signal?.addEventListener("abort", () => resolve(), { once: true })
            })
          }
        })()
      },
    },
  }
  return {
    ctx,
    servers,
    tools,
    commands,
    prompts,
    synthetics,
    created,
    push: (event: any) => {
      queue.push(event)
      wake?.()
      wake = undefined
    },
  }
}

const status = async (harness: ReturnType<typeof makeContext>) => JSON.parse((await harness.tools.find((tool) => tool.name === "office_status").execute({}, {})).content)

after(async () => {
  await office.close()
})

test("a linked worktree nobody connected resolves no credentials of the checkout's", () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  // The older per-repo token file is the checkout's as well, and a worktree
  // used to inherit that too.
  saveCredentials(credentialsPath(r, r, "", "repo"), credentials("legacy", "access-repo"))

  assert.equal(isUnconnectedWorktree(wt), true)
  const resolved = resolveCredentials(wt, r)
  assert.equal(resolved.unconnected, true)
  assert.equal(resolved.credentials, undefined, "no repo-pointer fallback")
  assert.equal(loadCredentials(resolved.credentialsFile), undefined, "and no legacy per-repo token")
  assert.notEqual(resolved.credentialsFile, agentCredentialsPath("agent-a"))
  assert.notEqual(resolved.credentialsFile, credentialsPath(r, r, "", "repo"))
  // Nothing was written for the worktree: it is unconnected, not redirected.
  assert.equal(loadAgentPointer(agentPointerPath(wt)), undefined)

  // The checkout it belongs to keeps its own agent.
  const main = resolveCredentials(r, r)
  assert.equal(main.unconnected, false)
  assert.equal(main.credentials?.access_token, "access-a")
})

test("an unconnected worktree gets no office MCP server, no polling and no status reports", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  office.calls.length = 0
  const harness = makeContext(wt, r, { pollInterval: 1_000 })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // A busy session would report `working` and register the office server
    // with the checkout's token if the worktree had one.
    harness.push({ type: "session.status", data: { sessionID: "wt", status: { type: "busy" } } })
    harness.push({ type: "session.idle", data: { sessionID: "wt" } })
    await new Promise((resolve) => setTimeout(resolve, 300))
    assert.equal(harness.servers.size, 0, "no office MCP server is registered")
    assert.deepEqual(mcpCalls(), [], "no checkin: no poll, no status report, no dispatch")
    const reported = await status(harness)
    assert.equal(reported.officeIdentity, "none (unconnected worktree: run /office here)")
    assert.equal(reported.lastPollAt, undefined)
  } finally {
    await cleanup()
  }
})

test("/office in the worktree connects it as its own agent; the checkout keeps its own", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  office.state.agent = { slug: "opencode", name: "OpenCode" }
  office.calls.length = 0
  const harness = makeContext(wt, r, { pollInterval: 1_000 })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // Before /office the worktree is nobody: nothing is polled or registered.
    harness.push({ type: "session.status", data: { sessionID: "wt", status: { type: "busy" } } })
    await new Promise((resolve) => setTimeout(resolve, 200))
    assert.equal(harness.servers.size, 0)
    assert.deepEqual(mcpCalls(), [])

    // `/office` there: the tool answers with the approval link, Azri approves.
    const started = await harness.tools.find((tool) => tool.name === "office_connect").execute({}, { sessionID: "wt", progress: async () => {} })
    const authorizeUrl = started.content.match(/^(http\S+\/oauth\/authorize\?\S+)$/m)?.[1]
    assert.ok(authorizeUrl, started.content)
    const back = await fetch(authorizeUrl, { redirect: "manual" })
    const landing = await fetch(back.headers.get("location")!)
    assert.match(await landing.text(), /OpenCode is connected to Agent Office/)

    await until(() => harness.servers.get("office") !== undefined, "the office MCP server")
    await until(() => existsSync(agentPointerPath(wt)), "the worktree's own agent pointer")
    assert.equal(loadAgentPointer(agentPointerPath(wt)), "opencode")
    assert.equal(agentPointerVia(agentPointerPath(wt)), "office", "marked: written on purpose by /office here")
    assert.equal(harness.servers.get("office").headers.Authorization, `Bearer ${agentToken("opencode")}`)
    // Connecting the worktree did not rewrite the token the checkout uses.
    assert.equal(agentToken("agent-a"), "access-a")

    // It resolves its own agent now, and polling started without a restart.
    const resolved = resolveCredentials(wt, r)
    assert.equal(resolved.unconnected, false)
    assert.equal(resolved.credentials?.access_token, agentToken("opencode"))
    office.calls.length = 0
    await until(() => office.calls.some((c) => c.body?.params?.name === "checkin"), "the worktree's first poll")
    // The main checkout still resolves to A, not to what the worktree chose.
    assert.equal(resolveCredentials(r, r).credentials?.access_token, "access-a")
  } finally {
    await cleanup()
  }
})

test("a pointer the old plugin left in a worktree grants no identity there", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  // What the old plugin wrote after every poll in a worktree: the identity it
  // had resolved, the checkout's agent, as an unmarked pointer of the
  // worktree's own. It is the shape a real machine still carries.
  saveAgentPointer(agentPointerPath(wt), "agent-a")
  assert.equal(loadAgentPointer(agentPointerPath(wt)), "agent-a", "the file stays; it is ignored, not deleted")

  // Read from the checkout, that pointer is the checkout's agent, as before.
  assert.equal(isUnconnectedWorktree(r), false)
  assert.equal(resolveCredentials(r, r).credentials?.access_token, "access-a")
  // Read from the worktree it counts for nothing.
  assert.equal(isUnconnectedWorktree(wt), true)
  const resolved = resolveCredentials(wt, r)
  assert.equal(resolved.unconnected, true)
  assert.equal(resolved.credentials, undefined)
  assert.equal(loadCredentials(resolved.credentialsFile), undefined)

  office.calls.length = 0
  const harness = makeContext(wt, r, { pollInterval: 1_000 })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "stale", status: { type: "busy" } } })
    await new Promise((resolve) => setTimeout(resolve, 300))
    assert.equal(harness.servers.size, 0, "no office MCP server")
    assert.deepEqual(mcpCalls(), [], "no checkin, no status report, no dispatch")
    // N1: the session is told once why it has no office tools, naming the
    // checkout's agent, which it does not act as.
    assert.match(JSON.stringify(harness.synthetics), /not connected to Agent Office \(the main checkout is, as @agent-a\)/)
    assert.equal(harness.synthetics.length, 1, "said once, not on every event")

    // A pointer written on purpose (what `/office` there writes) is accepted
    // without any cleanup of the old one.
    saveAgentPointer(agentPointerPath(wt), "agent-a", "office")
    assert.equal(isUnconnectedWorktree(wt), false)
    assert.equal(resolveCredentials(wt, r).credentials?.access_token, "access-a")
  } finally {
    if (typeof cleanup === "function") await cleanup()
    office.calls.length = 0
  }
})

test("a worktree with a pointer written on purpose resolves that agent, and only that one", () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  connect(wt, "agent-b", "access-b", "office")
  assert.equal(isUnconnectedWorktree(wt), false)
  assert.equal(resolveCredentials(wt, r).credentials?.access_token, "access-b")
  assert.equal(resolveCredentials(r, r).credentials?.access_token, "access-a")
})

test("a subdirectory of a worktree resolves as the worktree does", () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  mkdirSync(join(wt, "src"))
  // Unconnected, the subdirectory is unconnected too, not the checkout's agent.
  assert.equal(resolveCredentials(join(wt, "src"), r).credentials, undefined)
  // Connected at the worktree's root, its subdirectories find that pointer.
  connect(wt, "agent-b", "access-b", "office")
  assert.equal(resolveCredentials(join(wt, "src"), r).credentials?.access_token, "access-b")
  assert.equal(resolveCredentials(r, r).credentials?.access_token, "access-a")
})

test("a dispatched task worktree resolves as the agent that dispatched it", async () => {
  const r = repo()
  connect(r, "agent-a", "access-a")
  office.state.agent = { slug: "agent-a", name: "Agent A" }
  office.calls.length = 0
  const bin = join(mkdtempSync(join(tmpdir(), "ao-identity-hc-")), "hc")
  const log = `${bin}.log`
  writeFileSync(bin, `#!/bin/sh\nprintf '%s\\n' "$*" >> '${log}'\n[ "$1" = ls ] && echo '{"terminals":[]}'\nexit 0\n`, { mode: 0o755 })
  process.env.AGENT_OFFICE_HC_BIN = bin
  const harness = makeContext(r, r, { pollInterval: 1_000 })
  let cleanup: (() => Promise<void>) | undefined
  try {
    cleanup = await plugin.setup(harness.ctx)
    await until(() => mcpCalls().length > 0, "the baseline poll")
    harness.push({ type: "session.status", data: { sessionID: "parent", status: { type: "busy" } } })
    await until(() => office.calls.some((c) => c.body?.params?.name === "checkin"), "the working report")
    office.state.openTasks = [...office.state.openTasks, { id: "task-identity", title: "Fix it", status: "todo" }]
    await until(() => existsSync(log) && readFileSync(log, "utf8").includes("run --bg"), "the hc launch")

    const taskWorktree = readFileSync(log, "utf8").match(/--cwd (\S+)/)![1]
    // The dispatcher wrote the worktree's own pointer, so its worker resolves
    // the dispatching agent instead of falling back to the checkout's.
    assert.equal(loadAgentPointer(agentPointerPath(taskWorktree)), "agent-a")
    assert.equal(agentPointerVia(agentPointerPath(taskWorktree)), "dispatch", "marked: given on purpose")
    assert.equal(isUnconnectedWorktree(taskWorktree), false)
    assert.equal(resolveCredentials(taskWorktree, r).credentials?.access_token, "access-a")
  } finally {
    delete process.env.AGENT_OFFICE_HC_BIN
    office.state.openTasks = [{ id: "task-open-1", title: "Open task" }]
    await cleanup?.()
    office.calls.length = 0
  }
})

test("an explicit AGENT_OFFICE_TOKEN connects an unconnected worktree and adopts its own agent", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  saveCredentials(credentialsPath(r, r, "", "repo"), credentials("legacy", "access-repo"))
  office.state.agent = { slug: "agent-token", name: "Token Agent" }
  process.env.AGENT_OFFICE_TOKEN = "branch-token"
  const bin = join(mkdtempSync(join(tmpdir(), "ao-identity-hc-")), "hc")
  const log = `${bin}.log`
  writeFileSync(bin, `#!/bin/sh\nprintf '%s\\n' "$*" >> '${log}'\n[ "$1" = ls ] && echo '{"terminals":[]}'\nexit 0\n`, { mode: 0o755 })
  process.env.AGENT_OFFICE_HC_BIN = bin
  try {
    // The token is an identity handed in from outside: the worktree still
    // takes none of the checkout's, not its pointer and not its per-repo token.
    assert.equal(isUnconnectedWorktree(wt), false)
    const resolved = resolveCredentials(wt, r)
    assert.equal(resolved.credentials, undefined)
    assert.notEqual(resolved.credentialsFile, agentCredentialsPath("agent-a"))
    assert.notEqual(resolved.credentialsFile, credentialsPath(r, r, "", "repo"))

    office.calls.length = 0
    const harness = makeContext(wt, r, { pollInterval: 1_000 })
    const cleanup = await plugin.setup(harness.ctx)
    try {
      harness.push({ type: "session.status", data: { sessionID: "wt", status: { type: "busy" } } })
      await until(() => harness.servers.get("office") !== undefined, "the office MCP server")
      assert.equal(harness.servers.get("office").headers.Authorization, "Bearer branch-token")
      await until(() => mcpCalls().some((call) => call.startsWith("checkin ")), "the worktree's poll")
      assert.ok(office.calls.filter((c) => c.path === "/mcp").every((c) => c.auth === "Bearer branch-token"), "the checkout's token is never used")

      // A task dispatched from here is adopted for the agent this session is
      // (the token's), never for the agent the checkout's pointer names.
      office.state.openTasks = [...office.state.openTasks, { id: "task-token", title: "From a branch", status: "todo" }]
      await until(() => existsSync(log) && readFileSync(log, "utf8").includes("run --bg"), "the hc launch")
      const taskWorktree = readFileSync(log, "utf8").match(/--cwd (\S+)/)![1]
      assert.equal(loadAgentPointer(agentPointerPath(taskWorktree)), "agent-token")
      assert.equal(resolveCredentials(taskWorktree, r).credentials, undefined, "and no token file of the checkout's either")
      // The token is an identity handed in, not one this worktree keeps: a
      // poll with it must not leave a pointer that outlives it.
      assert.equal(existsSync(agentPointerPath(wt)), false, "the token left no pointer behind in the worktree")
    } finally {
      await cleanup()
    }
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    delete process.env.AGENT_OFFICE_HC_BIN
    office.state.openTasks = [{ id: "task-open-1", title: "Open task" }]
    office.calls.length = 0
  }
  // With the token gone the worktree is unconnected again: nothing is left.
  assert.equal(isUnconnectedWorktree(wt), true)
  assert.equal(resolveCredentials(wt, r).credentials, undefined)
})

test("a worktree /office that finds a token marks nothing and copies no token file", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  // A fresh slug, so no other test can have written its token file.
  office.state.agent = { slug: "agent-handed-in", name: "Handed In" }
  office.calls.length = 0
  // The token is set and removed inside this outer try, so a failure anywhere
  // cannot leave it set for the tests after this one.
  try {
    process.env.AGENT_OFFICE_TOKEN = "branch-token"
    const harness = makeContext(wt, r, { pollInterval: 600_000 })
    const cleanup = await plugin.setup(harness.ctx)
    try {
      // The bearer is already there, so `/office` asks no approval. Reaching
      // the token's agent is not a person connecting this worktree: nothing is
      // marked.
      const officeConnect = harness.tools.find((tool) => tool.name === "office_connect")
      const summary = String((await officeConnect.execute({}, { sessionID: "wt" })).content)
      assert.match(summary, /Connected to Agent Office as @agent-handed-in/)
      assert.equal(existsSync(agentPointerPath(wt)), false, "the token's agent is not written to the worktree")
      assert.equal(existsSync(agentCredentialsPath("agent-handed-in")), false, "no token file copied under the token's agent")
    } finally {
      if (typeof cleanup === "function") await cleanup()
    }
    // With the token gone, the identity it carried is gone with it.
    delete process.env.AGENT_OFFICE_TOKEN
    assert.equal(isUnconnectedWorktree(wt), true)
    assert.equal(resolveCredentials(wt, r).credentials, undefined)

    // The same call where the worktree is connected on purpose, as A, and the
    // token names another agent: A's OAuth must not be copied under that name.
    connect(wt, "agent-a", "access-a", "office")
    process.env.AGENT_OFFICE_TOKEN = "branch-token"
    const second = makeContext(wt, r, { pollInterval: 600_000 })
    const cleanupSecond = await plugin.setup(second.ctx)
    try {
      await second.tools.find((tool) => tool.name === "office_connect").execute({}, { sessionID: "wt" })
      assert.equal(loadAgentPointer(agentPointerPath(wt)), "agent-a", "the worktree's own pointer stays its own")
      assert.equal(agentPointerVia(agentPointerPath(wt)), "office")
      assert.equal(existsSync(agentCredentialsPath("agent-handed-in")), false, "A's OAuth is not saved under the token's agent")
    } finally {
      if (typeof cleanupSecond === "function") await cleanupSecond()
    }
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    office.calls.length = 0
  }
})

test("the main checkout, its subdirectories and plain directories are unchanged", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  mkdirSync(join(r, "src"))
  const plain = join(mkdtempSync(join(tmpdir(), "ao-identity-plain-")), "notes")
  mkdirSync(plain)
  saveCredentials(credentialsPath(plain), credentials("plain", "access-plain"))
  office.state.agent = { slug: "agent-a", name: "Agent A" }
  office.calls.length = 0
  const harness = makeContext(r, r, { pollInterval: 1_000 })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // The checkout polls and registers its own agent as it always has.
    harness.push({ type: "session.status", data: { sessionID: "main", status: { type: "busy" } } })
    await until(() => mcpCalls().some((call) => call.startsWith("checkin ")), "the checkout's poll")
    await until(() => harness.servers.get("office") !== undefined, "the office MCP server")
    assert.equal(harness.servers.get("office").headers.Authorization, "Bearer access-a")

    // A subdirectory of the checkout is not a worktree: the same agent.
    assert.equal(isUnconnectedWorktree(join(r, "src")), false)
    assert.equal(resolveCredentials(join(r, "src"), r).credentials?.access_token, "access-a")
    // A plain directory keeps its legacy per-repo token.
    assert.equal(isUnconnectedWorktree(plain), false)
    assert.equal(resolveCredentials(plain).credentials?.access_token, "access-plain")
    // The worktree beside them is the only one that lost the fallback.
    assert.equal(isUnconnectedWorktree(wt), true)
    // A marked pointer at the checkout still names its agent: nothing writes
    // one there, but both shapes read the same outside a worktree.
    saveAgentPointer(agentPointerPath(r), "agent-a", "office")
    assert.equal(isUnconnectedWorktree(r), false)
    assert.equal(resolveCredentials(r, r).credentials?.access_token, "access-a")
  } finally {
    await cleanup()
    office.calls.length = 0
  }
})

// A poller run, async so this process keeps answering as the fake Office: a
// spawnSync would block the event loop and the child's checkin would time out.
function poll(dir: string, bin: string): Promise<{ code: number | null; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", pollScript, "--dir", dir], {
      env: { ...process.env, AGENT_OFFICE_OPENCODE_BIN: bin },
      stdio: ["ignore", "ignore", "pipe"],
    })
    let stderr = ""
    child.stderr.on("data", (chunk) => (stderr += String(chunk)))
    child.on("error", reject)
    child.on("close", (code) => resolve({ code, stderr }))
  })
}

test("the cron poller in an unconnected worktree polls nothing and says why", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  office.state.agent = { slug: "agent-a", name: "Agent A" }
  const bin = join(mkdtempSync(join(tmpdir(), "ao-identity-opencode-")), "opencode")
  writeFileSync(bin, "#!/bin/sh\nexit 0\n", { mode: 0o755 })
  office.calls.length = 0
  const polled = await poll(wt, bin)
  assert.equal(polled.code, 2, polled.stderr)
  assert.match(polled.stderr, /a linked worktree nobody connected; run \/office in it/)
  assert.deepEqual(office.calls, [], "no checkin: nothing polled, claimed or dispatched")
  // The checkout it belongs to still polls as its own agent.
  office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-checkout", from: "azri", subject: "For the checkout" }]
  const main = await poll(r, bin)
  assert.equal(main.code, 0, main.stderr)
  assert.ok(office.calls.some((c) => c.path === "/mcp" && c.auth === "Bearer access-a"), "the checkout's poller checks in as A")
  office.state.unreadMessages = office.state.unreadMessages.filter((m) => m.id !== "message-checkout")
  office.calls.length = 0
})
