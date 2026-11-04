// GH-491 (port of ao-pi#3, GH-476): an Office extension may speak only as an
// agent whose approval is recorded against the session it is speaking for.
// OpenCode keeps one agent pointer per folder, so a child session with a new
// id - the same directory, environment and pointer files, no approval of its
// own - used to fall through to that pointer and post as the folder's agent.
// Here identity is per session: a session pointer written on approval. The
// folder pointer only pre-fills an approval.
//
// The session pointer is written directly, in the same shape the plugin
// writes, so this file also runs against the parent commit and can be seen
// failing there (the child instance checks in from the folder pointer).
import assert from "node:assert/strict"
import { execFileSync } from "node:child_process"
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"
import { after, test } from "node:test"
import { fakeOffice } from "./fake-office.mts"

const office = fakeOffice()
const tick = () => new Promise((r) => setTimeout(r, 20))
// A boot poll, a status report and a registration all settle well inside this;
// the tests that assert "nothing was sent" need it to be generous enough that
// the parent commit's eager poll has had its chance (the red run).
const settle = (ms = 300) => new Promise((r) => setTimeout(r, ms))
const until = async (check: () => boolean, what: string) => {
  for (let i = 0; i < 300; i++) {
    if (check()) return
    await tick()
  }
  assert.fail(`timed out waiting for ${what}`)
}

// Credentials and pointers live under $HOME, so this suite gets a home of its
// own; $HOME must be set before lib/office.ts loads.
const home = mkdtempSync(join(tmpdir(), "ao-session-identity-"))
process.env.HOME = home
process.env.AO_HOME = join(home, "ao-home")
process.env.AGENT_OFFICE_CALLBACK_PORT = "45521"
process.env.AGENT_OFFICE_URL = await office.listen()
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_OPENCODE_CREDENTIALS
delete process.env.AGENT_OFFICE_CHILD
delete process.env.AGENT_OFFICE_CHILD_NAME
// Never open a real browser from a test: the approval in the first case reaches
// the opener, and a real browser's own request for the approval URL lands in
// the next case's Office call log (az-iar/ao-pi#10, #13).
process.env.AGENT_OFFICE_OPEN_BROWSER = "off"
// Inside Herdr a linked worktree does not poll (only the main checkout does),
// so the worktree case below needs the ambient flag gone.
delete process.env.HERDR_ENV

const lib = await import("../lib/office.ts")
const { adoptWorktree, agentCredentialsPath, agentPointerPath, loadAgentPointer, saveAgentPointer, saveCredentials, workspaceKey } = lib
const { default: plugin } = await import("../index.ts")

// The per-session approval record, at the path the plugin writes it to.
const sessionPointer = (sessionID: string) => join(home, ".config", "agent-office", "opencode", "sessions", `${workspaceKey(sessionID)}.json`)
const approveSession = (sessionID: string, agent: string) => {
  mkdirSync(dirname(sessionPointer(sessionID)), { recursive: true })
  writeFileSync(sessionPointer(sessionID), JSON.stringify({ agent }, null, 2) + "\n", { mode: 0o600 })
}
const sessionAgent = (sessionID: string) => loadAgentPointer(sessionPointer(sessionID))

const mcpCalls = () => office.calls.filter((c) => c.path === "/mcp")
const checkins = () => mcpCalls().filter((c) => c.body?.params?.name === "checkin")

function repo() {
  const path = join(mkdtempSync(join(tmpdir(), "ao-session-identity-repo-")), "repo")
  mkdirSync(path)
  execFileSync("git", ["init", "-q", path])
  execFileSync("git", ["-C", path, "-c", "user.name=t", "-c", "user.email=t@t", "commit", "-q", "--allow-empty", "-m", "init"])
  return path
}

function worktree(r: string, branch = "task"): string {
  const path = `${r}-${branch}`
  execFileSync("git", ["-C", r, "worktree", "add", "-q", "-b", branch, path])
  return path
}

const credentials = (agent: string, access: string) => ({
  office: process.env.AGENT_OFFICE_URL!,
  client_id: `client-${agent}`,
  access_token: access,
  refresh_token: `refresh-${agent}`,
  expires_at: Date.now() + 3_600_000,
})

// What an approval leaves behind for a folder: the agent's token file and the
// folder's pointer to it, which every session in the folder then inherits.
function connect(directory: string, agent: string, access: string) {
  saveCredentials(agentCredentialsPath(agent), credentials(agent, access))
  saveAgentPointer(agentPointerPath(directory), agent)
}

const agentToken = (agent: string) => JSON.parse(readFileSync(agentCredentialsPath(agent), "utf8")).access_token

// The V2 plugin context, reduced to the domains this plugin uses. Transforms
// run immediately and again on reload, like the real registries. `stored`
// seeds the plugin's durable storage, which is where a pre-change approval's
// last session is remembered.
function makeContext(directory: string, options: Record<string, unknown> = {}, stored: Record<string, string> = {}) {
  const servers = new Map<string, any>()
  const tools: any[] = []
  const commands: any[] = []
  const synthetics: any[] = []
  const prompts: any[] = []
  const created: any[] = []
  const storage = new Map<string, unknown>(Object.entries(stored))
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
    location: { directory, project: { id: "project", directory, canonical: directory } },
    options: { autoDispatch: "off", pollInterval: 600_000, ...options },
    storage: {
      get: async (key: string) => storage.get(key),
      set: async (key: string, value: unknown) => void storage.set(key, value),
      remove: async (key: string) => void storage.delete(key),
    },
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
    synthetics,
    prompts,
    created,
    push: (event: any) => {
      queue.push(event)
      wake?.()
      wake = undefined
    },
  }
}

after(async () => {
  await office.close()
})

test("a session that approves here owns an identity: it registers and posts, and the approval is recorded against its id", async () => {
  const dir = repo()
  connect(dir, "host", "access-host")
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // Nothing is sent before a session owns an identity, even with a folder
    // pointer and its token present.
    await settle()
    assert.equal(harness.servers.size, 0, "the Office MCP server was registered before any session approved")
    assert.deepEqual(office.calls, [], "an Office request was made before any session approved")

    // The tool a session without an identity calls offers an approval, not
    // "already connected" (ao-pi's inherited-pointer case).
    const officeConnect = harness.tools.find((tool) => tool.name === "office_connect")
    const started = await officeConnect.execute({}, { sessionID: "host-session" })
    const authorizeUrl = String(started.content).match(/^(http\S+\/oauth\/authorize\?\S+)$/m)?.[1]
    assert.ok(authorizeUrl, String(started.content))

    // Azri approves in the browser: the fake Office sends the code back to the
    // plugin's loopback listener, which finishes the approval.
    await fetch(authorizeUrl, { redirect: "follow" })
    // The first registration is taken before the checkin names the agent; the
    // reload after adoption is the one that carries it (GH-476 on #24).
    await until(() => harness.servers.get("office")?.headers["X-Office-Session-Agent"] === "host", "the office MCP server naming the session's agent")
    await until(() => checkins().length > 0, "the approving session's first checkin")

    const token = agentToken("host")
    const server = harness.servers.get("office")
    assert.equal(server.headers.Authorization, `Bearer ${token}`)
    assert.ok(mcpCalls().every((c) => c.auth === `Bearer ${token}`))
    // The approval is now this session's own, not the folder's.
    assert.equal(sessionAgent("host-session"), "host", "the approval is not recorded against the session that made it")
    assert.ok(harness.synthetics.some((m) => m.sessionID === "host-session" && /^Connected to Agent Office as @host/.test(m.text)))
  } finally {
    await cleanup()
    office.calls.length = 0
  }
})

test("a child session with the same directory, environment and pointer files but no approval of its own registers nothing and sends nothing", async () => {
  const dir = repo()
  // The host approved in this folder. Its token and its folder pointer are
  // exactly what a child inherits - and what it must not be able to speak with.
  connect(dir, "host", "access-host")
  approveSession("host-session", "host")
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // A different session id in the same folder: it has no approval of its own.
    harness.push({ type: "session.status", data: { sessionID: "child-session", status: { type: "busy" } } })
    harness.push({ type: "session.idle", data: { sessionID: "child-session" } })
    // Long enough for a boot poll, a status report and a registration to land.
    await settle()

    // The assertion that matters on the parent commit: the served-request list
    // is empty, not merely "no error was thrown".
    assert.deepEqual(office.calls.map((c) => c.path), [], "a child with no approval of its own still called the Office")
    assert.equal(harness.servers.size, 0, "the Office MCP server was registered for a child with no approval of its own")
    const notice = harness.synthetics.map((m: any) => m.text).join(" | ")
    assert.match(notice, /no Office identity of its own/, `the child was not told why it cannot post: ${notice}`)
    // The refusal has to be traceable: it names the session and the agent it
    // would have spoken as, because a sender line is unverifiable after the fact.
    assert.match(notice, /child-session/, `the notice does not name the session: ${notice}`)
    assert.match(notice, /\bhost\b/, `the notice does not name the agent it would have spoken as: ${notice}`)
    assert.match(notice, /\/office/, `the notice does not say how to get an identity of its own: ${notice}`)
  } finally {
    await cleanup()
    office.calls.length = 0
  }
})

test("a child launched with AGENT_OFFICE_CHILD=1 is refused too", async () => {
  const dir = repo()
  connect(dir, "host", "access-host")
  office.state.agent = { slug: "host", name: "Host" }
  process.env.AGENT_OFFICE_CHILD = "1"
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "flagged-child", status: { type: "busy" } } })
    await settle()
    assert.deepEqual(office.calls.map((c) => c.path), [], "a flagged child still called the Office")
    assert.equal(harness.servers.size, 0)
    const notice = harness.synthetics.map((m: any) => m.text).join(" | ")
    assert.match(notice, /Child session flagged-child/, `the refusal is not the child form: ${notice}`)
    assert.match(notice, /\bhost\b/)
  } finally {
    delete process.env.AGENT_OFFICE_CHILD
    await cleanup()
    office.calls.length = 0
  }
})

test("AGENT_OFFICE_TOKEN is the process's own identity for a non-child, with no session approval", async () => {
  const dir = repo()
  // A folder pointer naming another agent: the token wins, and nothing is
  // recorded against a session, because a grant is not an approval.
  connect(dir, "folder-agent", "access-folder")
  office.state.agent = { slug: "runner", name: "Runner" }
  process.env.AGENT_OFFICE_TOKEN = "runner-token"
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // The token is the process's identity from boot: the first poll checks in
    // with it, before any session has spoken.
    await until(() => checkins().length > 0, "the token's boot checkin")
    assert.ok(mcpCalls().every((c) => c.auth === "Bearer runner-token"), "the folder's token was used instead")
    harness.push({ type: "session.status", data: { sessionID: "runner-session", status: { type: "idle" } } })
    await until(() => harness.servers.get("office") !== undefined, "the token's registration")
    assert.equal(harness.servers.get("office").headers.Authorization, "Bearer runner-token")
    // A grant is not an approval, so nothing is recorded against the session
    // that happened to be speaking.
    await tick()
    assert.equal(sessionAgent("runner-session"), undefined, "a token recorded a session approval")
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    office.calls.length = 0
  }
})

test("a child with AGENT_OFFICE_TOKEN and no approval of its own is still refused", async () => {
  const dir = repo()
  connect(dir, "host", "access-host")
  process.env.AGENT_OFFICE_TOKEN = "runner-token"
  process.env.AGENT_OFFICE_CHILD = "1"
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "token-child", status: { type: "busy" } } })
    await settle()
    assert.deepEqual(office.calls.map((c) => c.path), [], "a child with an inherited token called the Office")
    assert.equal(harness.servers.size, 0, "a child registered on an inherited token")
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    delete process.env.AGENT_OFFICE_CHILD
    await cleanup()
    office.calls.length = 0
  }
})

test("a dispatched worker keeps the identity its dispatcher wrote into its worktree", async () => {
  const r = repo()
  const wt = worktree(r)
  connect(r, "agent-a", "access-a")
  // What dispatchHc does before the worker starts: the worktree's own pointer,
  // marked as a dispatch. That mark is the grant; no session approves.
  adoptWorktree(wt, "agent-a")
  office.state.agent = { slug: "agent-a", name: "Agent A" }
  office.calls.length = 0
  const harness = makeContext(wt)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // The worker's own status events; the grant is what it acts with, not a
    // session approval.
    harness.push({ type: "session.status", data: { sessionID: "worker-session", status: { type: "busy" } } })
    await until(() => harness.servers.get("office") !== undefined, "the worker's registration")
    assert.equal(harness.servers.get("office").headers.Authorization, "Bearer access-a")
    assert.equal(harness.servers.get("office").headers["X-Office-Session-Agent"], "agent-a", "the worker does not name the agent it was given")
    await until(() => checkins().length > 0, "the worker's checkin")
    assert.equal(sessionAgent("worker-session"), undefined, "the grant was recorded as a session approval")
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    office.calls.length = 0
  }
})

test("an approval made before this change keeps working when OpenCode still remembers the session it belonged to", async () => {
  const dir = repo()
  // The shape this upgrade meets: a folder approval, no session pointer (the
  // old plugin kept none), and OpenCode's own record of the last session that
  // worked here as that agent.
  connect(dir, "host", "access-host")
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir, {}, { [`last-session:host:${workspaceKey(dir)}`]: "host-session" })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // Still nothing until that session shows up - the record names one id.
    await settle()
    assert.equal(harness.servers.size, 0)
    assert.deepEqual(office.calls, [])
    harness.push({ type: "session.status", data: { sessionID: "host-session", status: { type: "busy" } } })
    await until(() => harness.servers.get("office")?.headers["X-Office-Session-Agent"] === "host", "the pre-change approval's registration")
    // Recorded against the session from now on, so the next start needs
    // neither the storage record nor an approval.
    assert.equal(sessionAgent("host-session"), "host")
    await until(() => checkins().length > 0, "the pre-change approval's checkin")
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    office.calls.length = 0
  }
})

test("a folder approval with no session record is pre-fill only: the session is asked to run /office", async () => {
  const dir = repo()
  connect(dir, "host", "access-host")
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "unknown-session", status: { type: "busy" } } })
    await settle()
    assert.deepEqual(office.calls.map((c) => c.path), [])
    assert.equal(harness.servers.size, 0)
    const notice = harness.synthetics.map((m: any) => m.text).join(" | ")
    assert.match(notice, /unknown-session/)
    assert.match(notice, /\bhost\b/)
    assert.match(notice, /Run \/office/, `the session is not told to approve: ${notice}`)
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    office.calls.length = 0
  }
})
