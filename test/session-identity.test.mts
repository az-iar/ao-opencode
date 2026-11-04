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
// agent-office#560: session approvals are host-wide, not under the runtime.
const sessionPointer = (sessionID: string) => join(home, ".config", "agent-office", "sessions", `${workspaceKey(sessionID)}.json`)
const approveSession = (sessionID: string, agent: string) => {
  mkdirSync(dirname(sessionPointer(sessionID)), { recursive: true })
  writeFileSync(sessionPointer(sessionID), JSON.stringify({ agent }, null, 2) + "\n", { mode: 0o600 })
}
const sessionAgent = (sessionID: string) => loadAgentPointer(sessionPointer(sessionID))

// Every test that connects gives its folder a token of its own, and every
// assertion counts only the calls that carry the tokens this test installed.
// The per-test token is also the generation every call is stamped with: the fake
// keeps each call with the Authorization it arrived with, so a late call is
// attributed to the test that made it rather than dropped.
//
// The flake this replaces: the previous test's plugin can still have an Office
// request in flight when its `cleanup()` resolves - the approval's summary
// checkin, or the first poll that approval schedules - and on a slow runner that
// request lands after the next test cleared the log, where it was counted
// against a child that sent nothing (az-iar/ao-opencode#35 and #36: run
// 41794077476's `not ok 108` at 267a006, run 41792874366's test 91 at d0c0ce7,
// whose attempt 1 failed and attempt 2 passed on the same head, and once on main
// during #28's review). A straggler speaks with the previous test's token, so it
// cannot be counted here; a real child leak speaks with this test's folder token
// or with no token at all, so it still is.
let tokens = 0
const testToken = (what: string) => `access-${what}-${++tokens}`
// The calls this test made: what the fake holds under this test's tokens.
const authored = (mine: string[]) => mine.flatMap((token) => office.byToken(token))
// What a test that expects nothing must not see: a call carrying this test's
// tokens, or a call carrying no token at all. A call with no Authorization
// header cannot be attributed to a test, so it counts against the test that was
// running when it arrived - the leak this test exists to catch is a session
// speaking without an approval of its own, and a filter on the token alone would
// miss exactly that call.
const strays = (mine: string[]) => office.calls.filter((c) => !c.auth || mine.some((t) => c.auth === `Bearer ${t}`))
// Every call in this test's window with its generation, so a failure says which
// calls are this test's and which arrived from another test - or from none.
const windowReport = (mine: string[]) =>
  office.calls
    .map((c) => {
      const token = c.auth?.replace(/^Bearer /, "")
      const whose = !c.auth ? "no token" : mine.includes(token!) ? "this test's" : "another test's straggler"
      return `${c.method} ${c.path}${c.body?.params?.name ? ` ${c.body.params.name}` : ""} [${whose}: ${token ?? "none"}]`
    })
    .join("; ") || "no calls"
// The notices the plugin gave the session, and whether one matches: the fact
// that a refusing session has had its turn, in place of waiting long enough to
// hope it has.
const notices = (harness: { synthetics: { text: string }[] }) => harness.synthetics.map((m) => m.text).join(" | ")
const noticed = (harness: { synthetics: { text: string }[] }, pattern: RegExp) => harness.synthetics.some((m) => pattern.test(m.text))
// A /mcp call that arrived with no token: the Office refused it, and no test's
// token filter can claim it.
const unauthenticatedMcp = () => office.calls.filter((c) => !c.auth && c.path === "/mcp")
const mcpCalls = (mine: string[]) => authored(mine).filter((c) => c.path === "/mcp")
const checkins = (mine: string[]) => mcpCalls(mine).filter((c) => c.body?.params?.name === "checkin")
// Ends a test by waiting for the Office to be answering nothing before the log
// is cleared: the fake's in-flight count is the fact, never a quiet period. The
// log is reset per test so it stays bounded, and a call that arrives after this
// drain is still recorded - under the finished test's token - so it is
// attributed to that test rather than vanishing.
const drained = async () => {
  await office.idle()
  office.calls.length = 0
}

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
  const folder = testToken("host-folder")
  const approved = testToken("host")
  connect(dir, "host", folder)
  // The approval issues this test's token too, so the calls the approving
  // plugin makes are this test's and no other test's.
  office.state.issuedToken = approved
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // Nothing is sent before a session owns an identity, even with a folder
    // pointer and its token present. The fact to wait on is that the plugin had
    // a session in hand and declined to speak for it: the notice it gives a
    // session with no approval of its own.
    harness.push({ type: "session.status", data: { sessionID: "unapproved-session", status: { type: "busy" } } })
    await until(() => noticed(harness, /Run \/office/), "the unapproved session's notice")
    assert.equal(harness.servers.size, 0, "the Office MCP server was registered before any session approved")
    assert.deepEqual(strays([folder, approved]), [], `an Office request was made before any session approved: ${windowReport([folder, approved])}`)

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
    await until(() => checkins([approved]).length > 0, "the approving session's first checkin")

    const token = agentToken("host")
    assert.equal(token, approved, "the approval did not issue this test's token, so its calls cannot be attributed")
    const server = harness.servers.get("office")
    assert.equal(server.headers.Authorization, `Bearer ${token}`)
    // Every Office call this test made carried the approval's token: nothing
    // spoke with the folder's token instead, and nothing went out without one.
    assert.ok(mcpCalls([folder, approved]).every((c) => c.auth === `Bearer ${token}`), "an Office call did not carry the session's approval")
    assert.deepEqual(unauthenticatedMcp(), [], "an Office call went out with no token")
    // The approval is now this session's own, not the folder's.
    assert.equal(sessionAgent("host-session"), "host", "the approval is not recorded against the session that made it")
    assert.ok(harness.synthetics.some((m) => m.sessionID === "host-session" && /^Connected to Agent Office as @host/.test(m.text)))
  } finally {
    office.state.issuedToken = undefined
    await cleanup()
    await drained()
  }
})

test("a child session with the same directory, environment and pointer files but no approval of its own registers nothing and sends nothing", async () => {
  const dir = repo()
  const host = testToken("host")
  // The host approved in this folder. Its token and its folder pointer are
  // exactly what a child inherits - and what it must not be able to speak with.
  connect(dir, "host", host)
  approveSession("host-session", "host")
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // A different session id in the same folder: it has no approval of its own.
    harness.push({ type: "session.status", data: { sessionID: "child-session", status: { type: "busy" } } })
    harness.push({ type: "session.idle", data: { sessionID: "child-session" } })
    // The child's turn is over when it has been told it cannot post: the refusal
    // is what the plugin emits after deciding not to register, so the wait is on
    // that fact, not on a duration.
    await until(() => noticed(harness, /no Office identity of its own/), "the child's refusal notice")

    // The assertion that matters on the parent commit: nothing was served in
    // reply to this folder, neither with its token nor without one. A child
    // that leaked would speak with the token it inherited, so the filter cannot
    // hide it; the previous test's straggler carries that test's token and is
    // not counted here.
    assert.deepEqual(strays([host]), [], `a child with no approval of its own still called the Office: ${windowReport([host])}`)
    assert.equal(harness.servers.size, 0, "the Office MCP server was registered for a child with no approval of its own")
    const notice = notices(harness)
    assert.match(notice, /no Office identity of its own/, `the child was not told why it cannot post: ${notice}`)
    // The refusal has to be traceable: it names the session and the agent it
    // would have spoken as, because a sender line is unverifiable after the fact.
    assert.match(notice, /child-session/, `the notice does not name the session: ${notice}`)
    assert.match(notice, /\bhost\b/, `the notice does not name the agent it would have spoken as: ${notice}`)
    assert.match(notice, /\/office/, `the notice does not say how to get an identity of its own: ${notice}`)
  } finally {
    await cleanup()
    await drained()
  }
})

test("a child launched with AGENT_OFFICE_CHILD=1 is refused too", async () => {
  const dir = repo()
  const host = testToken("host")
  connect(dir, "host", host)
  office.state.agent = { slug: "host", name: "Host" }
  process.env.AGENT_OFFICE_CHILD = "1"
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "flagged-child", status: { type: "busy" } } })
    await until(() => noticed(harness, /Child session flagged-child/), "the flagged child's refusal notice")
    assert.deepEqual(strays([host]), [], `a flagged child still called the Office: ${windowReport([host])}`)
    assert.equal(harness.servers.size, 0)
    const notice = notices(harness)
    assert.match(notice, /Child session flagged-child/, `the refusal is not the child form: ${notice}`)
    assert.match(notice, /\bhost\b/)
  } finally {
    delete process.env.AGENT_OFFICE_CHILD
    await cleanup()
    await drained()
  }
})

test("AGENT_OFFICE_TOKEN is the process's own identity for a non-child, with no session approval", async () => {
  const dir = repo()
  const folder = testToken("folder-agent")
  const runner = testToken("runner")
  // A folder pointer naming another agent: the token wins, and nothing is
  // recorded against a session, because a grant is not an approval.
  connect(dir, "folder-agent", folder)
  office.state.agent = { slug: "runner", name: "Runner" }
  process.env.AGENT_OFFICE_TOKEN = runner
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // The token is the process's identity from boot: the first poll checks in
    // with it, before any session has spoken.
    await until(() => checkins([runner]).length > 0, "the token's boot checkin")
    assert.ok(mcpCalls([folder, runner]).every((c) => c.auth === `Bearer ${runner}`), "the folder's token was used instead")
    assert.deepEqual(unauthenticatedMcp(), [], "an Office call went out with no token")
    harness.push({ type: "session.status", data: { sessionID: "runner-session", status: { type: "idle" } } })
    await until(() => harness.servers.get("office") !== undefined, "the token's registration")
    assert.equal(harness.servers.get("office").headers.Authorization, `Bearer ${runner}`)
    // A grant is not an approval, so nothing is recorded against the session
    // that happened to be speaking.
    await tick()
    assert.equal(sessionAgent("runner-session"), undefined, "a token recorded a session approval")
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    await drained()
  }
})

test("a child with AGENT_OFFICE_TOKEN and no approval of its own is still refused", async () => {
  const dir = repo()
  const host = testToken("host")
  const runner = testToken("runner")
  connect(dir, "host", host)
  process.env.AGENT_OFFICE_TOKEN = runner
  process.env.AGENT_OFFICE_CHILD = "1"
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "token-child", status: { type: "busy" } } })
    await until(() => noticed(harness, /Child session token-child/), "the token child's refusal notice")
    assert.deepEqual(strays([host, runner]), [], `a child with an inherited token called the Office: ${windowReport([host, runner])}`)
    assert.equal(harness.servers.size, 0, "a child registered on an inherited token")
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
    delete process.env.AGENT_OFFICE_CHILD
    await cleanup()
    await drained()
  }
})

test("a dispatched worker keeps the identity its dispatcher wrote into its worktree", async () => {
  const r = repo()
  const wt = worktree(r)
  const access = testToken("agent-a")
  connect(r, "agent-a", access)
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
    assert.equal(harness.servers.get("office").headers.Authorization, `Bearer ${access}`)
    assert.equal(harness.servers.get("office").headers["X-Office-Session-Agent"], "agent-a", "the worker does not name the agent it was given")
    await until(() => checkins([access]).length > 0, "the worker's checkin")
    assert.equal(sessionAgent("worker-session"), undefined, "the grant was recorded as a session approval")
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    await drained()
  }
})

test("an approval made before this change keeps working when OpenCode still remembers the session it belonged to", async () => {
  const dir = repo()
  // The shape this upgrade meets: a folder approval, no session pointer (the
  // old plugin kept none), and OpenCode's own record of the last session that
  // worked here as that agent.
  const host = testToken("host")
  connect(dir, "host", host)
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir, {}, { [`last-session:host:${workspaceKey(dir)}`]: "host-session" })
  const cleanup = await plugin.setup(harness.ctx)
  try {
    // Still nothing until that session shows up - the record names one id. The
    // fact to wait on is that the plugin had a session in hand and declined to
    // speak for it: a stranger in this folder is refused by name, and the
    // recorded session has not been handed anything yet.
    harness.push({ type: "session.status", data: { sessionID: "stranger-session", status: { type: "busy" } } })
    await until(() => noticed(harness, /stranger-session has no Office identity/), "the stranger's refusal notice")
    assert.equal(harness.servers.size, 0)
    assert.deepEqual(strays([host]), [], `an Office request was made before the recorded session spoke: ${windowReport([host])}`)
    harness.push({ type: "session.status", data: { sessionID: "host-session", status: { type: "busy" } } })
    await until(() => harness.servers.get("office")?.headers["X-Office-Session-Agent"] === "host", "the pre-change approval's registration")
    // Recorded against the session from now on, so the next start needs
    // neither the storage record nor an approval.
    assert.equal(sessionAgent("host-session"), "host")
    await until(() => checkins([host]).length > 0, "the pre-change approval's checkin")
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    await drained()
  }
})

test("a folder approval with no session record is pre-fill only: the session is asked to run /office", async () => {
  const dir = repo()
  const host = testToken("host")
  connect(dir, "host", host)
  office.state.agent = { slug: "host", name: "Host" }
  office.calls.length = 0
  const harness = makeContext(dir)
  const cleanup = await plugin.setup(harness.ctx)
  try {
    harness.push({ type: "session.status", data: { sessionID: "unknown-session", status: { type: "busy" } } })
    await until(() => noticed(harness, /Run \/office/), "the session's approval notice")
    assert.deepEqual(strays([host]), [], `a session with no approval sent an Office request: ${windowReport([host])}`)
    assert.equal(harness.servers.size, 0)
    const notice = notices(harness)
    assert.match(notice, /unknown-session/)
    assert.match(notice, /\bhost\b/)
    assert.match(notice, /Run \/office/, `the session is not told to approve: ${notice}`)
  } finally {
    office.state.agent = { slug: "opencode", name: "OpenCode" }
    await cleanup()
    await drained()
  }
})
