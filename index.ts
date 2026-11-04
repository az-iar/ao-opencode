// Agent Office plugin for OpenCode (https://opencode.ai/docs/plugins/).
//
//   /office [name]     connects: OAuth in the browser, Azri approves, the token
//                      is stored under the agent name, and the Office MCP tools
//                      appear in OpenCode. The name pre-fills a new agent.
//   while running      `working` while any session runs, a heartbeat every
//                      90 s, `idle` once every session is quiet; `blocked` is
//                      never overridden. New mail runs in an inbox child of
//                      the chat's session. A new assigned task is claimed in a
//                      worktree and queued in this session, or in one Herdr
//                      agent, up to two tasks at a time. No supervisor/worker
//                      subagent pair.
//
//   AGENT_OFFICE_URL                   dashboard origin, default https://office.byma.tech
//   AGENT_OFFICE_TOKEN                 an agent API token instead of OAuth (optional)
//   AGENT_OFFICE_OPENCODE_CREDENTIALS  exact credentials file, overrides agent-scoped storage
//   AGENT_OFFICE_CALLBACK_PORT         loopback port for the OAuth callback, default 41997
//   AGENT_OFFICE_POLL_MS               inbox poll interval, default 60000
//   AGENT_OFFICE_AUTODISPATCH          on (default), off, herdr, or session
//   AGENT_OFFICE_CLIENT_NAME           OAuth client name; defaults to folder@host
//
// Identity is per session (GH-491): an approval is recorded against the
// session that made it (`~/.config/agent-office/opencode/sessions/<id>.json`),
// and a session with no approval of its own registers no office MCP server and
// makes no request rather than posting as the folder's agent. The folder
// pointer only pre-fills an approval; AGENT_OFFICE_TOKEN and the credentials
// file are grants handed to the process on purpose.
//
// Tokens are never logged.

import { Plugin } from "@opencode/plugin"
import { spawn } from "node:child_process"
import { createHash, randomBytes } from "node:crypto"
import { mkdirSync, existsSync } from "node:fs"
import { createServer } from "node:http"
import { hostname } from "node:os"
import { basename } from "node:path"
import { fileURLToPath } from "node:url"
import {
  type Credentials,
  acquireDispatch,
  addWorktree,
  adoptWorktree,
  hcBin,
  holdDispatch,
  agentCredentialsPath,
  agentPointerVia,
  agentStateDir,
  loadAgentPointer,
  loadCredentials,
  blockTask,
  CREDENTIALS_OVERRIDE,
  claimNotice,
  claimTask,
  dispatchRefusal,
  createBearer,
  createClient,
  cronLine,
  DISPATCH_STALE_MS,
  fromOthers,
  inboxNotice,
  inboxPrompt,
  itemKey,
  MAX_DISPATCHES,
  OFFICE,
  officeData,
  inHerdrWorktree,
  onTaskBranch,
  REFRESH_EARLY_MS,
  releaseDispatch,
  releaseNotice,
  resolveCredentials,
  saveAgentPointer,
  saveCredentials,
  sessionPointerPath,
  summary,
  taskPrompt,
  taskID,
  taskLabel,
  taskShort,
  tokenRequest,
  workspaceKey,
  dispatchTarget,
  scratchTaskDir,
  scratchPrompt,
} from "./lib/office.ts"

export { agentCredentialsPath, agentPointerPath, credentialsPath, workspacePointerPath } from "./lib/office.ts"

const CALLBACK_BASE_PORT = Number(process.env.AGENT_OFFICE_CALLBACK_PORT ?? 41997)
const CALLBACK_TRIES = 20
const POLL_SCRIPT = fileURLToPath(new URL("./poll.ts", import.meta.url))
const callbackUrl = (port: number) => `http://127.0.0.1:${port}/callback`
const SCOPES = "office:read office:write office:events"
const HEARTBEAT_MS = 90_000
const APPROVAL_TIMEOUT_MS = 5 * 60_000
const DEFAULT_POLL_MS = 60_000
const SESSION_TIMEOUT_MS = 60_000
// The inbox child's first line: nobody in the human's chat sees its replies.
const INBOX_SESSION = "This is your Office inbox session, apart from the human's chat; nobody reads your replies here, so answer and ask only through Agent Office."

type PendingCallback = { resolve: (code: string) => void; reject: (err: Error) => void; close: () => void }

// A pasted callback URL can finish an approval without the browser reaching
// the loopback listener, which matters on remote machines and headless hosts.
const pendingCallbacks = new Map<string, PendingCallback>()

function completeCallback(raw: string): { handled: boolean; ok: boolean; message: string } {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return { handled: false, ok: false, message: "that is not a valid URL" }
  }
  const state = url.searchParams.get("state")
  if (!state) return { handled: false, ok: false, message: "the URL has no state" }
  const pending = pendingCallbacks.get(state)
  if (!pending) return { handled: false, ok: false, message: "no pending approval matches this URL; run /office again" }

  const code = url.searchParams.get("code")
  const error = url.searchParams.get("error_description") ?? url.searchParams.get("error")
  pendingCallbacks.delete(state)
  pending.close()
  if (error || !code) {
    pending.reject(new Error(`Office authorization failed: ${error ?? "no code"}`))
    return { handled: true, ok: false, message: `approval failed: ${error ?? "no code"}` }
  }
  pending.resolve(code)
  return { handled: true, ok: true, message: "callback accepted; finishing the Office connection" }
}

function run(command: string, args: string[], timeoutMs = 30_000): Promise<{ code: number | null; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: ["ignore", "pipe", "pipe"] })
    let stdout = ""
    let stderr = ""
    let timer: ReturnType<typeof setTimeout> | undefined
    if (timeoutMs > 0) {
      timer = setTimeout(() => {
        child.kill("SIGTERM")
        setTimeout(() => child.kill("SIGKILL"), 5_000).unref()
      }, timeoutMs)
      timer.unref()
    }
    child.stdout.on("data", (chunk) => (stdout += String(chunk)))
    child.stderr.on("data", (chunk) => (stderr += String(chunk)))
    child.on("error", (err) => {
      clearTimeout(timer)
      reject(err)
    })
    child.on("close", (code) => {
      clearTimeout(timer)
      resolve({ code, stdout, stderr })
    })
  })
}

// OpenCode's session and storage calls have no timeout of their own; one that
// never settles must not hold the poll guard for good.
function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined
  const expired = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${what} timed out after ${ms} ms`)), ms)
    timer.unref()
  })
  return Promise.race([promise, expired]).finally(() => clearTimeout(timer))
}

// The client's own check: OpenCode tags a missing session's error.
const sessionNotFound = (err: unknown) => typeof err === "object" && err !== null && (err as { _tag?: unknown })._tag === "SessionNotFoundError"

async function registerClient(callback: string, clientName: string): Promise<string> {
  const res = await fetch(`${OFFICE}/oauth/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      client_name: clientName,
      redirect_uris: [callback],
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      token_endpoint_auth_method: "none",
      scope: SCOPES,
    }),
    signal: AbortSignal.timeout(15_000),
  })
  const body = (await res.json()) as { client_id: string; error_description?: string; error?: string }
  if (!res.ok) throw new Error(`Office client registration failed: ${body.error_description ?? body.error ?? res.status}`)
  return body.client_id
}

// openBrowser opens the approval URL in the person's browser. It is the one
// place this extension reaches outside the process, so it is also the boundary
// the tests stub: with AGENT_OFFICE_OPEN_BROWSER=off nothing is spawned at all,
// which is what stops an approval test from launching a real browser on the
// host (az-iar/ao-pi#10 and #13, where a suite run left hundreds of stray
// browser processes - and the browser's own request for the approval URL then
// turns up in another test's Office call log).
export function browserOpeningAllowed(): boolean {
  const mode = (process.env.AGENT_OFFICE_OPEN_BROWSER ?? "").trim().toLowerCase()
  return mode !== "off" && mode !== "0" && mode !== "false" && mode !== "no"
}

export function openBrowser(url: string) {
  if (!browserOpeningAllowed()) return
  try {
    const [cmd, args] = process.platform === "win32" ? ["cmd", ["/c", "start", "", url]] : process.platform === "darwin" ? ["open", [url]] : ["xdg-open", [url]]
    spawn(cmd, args, { detached: true, stdio: "ignore" }).on("error", () => {}).unref()
  } catch {}
}

// Binds a loopback callback for this state, falling back to the next free port
// when another OpenCode plugin instance already holds the preferred one.
function listenForCode(state: string): Promise<{ port: number; code: Promise<string>; close: () => void }> {
  return new Promise((resolve, reject) => {
    const attempt = (offset: number) => {
      if (offset >= CALLBACK_TRIES) {
        reject(new Error(`cannot listen on ports ${CALLBACK_BASE_PORT}-${CALLBACK_BASE_PORT + CALLBACK_TRIES - 1}`))
        return
      }
      const port = CALLBACK_BASE_PORT + offset
      const callback = callbackUrl(port)
      let resolveCode!: (code: string) => void
      let rejectCode!: (err: Error) => void
      const code = new Promise<string>((resolve, reject) => {
        resolveCode = resolve
        rejectCode = reject
      })
      let settled = false
      let timer: ReturnType<typeof setTimeout> | undefined
      const server = createServer((req, res) => {
        const result = completeCallback(`http://127.0.0.1:${port}${req.url ?? "/"}`)
        if (!result.handled) {
          res.writeHead(404).end()
          return
        }
        res.writeHead(result.ok ? 200 : 400, { "Content-Type": "text/html; charset=utf-8" })
        res.end(`<!doctype html><title>Agent Office</title><p>${result.ok ? "OpenCode is connected to Agent Office. You can close this tab." : `Not connected: ${result.message}`}</p>`)
      })
      server.once("error", (err: NodeJS.ErrnoException) => {
        clearTimeout(timer)
        if (!server.listening && err.code === "EADDRINUSE") {
          attempt(offset + 1)
          return
        }
        if (!settled) {
          settled = true
          pendingCallbacks.delete(state)
          rejectCode(new Error(`cannot listen on ${callback}: ${err.message}`))
        }
      })
      server.listen(port, "127.0.0.1", () => {
        if (settled) return
        timer = setTimeout(() => {
          if (settled) return
          settled = true
          pendingCallbacks.delete(state)
          server.close()
          rejectCode(new Error("timed out waiting for the approval in the browser"))
        }, APPROVAL_TIMEOUT_MS)
        pendingCallbacks.set(state, {
          resolve: resolveCode,
          reject: rejectCode,
          close: () => {
            clearTimeout(timer)
            server.close()
          },
        })
        resolve({
          port,
          code,
          close: () => {
            pendingCallbacks.delete(state)
            clearTimeout(timer)
            server.close()
          },
        })
      })
    }
    attempt(0)
  })
}

// Starts an approval and returns its URL at once. The tokens follow when the
// browser reaches the loopback callback or the callback URL is pasted back,
// so a remote machine never blocks on a browser it cannot open.
async function authorize(clientName: string): Promise<{ url: string; state: string; credentials: Promise<Credentials>; close: () => void }> {
  const verifier = randomBytes(32).toString("base64url")
  const state = randomBytes(16).toString("base64url")
  const listener = await listenForCode(state)
  const callback = callbackUrl(listener.port)
  let client_id: string
  try {
    client_id = await registerClient(callback, clientName)
  } catch (err) {
    listener.close()
    throw err
  }
  const url = new URL(`${OFFICE}/oauth/authorize`)
  url.search = new URLSearchParams({
    response_type: "code",
    client_id,
    redirect_uri: callback,
    scope: SCOPES,
    state,
    code_challenge: createHash("sha256").update(verifier).digest("base64url"),
    code_challenge_method: "S256",
    resource: `${OFFICE}/mcp`,
  }).toString()
  openBrowser(url.href)
  const credentials = listener.code.then((code) =>
    tokenRequest(client_id, { grant_type: "authorization_code", code, redirect_uri: callback, code_verifier: verifier }),
  )
  return { url: url.href, state, credentials, close: listener.close }
}

// A pasted callback, tolerating the quotes or angle brackets a terminal or
// chat client may add around it.
const pastedCallback = (text: string) => {
  const url = text.trim().replace(/^[<"'`]+|[>"'`]+$/g, "")
  return /^http:\/\/(127\.0\.0\.1|localhost):\d+\/callback\?\S*$/.test(url) ? url : undefined
}

function approvalText(name: string, url: string) {
  return [
    `Open this link and approve "${name}" as the Office admin:`,
    url,
    "",
    "On this machine the browser opens it by itself and the connection finishes on its own.",
    "On a remote machine, open the link in any browser. After you approve, the browser lands on a page that does not load: copy that page's full address (http://127.0.0.1:.../callback?...) and paste it here, or run /office followed by the address.",
    `The link works for ${APPROVAL_TIMEOUT_MS / 60_000} minutes. Show the user the link exactly as written.`,
  ].join("\n")
}


// The child name the office will accept: lowercase, no leading punctuation, only
// [a-z0-9._-], and at most 64 characters - #512 answers 400 outside that rule,
// so normalising here is what stops a launcher's name turning a check-in into a
// refusal. Sunny's fallback, one across the three extensions: with no name, the
// child is `s-` plus the first eight hex of its own session id.
export function normaliseChildName(raw: string | undefined, sessionID: string | undefined): string | undefined {
  const hex = (sessionID ?? "").replace(/[^0-9a-fA-F]/g, "")
  // A session id too short to give eight hex characters has no honest name.
  const source = raw?.trim() || (hex.length >= 8 ? `s-${hex.slice(0, 8)}` : "")
  // `s-` keeps its dash because only *leading* non-alphanumerics are stripped;
  // a trailing separator is dropped so "Office Inbox!" does not end in a dash.
  const name = source
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/^[^a-z0-9]+/, "")
    .replace(/[^a-z0-9]+$/, "")
    .slice(0, 64)
    // Again after the cut: 64 characters can land on a separator, and while the
    // office accepts that, a name ending in "-" reads like a bug in a header.
    .replace(/[^a-z0-9]+$/, "")
  return name || undefined
}

// identityHeaders is the header set for one Office call. Pure, so the rules can
// be tested without a folder, a token or an approval: no parent means no via at
// all rather than a half-header, and a child sends a via only when it also has a
// parent to name.
export function identityHeaders(input: { approved?: string; child?: boolean; childName?: string; sessionID?: string; runtime: string }): Record<string, string> {
  const headers: Record<string, string> = {}
  if (input.approved) headers["X-Office-Session-Agent"] = input.approved
  if (input.child && input.approved) {
    const child = normaliseChildName(input.childName, input.sessionID)
    if (child) headers["X-Office-Via"] = `child=${child}; parent=${input.approved}; runtime=${input.runtime}`
  }
  return headers
}

export default Plugin.define({
  id: "agent-office",
  async setup(ctx) {
    // Tokens are never part of a log line; errors only carry status codes,
    // messages from the Office and the names of the calls that failed.
    const warn = (message: string, extra?: Record<string, unknown>) => {
      console.warn(`[agent-office] ${message}${extra ? ` ${JSON.stringify(extra)}` : ""}`)
    }
    const life = new AbortController()
    const locationKey = workspaceKey(ctx.location.directory)
    const clientName = process.env.AGENT_OFFICE_CLIENT_NAME ?? `${basename(ctx.location.directory) || "OpenCode"}@${hostname().split(".")[0]}`

    const resolved = resolveCredentials(ctx.location.directory, ctx.location.project?.canonical)
    const { legacyCredentialsFile, pointerFile, repoPointerFile } = resolved
    let currentSessionID: string | undefined
    let approvedAgent: string | undefined
    let credentialsFile = resolved.credentialsFile
    const usingCredentialsOverride = Boolean(CREDENTIALS_OVERRIDE)
    const auth = createBearer(() => credentialsFile, resolved.credentials)
    // A linked worktree nobody connected has no Office identity at all, and
    // must not borrow the checkout's: no polling, no status reports, no
    // dispatch and no office MCP server (whose tools would act as that agent).
    // A `/office` run here (adoptAgent) or a token in the environment is an
    // explicit identity; a task worktree's worker is handed one by its
    // dispatcher (adoptWorktree). Assigned once polling is set up, below, so a
    // worktree connected by /office starts polling without a restart.
    const noIdentity = () => resolved.unconnected && !auth.credentials && !process.env.AGENT_OFFICE_TOKEN
    let startPolling: () => void = () => {}
    // Identity is per session (#491, ao-pi#3). A session owns one when an
    // approval is recorded against its own id - `sessions/<id>.json`, the same
    // record ao-pi keeps - or when an approval finished in this process for it.
    // The folder pointer and the checkout's token are pre-fill only: a child
    // session inherits both, and inheriting them is the bug (#476, #491).
    let ownedSessionID: string | undefined
    let approvedSessionID: string | undefined
    // An approval recorded before this change was kept for the folder, not a
    // session. OpenCode's own storage names the last session that worked here
    // as that agent, so that session keeps the approval it acted under; every
    // other session approves for itself. Read once, in setup.
    let legacySessionID: string | undefined
    let legacyAgent: string | undefined
    // owns is whether this process may act as an Office agent at all. True only
    // when a session it serves owns an approval, or an identity was handed to
    // the process on purpose: AGENT_OFFICE_TOKEN for a non-child (the runner's
    // own identity), AGENT_OFFICE_OPENCODE_CREDENTIALS (a dispatcher's grant),
    // or a connected worktree's `dispatch` mark (the identity its dispatcher
    // wrote there for this worker). A `/office` mark is not a grant: that
    // approval is recorded against the session that ran it.
    let owns = false
    const sessionAgent = (id: string | undefined) => (id ? loadAgentPointer(sessionPointerPath(id)) : undefined)
    const tokenGranted = () => Boolean(process.env.AGENT_OFFICE_TOKEN) && process.env.AGENT_OFFICE_CHILD !== "1"
    const dispatchGranted = () => Boolean(resolved.worktree && agentPointerVia(pointerFile) === "dispatch")
    const granted = () => tokenGranted() || usingCredentialsOverride || dispatchGranted()
    // Write an approval against the session that owns it, never against
    // whoever happens to be speaking: a process acting on a token or a
    // dispatcher's grant has no session approval to record.
    const rememberSessionApproval = (slug: string) => {
      // A grant is not an approval: under a token or a dispatcher's credentials
      // file the process already has an identity, and nothing is recorded
      // against a session (ao-pi's adoptAgent returns before its pointer too).
      if (usingCredentialsOverride || tokenGranted()) return
      const id = approvedSessionID ?? ownedSessionID
      if (!id || loadAgentPointer(sessionPointerPath(id)) === slug) return
      saveAgentPointer(sessionPointerPath(id), slug)
    }
    // `via` is the mark a linked worktree's pointer needs to count: "office"
    // for a `/office` here. A plain poll passes none and, in a worktree, writes
    // no pointer unless one there is already marked (a dispatcher's worker
    // keeps its identity). With a token in the environment a worktree writes
    // none at all. The main checkout keeps the unmarked pointer it always had.
    const adoptAgent = (slug?: string, via?: "office" | "dispatch") => {
      // The office told us which agent this token is, so keep it even when the
      // credentials override means there is no pointer to write - the identity
      // header needs it (GH-476). A dispatcher hands a worker its token and no
      // agent pointer, and without this the header would simply be missing.
      if (slug) approvedAgent = slug
      if (slug) rememberSessionApproval(slug)
      if (usingCredentialsOverride || !slug) return
      // In a linked worktree the token is an identity handed in from outside,
      // not this workspace's: a poll with it ("/office" here is different)
      // must not leave a pointer or another agent's token behind.
      if (resolved.worktree && process.env.AGENT_OFFICE_TOKEN && !via) return startPolling()
      const target = agentCredentialsPath(slug)
      if (auth.credentials) saveCredentials(target, auth.credentials)
      credentialsFile = target
      if (!resolved.worktree) saveAgentPointer(pointerFile, slug)
      else {
        const mark = via ?? agentPointerVia(pointerFile)
        if (mark) saveAgentPointer(pointerFile, slug, mark)
      }
      startPolling()
    }
    const dispatchDir = () => agentStateDir(legacyCredentialsFile, "dispatch", ownAgent?.slug ?? locationKey)
    const noticeDir = () => agentStateDir(legacyCredentialsFile, "notices", ownAgent?.slug ?? locationKey)
    // New mail goes to the open session as an ids-only notice by default, like
    // ao-claude's session inbox; "headless" keeps the child-session worker,
    // "off" leaves mail unread for a human.
    const configuredInbox = String(ctx.options.inbox ?? process.env.AGENT_OFFICE_INBOX ?? "session").toLowerCase()
    const inboxMode = configuredInbox === "headless" ? "headless" : configuredInbox === "off" ? "off" : "session"
    const configuredDispatch = String(ctx.options.autoDispatch ?? process.env.AGENT_OFFICE_AUTODISPATCH ?? "on").toLowerCase()
    // ao-claude never queues a task into the user's session: every mode spawns
    // its own worker in the worktree. "session" and "herdr" are kept as aliases
    // for hc, so an old setting still dispatches and no longer edits the main
    // checkout.
    const dispatchMode = configuredDispatch === "off" || configuredDispatch === "false" ? "off" : "hc"
    const bearer = auth.bearer
    // ownedBearer is the only door to identity: while this process owns none it
    // answers nothing, so checkin, the office call, the MCP registration, the
    // status reports and polling all fail closed instead of each resolving the
    // folder pointer's token itself (#491).
    const ownedBearer = async () => (owns ? bearer() : undefined)
    // officeIdentity is the header set every Office call carries (GH-476): the
    // agent this session's own approval named, and, for a child session, the
    // parent it was started from plus its runtime.
    //
    // Session-scoped now: the agent comes from the owning session's pointer -
    // the same record that lets it act - and never from the folder pointer,
    // which is pre-fill only. Under a token or a credentials override the
    // token's own agent is the only truth: the pointer may name a different
    // agent, and mixing A's slug with B's bearer is a 409 (and a wrong parent
    // for a child). A dispatched worktree's worker is the exception: its
    // identity is the marked pointer its dispatcher wrote there.
    const officeIdentity = (): Record<string, string> =>
      identityHeaders({
        approved:
          process.env.AGENT_OFFICE_TOKEN || usingCredentialsOverride
            ? approvedAgent
            : (sessionAgent(ownedSessionID) ?? (resolved.worktree ? loadAgentPointer(pointerFile) : undefined) ?? approvedAgent),
        childName: process.env.AGENT_OFFICE_CHILD_NAME,
        child: process.env.AGENT_OFFICE_CHILD === "1",
        sessionID: currentSessionID,
        runtime: "opencode",
      })
    const { call, checkin } = createClient(ownedBearer, officeIdentity)

    // Hands OpenCode the office MCP server with the current access token, and
    // again whenever the token is refreshed. The transform captures the token
    // and `reload()` replays it after every refresh.
    let token: string | undefined
    await ctx.mcp.transform((editor) => {
      if (!token) return
      // Same identity headers as createClient: a tool call through this server
      // has to look like a call from this session (GH-476).
      editor.set("office", { type: "remote", url: `${OFFICE}/mcp`, oauth: false, headers: { Authorization: `Bearer ${token}`, ...officeIdentity() } })
    })
    let refreshTimer: ReturnType<typeof setTimeout> | undefined
    // syncMcpIdentity replays the registration only when the identity headers
    // changed: connectedSummary is not just the connect path (a reconnect and
    // /office use it too), so a blind reload would churn the MCP server on every
    // summary. Comparing the header set also makes the second-/mcp assertion in
    // the test deterministic rather than order-dependent.
    let lastIdentity = ""
    // Callers run this after a checkin, which is also what sets `token`; the
    // reload is a no-op without one, so the assumption is only that they got
    // this far into the connect path.
    const syncMcpIdentity = async () => {
      if (!token) return
      const next = JSON.stringify(officeIdentity())
      if (next === lastIdentity) return
      lastIdentity = next
      await ctx.mcp.reload()
    }
    const connectMcp = async () => {
      // Office tools act as the agent they carry, so an unconnected worktree
      // gets none of them, and neither does a process with no identity it owns:
      // registering the server is what puts the tools in front of a session (#491).
      if (noIdentity() || !owns) return
      token = await ownedBearer()
      if (!token) return
      // Record what is being registered here too: a session that already had a
      // pointer would otherwise pay one redundant reload on its first summary,
      // because syncMcpIdentity would see a "change" from its empty starting value.
      lastIdentity = JSON.stringify(officeIdentity())
      await ctx.mcp.reload()
      clearTimeout(refreshTimer)
      if (auth.credentials && !process.env.AGENT_OFFICE_TOKEN) {
        const delay = Math.max(auth.credentials.expires_at - Date.now() - REFRESH_EARLY_MS, 10_000)
        refreshTimer = setTimeout(() => connectMcp().catch((err) => warn("could not refresh the Office connection", { error: String(err) })), delay)
        refreshTimer.unref()
      }
    }
    let connected: Promise<void> | undefined
    const ensureConnected = () => (connected ??= connectMcp().catch((err) => warn("could not connect the Office MCP server", { error: String(err) })))

    const note = `opencode: ${basename(ctx.location.directory)}`
    // Tests shorten it.
    const configuredTimeout = Number(ctx.options.sessionTimeout ?? SESSION_TIMEOUT_MS)
    const sessionTimeout = Number.isFinite(configuredTimeout) && configuredTimeout > 0 ? configuredTimeout : SESSION_TIMEOUT_MS
    const bounded = <T>(promise: Promise<T>, what: string) => withTimeout(promise, sessionTimeout, what)
    const busy = new Set<string>()
    let lastSessionID: string | undefined
    let heartbeat: ReturnType<typeof setInterval> | undefined
    let pollTimer: ReturnType<typeof setInterval> | undefined
    let queue = Promise.resolve()
    let lastPollAt: number | undefined
    let lastPollError: string | undefined
    let lastEventAt: number | undefined
    let lastEventType: string | undefined
    let lastEventDirectory: string | undefined
    let lastEventSessionID: string | undefined
    let lastPromptAt: number | undefined
    let lastPromptError: string | undefined
    let lastPending: Record<string, number> | undefined
    let ownAgent: { id?: string; slug?: string } | undefined

    // Reports run one after another so a short turn cannot land idle before working.
    const report = (status: "working" | "idle") => {
      queue = queue.then(async () => {
        try {
          if (!(await ownedBearer())) return
          if (status === "idle" && (await checkin()).agent?.work_status === "blocked") return
          await checkin({ status, note })
        } catch (err) {
          warn(`could not report ${status}`, { error: String(err) })
        }
      })
      return queue
    }

    const start = (sessionID: string) => {
      // The child name falls back to this when no launcher name was given.
      currentSessionID = sessionID
      noticeUnconnected(sessionID)
      if (busy.size === 0) {
        heartbeat = setInterval(() => {
          if (owns) checkin().catch(() => {})
        }, HEARTBEAT_MS)
        heartbeat.unref()
        void report("working")
      }
      busy.add(sessionID)
    }

    const stop = (sessionID: string) => {
      if (!busy.delete(sessionID) || busy.size > 0) return
      clearInterval(heartbeat)
      void report("idle")
    }

    // Inbox children by id, each mapped to the root session it works for.
    const inboxSessions = new Map<string, string>()
    // Inbox children this plugin started count as busy work, but never as a target.
    const officeChild = (sessionID: string) => inboxSessions.has(sessionID)
    const targetSession = () => [...busy].filter((id) => !id.startsWith("herdr:") && !officeChild(id)).at(-1) ?? lastSessionID
    // One OpenCode server can stream events for several windows, so a session
    // counts only when OpenCode places it in this folder or this plugin
    // started it.
    const ownership = new Map<string, Promise<boolean>>()
    const ownsSession = (sessionID: string) => {
      if (officeChild(sessionID)) return Promise.resolve(true)
      let owned = ownership.get(sessionID)
      if (!owned) {
        owned = bounded(ctx.session.get({ sessionID }), "session.get").then(
          // A subagent's session is its parent's business, never a mail target.
          (info) => !info.parentID && info.location.directory === ctx.location.directory,
          () => {
            ownership.delete(sessionID)
            return false
          },
        )
        ownership.set(sessionID, owned)
      }
      return owned
    }
    const sessionStoreKey = () => (ownAgent?.slug ? `last-session:${ownAgent.slug}:${locationKey}` : undefined)
    const targetSessionAsync = async () => {
      const known = targetSession()
      const key = sessionStoreKey()
      if (known || !key) return known
      const stored = await bounded(ctx.storage.get(key), "storage.get").catch(() => undefined)
      if (typeof stored !== "string" || !(await ownsSession(stored))) return undefined
      lastSessionID = stored
      return stored
    }
    const rememberSession = (sessionID: string) => {
      lastSessionID = sessionID
      const key = sessionStoreKey()
      if (key) void ctx.storage.set(key, sessionID).catch(() => {})
    }
    const notify = (sessionID: string | undefined, text: string) => {
      if (!sessionID) return
      void ctx.session.synthetic({ sessionID, text, description: "Agent Office" }).catch((err) => warn("could not notify the session", { error: String(err) }))
    }

    // An unconnected worktree just has no office tools; say why once in the
    // session, and only when the checkout it belongs to really is connected,
    // so this is not noise in a project that never used Agent Office.
    const unconnectedNoticed = new Set<string>()
    const noticeUnconnected = (sessionID: string) => {
      if (!noIdentity() || unconnectedNoticed.has(sessionID)) return
      const checkoutAgent = loadAgentPointer(repoPointerFile)
      if (!checkoutAgent) return
      unconnectedNoticed.add(sessionID)
      notify(sessionID, `This worktree is not connected to Agent Office (the main checkout is, as @${checkoutAgent}), so it has no office tools, polling or notices. Run /office here to connect it.`)
    }

    // Office children run on their root session's model, not on OpenCode's
    // default. The agent is not copied: a read-only one such as Plan would
    // leave the child's permission asks where nobody is watching, and the
    // SDK does not say which agents can edit. A root without an explicit
    // model still gets its child, on the default model, and its chat is told
    // once.
    const defaultModelNoticed = new Set<string>()
    // Roots with a child being created; a deletion meanwhile marks them.
    const creating = new Set<{ parent: string; deleted: boolean }>()
    const createChild = async (parent: string, title: string, location?: { directory: string }) => {
      const pending = { parent, deleted: false }
      creating.add(pending)
      try {
        const info = await bounded(ctx.session.get({ sessionID: parent }), "session.get")
        const created = ctx.session.create({
          parentID: parent,
          title,
          location: location ?? { directory: info.location.directory },
          ...(info.model ? { model: info.model } : {}),
        })
        // A create that lands after its timeout is removed, so no empty child is left.
        const child = await bounded(created, "session.create").catch((err) => {
          void created.then((late) => ctx.session.remove({ sessionID: late.id })).catch(() => {})
          throw err
        })
        if (pending.deleted) {
          void ctx.session.remove({ sessionID: child.id }).catch(() => {})
          throw new Error(`session ${parent} was deleted while its Office child was created`)
        }
        if (!info.model && !defaultModelNoticed.has(parent)) {
          defaultModelNoticed.add(parent)
          notify(parent, "Office work is using OpenCode's default model; pick a model in this chat to choose one.")
        }
        return child
      } finally {
        creating.delete(pending)
      }
    }

    // New mail runs in one inbox child per root session, reused while it
    // lives, so Office work stays out of the human's chat history. Only a
    // session OpenCode reports missing is replaced; any other error is retried.
    // A queue into the reused child can land after it timed out, so it is
    // handed back in `sent` for the caller to settle the mail from.
    const deliverInbox = async (parent: string, text: string, sent: { queued?: Promise<unknown> } = {}) => {
      let sessionID = [...inboxSessions].find(([, root]) => root === parent)?.[0]
      if (sessionID) {
        const gone = await bounded(ctx.session.get({ sessionID }), "session.get").then(
          () => false,
          (err) => {
            if (sessionNotFound(err)) return true
            throw err
          },
        )
        if (gone) {
          inboxSessions.delete(sessionID)
          sessionID = undefined
        }
      }
      if (sessionID) {
        sent.queued = ctx.session.prompt({ sessionID, text, delivery: "queue" })
        await bounded(sent.queued, "session.prompt")
        return
      }
      const child = await createChild(parent, "Office inbox")
      inboxSessions.set(child.id, parent)
      try {
        await bounded(ctx.session.prompt({ sessionID: child.id, text, delivery: "queue" }), "session.prompt")
      } catch (err) {
        inboxSessions.delete(child.id)
        void ctx.session.remove({ sessionID: child.id }).catch(() => {})
        throw err
      }
    }

    // `via` marks a worktree pointer as written on purpose, so it is passed
    // only where a person connected this workspace: after an approval. A
    // `/office` that finds a bearer already (say AGENT_OFFICE_TOKEN) marks
    // nothing, or the token's identity would outlive the token.
    const connectedSummary = async (via?: "office" | "dispatch") => {
      await ensureConnected()
      const me = await checkin()
      adoptAgent(me.agent?.slug, via)
      ownAgent = me.agent
      await syncMcpIdentity()
      await observePending(me.pending, false)
      const pending = me.pending ?? {}
      const summary = `Connected to Agent Office as @${me.agent?.slug} (${me.agent?.name}). Open tasks: ${pending.open_tasks?.length ?? 0}, unclaimed team tasks: ${pending.unclaimed_team_tasks?.length ?? 0}, unread messages: ${pending.unread_messages?.length ?? 0}. The office MCP tools are available; call checkin first.`
      if (!existsSync(POLL_SCRIPT)) return summary
      return `${summary}\nTo keep handling Office work while OpenCode is closed, add this line to the crontab (crontab -e), or ask to have it added:\n${cronLine(ctx.location.directory, POLL_SCRIPT)}`
    }

    // The approval in flight. A loopback callback finishes it in the
    // background and tells the session; a pasted URL reports it in reply.
    let approval: { state: string; done: Promise<string>; pasted: boolean; close: () => void } | undefined
    const startApproval = async (approvalName: string, sessionID: string | undefined) => {
      // A new approval replaces the old one, so its port is freed at once.
      approval?.close()
      const started = await authorize(approvalName)
      const current = {
        state: started.state,
        close: started.close,
        pasted: false,
        done: started.credentials.then(async (next) => {
          auth.credentials = next
          saveCredentials(credentialsFile, next)
          // The approval is this session's own grant, held from the moment its
          // credentials are persisted and before anything acts on them. A grant
          // that could not be recorded is not one this process holds (#476).
          owns = true
          if (sessionID) {
            approvedSessionID = sessionID
            ownedSessionID = sessionID
          }
          connected = undefined
          return connectedSummary("office")
        }),
      }
      approval = current
      current.done.then(
        (summary) => current.pasted || notify(sessionID, summary),
        (err) => current.pasted || notify(sessionID, `Agent Office approval did not finish: ${err instanceof Error ? err.message : String(err)}. Run /office again.`),
      )
      return started.url
    }
    const finishWithCallback = async (raw: string) => {
      let state = ""
      try {
        state = new URL(raw.trim()).searchParams.get("state") ?? ""
      } catch {}
      // Another OpenCode in this process may own the approval this URL answers.
      const current = approval?.state === state ? approval : undefined
      if (current) current.pasted = true
      const result = completeCallback(raw.trim())
      if (!result.handled || !result.ok) {
        if (current) current.pasted = false
        return result.message
      }
      if (!current) return `${result.message} in the OpenCode that started it.`
      try {
        return await current.done
      } catch (err) {
        return `Agent Office approval did not finish: ${err instanceof Error ? err.message : String(err)}. Run /office again.`
      }
    }

    const backlog = new Set<string>()
    const dispatching = new Set<string>()

    const repo = ctx.location.project?.canonical ?? ctx.location.directory

    // Herdr-core mode: one `hc` child does the whole task. The worktree is
    // ours (git), and `hc run --bg` starts an OpenCode with the task as its
    // first prompt, so there is no supervisor and no worker subagent.
    // dispatchScratch runs a repo-less task in a directory that is not a git
    // checkout: no branch, no worktree, and a prompt that says so (agent-office#556).
    const dispatchScratch = async (id: string, task: any, dir: string, parent: string) => {
      // The scratch dir needs the same Office identity the worktree path gives its
      // worker: the prompt tells the worker to post_comment and set the task
      // blocked, and without a pointer there is no agent to do it as (neo, #32).
      // A missing identity is loud rather than a silent no-op: launching a worker
      // that cannot post is worse than not launching it.
      if (!usingCredentialsOverride && !ownAgent?.slug) {
        throw new Error(`cannot dispatch task ${id} into ${dir}: this session has no Office agent yet, so the worker would have no identity to post with`)
      }
      adoptWorktree(dir, usingCredentialsOverride ? undefined : ownAgent?.slug)
      const name = `office-${taskShort(id)}-${randomBytes(2).toString("hex")}`
      const started = await run(
        hcBin(),
        ["run", "--bg", "--outlive-parent", "--no-notify", "--name", name, "--label", taskLabel(id), "--cwd", dir, "opencode", "--prompt", scratchPrompt(id, task, dir)],
        60_000,
      )
      if (started.code !== 0) throw new Error(`hc run failed: ${started.stderr || started.stdout}`)
      holdDispatch(dispatchDir(), id, { terminal: name })
      notify(parent, `Office task ${id} dispatched to hc worker ${name} with no worktree at ${dir}.`)
    }

    const dispatchHc = async (id: string, task: any, title: string, parent: string) => {
      // A repo-less task gets no worktree (Azri's card 53828e66): it runs in a
      // scratch directory with no repository, and the prompt says so, rather than
      // borrowing this session's checkout (agent-office#556).
      const target = dispatchTarget(repo, String(task?.repo ?? ""))
      if (target.kind === "scratch") {
        const scratch = scratchTaskDir(id)
        mkdirSync(scratch, { recursive: true })
        return dispatchScratch(id, task, scratch, parent)
      }
      const { path, branch } = await addWorktree(target.repo, id, title)
      // The worker's OpenCode opens in a linked worktree, which has no Office
      // identity unless it is given one: it gets this agent's, marked as a
      // dispatch so the worktree keeps it. Under the credentials override
      // nothing is adopted: an hc worker cannot inherit that variable, so a
      // pointer here would name an identity the worker cannot use.
      adoptWorktree(path, usingCredentialsOverride ? undefined : ownAgent?.slug)
      // Names must be unique among live terminals: a retry while the first
      // child still runs would otherwise fail the launch.
      const name = `office-${taskShort(id)}-${randomBytes(2).toString("hex")}`
      const started = await run(
        hcBin(),
        ["run", "--bg", "--outlive-parent", "--no-notify", "--name", name, "--label", taskLabel(id), "--cwd", path, "opencode", "--prompt", taskPrompt(id, task, path, branch)],
        60_000,
      )
      if (started.code !== 0) throw new Error(`hc run failed: ${started.stderr || started.stdout}`)
      // The slot outlives this process: the lock names the terminal, and the
      // next acquire frees it once hc says it is finished or gone.
      holdDispatch(dispatchDir(), id, { terminal: name })
      notify(parent, `Office task ${id} dispatched to hc worker ${name} (${branch}) at ${path}.`)
    }

    const dispatchTask = async (id: string) => {
      let claimed = false
      let held = false
      let task: any
      const parent = await targetSessionAsync()
      try {
        if (!parent) return
        const detail = await call("get_task", { task_id: id }).catch(() => undefined)
        task = detail?.task ?? detail
        const title = String(task?.title ?? id)
        // Another dispatcher (ao-claude, a hand-started hc worker) may already
        // run this task: skip it unclaimed when it is no longer todo or a live
        // terminal carries its ao.task label. It stays seen, so it is not
        // retried; a task that comes back to todo is new to the next checkin.
        // Not airtight: two dispatchers can both pass in the moment before
        // the label is set, and the claim is the last line.
        const refused = dispatchRefusal(id, task, true)
        if (refused) {
          warn("not dispatching Office task", { task: id, reason: refused })
          notify(parent, `Not dispatching Office task ${id}: ${refused}.`)
          return
        }
        await claimTask(call, id)
        claimed = true
        await dispatchHc(id, task, title, parent)
        held = true
      } catch (err) {
        warn("could not dispatch Office task", { task: id, error: String(err) })
        notify(parent, `Could not dispatch Office task ${id}: ${String(err)}`)
        // The finally drops it from the dispatch state, and it stays seen, so
        // a lasting failure is not claimed again.
        if (claimed) await blockTask(call, id, err, task, ownAgent?.slug).catch((blockErr) => warn("could not mark the Office task blocked", { task: id, error: String(blockErr) }))
      } finally {
        if (!held) releaseDispatch(dispatchDir(), id)
        // The cap is the lock files, not this set: holding the id here would
        // stop every later pump in a long-lived OpenCode.
        dispatching.delete(id)
        pumpDispatch()
      }
    }

    // One pump at a time, including while it looks up the remembered session.
    let pumping = false
    const pumpDispatch = () => {
      if (life.signal.aborted || dispatchMode === "off" || noIdentity() || !owns || pumping) return
      pumping = true
      if (!targetSession()) {
        void targetSessionAsync()
          .catch(() => undefined)
          .then((id) => {
            pumping = false
            if (id) pumpDispatch()
          })
        return
      }
      try {
        for (const id of [...backlog]) {
          if (dispatching.size >= MAX_DISPATCHES) return
          const acquired = acquireDispatch(dispatchDir(), id)
          if (acquired === "full") return
          backlog.delete(id)
          if (acquired === "busy") continue
          dispatching.add(id)
          void dispatchTask(id)
        }
      } finally {
        pumping = false
      }
    }

    const seenMessages = new Set<string>()
    const seenTasks = new Set<string>()
    const seenAssigned = new Set<string>()
    let pollPrimed = false

    // Polling reports only new items. The first successful checkin is a
    // baseline: existing unread mail is already visible through checkin.
    const observePending = async (pending: any, live: boolean) => {
      const messages: any[] = (Array.isArray(pending?.unread_messages) ? pending.unread_messages : []).filter((item: any) => fromOthers(item, ownAgent))
      const tasks: any[] = Array.isArray(pending?.unclaimed_team_tasks) ? pending.unclaimed_team_tasks : []
      const assigned: any[] = Array.isArray(pending?.open_tasks) ? pending.open_tasks : []
      const freshMessages = messages.filter((item) => !seenMessages.has(itemKey(item, "message")))
      const freshTasks = tasks.filter((item) => !seenTasks.has(itemKey(item, "task")))
      const freshAssigned = assigned.filter((item) => !seenAssigned.has(itemKey(item, "task")))
      lastPending = { messages: messages.length, unclaimed: tasks.length, assigned: assigned.length }

      if (!pollPrimed) {
        for (const item of tasks) seenTasks.add(itemKey(item, "task"))
        for (const item of assigned) seenAssigned.add(itemKey(item, "task"))
        freshTasks.length = 0
        freshAssigned.length = 0
        pollPrimed = true
      }
      if (!live || (freshMessages.length === 0 && freshTasks.length === 0 && freshAssigned.length === 0)) return

      const sessionID = await targetSessionAsync()
      if (!sessionID) return // keep the items unseen until a session can see them

      for (const item of tasks) seenTasks.add(itemKey(item, "task"))
      for (const item of assigned) seenAssigned.add(itemKey(item, "task"))

      for (const item of freshAssigned) {
        if (String(item?.status ?? "todo").toLowerCase() === "todo") backlog.add(taskID(item))
      }

      if (freshMessages.length > 0 && inboxMode === "off") {
        // Left unread on purpose: the ids stay seen so one batch is not
        // announced on every poll, and a human sees them in Office.
        for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
      } else if (freshMessages.length > 0 && inboxMode === "session") {
        try {
          await bounded(ctx.session.prompt({ sessionID, text: inboxNotice(freshMessages), delivery: "queue" }), "session.prompt")
          lastPromptAt = Date.now()
          lastPromptError = undefined
          for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
        } catch (err) {
          lastPromptError = String(err)
          warn("could not hand new Office messages to the session", { error: String(err) })
        }
      } else if (freshMessages.length > 0) {
        const ids = freshMessages.map((item) => itemKey(item, "message"))
        if (!claimNotice(noticeDir(), ids)) {
          for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
        } else {
          const text = `${INBOX_SESSION}\n${inboxPrompt(freshMessages)}`
          const sent: { queued?: Promise<unknown> } = {}
          try {
            await deliverInbox(sessionID, text, sent)
            lastPromptAt = Date.now()
            lastPromptError = undefined
            for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
          } catch (err) {
            if (sent.queued) {
              // The reused child may still get this mail: it stays claimed and
              // seen, and is released for a retry only if the queue fails. A
              // queue that never settles keeps its mail seen here until a
              // restart, and the cron poller can take it once the notice
              // expires after 10 minutes: the price of never queueing it twice.
              for (const id of ids) seenMessages.add(id)
              void sent.queued.catch(() => {
                for (const id of ids) seenMessages.delete(id)
                releaseNotice(noticeDir(), ids)
              })
            } else releaseNotice(noticeDir(), ids)
            lastPromptError = String(err)
            warn("could not wake the session for new Office messages", { error: String(err) })
          }
        }
      }

      if (freshTasks.length > 0 || freshAssigned.length > 0) {
        const listed = [
          ...freshTasks.slice(0, 5).map((item) => ({ kind: "unclaimed", id: itemKey(item, "task"), title: summary(item) })),
          ...freshAssigned.slice(0, 5).map((item) => ({ kind: "assigned", id: itemKey(item, "task"), status: String(item?.status ?? "todo"), title: summary(item) })),
        ]
        notify(sessionID, `Agent Office poll: ${freshTasks.length} new unclaimed task(s), ${freshAssigned.length} new assigned task(s). Use list_tasks in the office MCP server. As data:\n${officeData(listed)}`)
      }
      if (backlog.size > 0) pumpDispatch()
    }

    // A slow poll (a queued prompt, a stuck checkin) must not overlap the next.
    let polling = false
    const poll = async () => {
      if (polling) return
      polling = true
      try {
        if (!(await ownedBearer())) return
        const me = await checkin()
        adoptAgent(me.agent?.slug)
        ownAgent = me.agent
        await syncMcpIdentity()
        lastPollAt = Date.now()
        lastPollError = undefined
        await observePending(me.pending, true)
      } catch (err) {
        lastPollError = String(err)
        warn("could not poll the Office inbox", { error: String(err) })
      } finally {
        polling = false
      }
    }

    const configuredPoll = Number(ctx.options.pollInterval ?? process.env.AGENT_OFFICE_POLL_MS ?? DEFAULT_POLL_MS)
    const pollInterval = Number.isFinite(configuredPoll) ? Math.max(configuredPoll, 1_000) : DEFAULT_POLL_MS
    // Only the main checkout reads the inbox: not a session on an office/
    // branch, not a Herdr worktree, not a worktree nobody connected, and not a
    // process whose session has no identity of its own (#491).
    startPolling = () => {
      if (pollTimer || noIdentity() || !owns || onTaskBranch(ctx.location.directory) || inHerdrWorktree(ctx.location.directory)) return
      pollTimer = setInterval(() => void poll(), pollInterval)
      pollTimer.unref()
      void poll()
    }
    // An identity handed to the process on purpose is owned from the start: a
    // runner's token, a dispatcher's credentials file, or the pointer a
    // dispatcher wrote into this task worktree. Everything else waits for the
    // session whose approval it is (#491).
    if (granted() && !noIdentity()) owns = true
    // A pre-change approval was kept for the folder, not a session. OpenCode's
    // storage names the last session that worked here as that agent (the key
    // `sessionStoreKey` writes), so that is the session the approval belonged
    // to: it inherits it once, when it appears, and every other session - a
    // child's, a second window's - approves for itself. A checkout with no such
    // record has no session to attribute the approval to, so its session is
    // asked to run /office (the folder pointer still pre-fills the agent).
    if (!granted() && !noIdentity() && !resolved.worktree) {
      legacyAgent = loadAgentPointer(pointerFile) ?? loadAgentPointer(repoPointerFile)
      if (legacyAgent) {
        const stored = await bounded(ctx.storage.get(`last-session:${legacyAgent}:${locationKey}`), "storage.get").catch(() => undefined)
        if (typeof stored === "string" && stored) legacySessionID = stored
      }
    }
    startPolling()

    // The one door from a session id to an identity. A session owns one when
    // its approval is recorded against that id, when an approval finished in
    // this process for it, or when it is the session a pre-change approval
    // belonged to. A grant owns without a session and records nothing.
    const claimSession = (sessionID: string | undefined) => {
      if (!sessionID || owns || noIdentity()) return
      let agent = sessionAgent(sessionID)
      if (!agent && legacyAgent && sessionID === legacySessionID) {
        // The grandfather, recorded once and against this id alone.
        saveAgentPointer(sessionPointerPath(sessionID), legacyAgent)
        agent = legacyAgent
      }
      if (!agent) return
      if (!process.env.AGENT_OFFICE_TOKEN && !usingCredentialsOverride) {
        // The session's own approval names its own token: two agents can share
        // a folder, and this session speaks as the one it approved (as ao-pi's
        // session pointer does in a shared checkout). A token stays the truth
        // when the runner set one.
        credentialsFile = agentCredentialsPath(agent)
        auth.credentials = loadCredentials(credentialsFile)
      }
      owns = true
      ownedSessionID = sessionID
      connected = undefined
      void ensureConnected()
      startPolling()
    }

    // A session that owns nothing must not speak as whoever approved the
    // folder. Say so once per session, naming the session and the agent it
    // would have spoken as: the whole point of the ticket is that a sender line
    // is unverifiable after the fact, so the refusal has to be traceable in the
    // log too - a headless run shows no UI (#476, #491).
    const refusedSessions = new Set<string>()
    const refuseSession = (sessionID: string | undefined) => {
      if (!sessionID || owns || noIdentity() || refusedSessions.has(sessionID)) return
      const wouldSpeakAs = loadAgentPointer(pointerFile) ?? loadAgentPointer(repoPointerFile)
      // Nothing to inherit: a folder nobody connected is not this ticket.
      if (!wouldSpeakAs && !auth.credentials && !process.env.AGENT_OFFICE_TOKEN) return
      refusedSessions.add(sessionID)
      const agent = wouldSpeakAs ?? "the host's agent"
      warn("a session with no Office approval of its own will not post", { session: sessionID, wouldSpeakAs: agent })
      notify(
        sessionID,
        process.env.AGENT_OFFICE_CHILD === "1"
          ? `Child session ${sessionID} has no Office identity of its own, so it cannot post: without an approval of its own it would have spoken as ${agent}. Run /office here if this session needs to speak as an agent.`
          : `Session ${sessionID} has no Office identity of its own, so it cannot post: without an approval of its own it would have spoken as ${agent}. Run /office here to approve one.`,
      )
    }

    await ctx.command.transform((editor) => {
      editor.add({
        name: "office",
        description: "Connect to Agent Office: register an agent, check status, or reconnect",
        execute: async ({ sessionID, prompt, delivery }) => {
          claimSession(sessionID)
          const input = prompt.text.trim()
          const callback = pastedCallback(prompt.text)
          if (callback) {
            notify(sessionID, await finishWithCallback(callback))
            return
          }
          const [verb, ...rest] = input.split(/\s+/)
          const argument = rest.join(" ").trim()
          let request: string
          if (!input) {
            request = "Connect this OpenCode to Agent Office with the office_connect tool and tell me what it returned."
          } else if (verb === "status") {
            request = "Call the office_status tool and report the result in short bullet points."
          } else if (verb === "reconnect") {
            request = "Call office_connect with reconnect true, then tell me what it returned."
          } else if (verb === "register") {
            request = argument
              ? `Connect this OpenCode to Agent Office with the office_connect tool using agent_name ${JSON.stringify(argument)} and tell me what it returned.`
              : "Ask me for the agent name, then call office_connect with agent_name and tell me what it returned."
          } else {
            request = `Connect this OpenCode to Agent Office with the office_connect tool using agent_name ${JSON.stringify(input)} and tell me what it returned.`
          }
          await ctx.session.prompt({
            ...prompt,
            sessionID,
            text: request,
            delivery,
          })
        },
      })
    })

    await ctx.tool.transform((editor) => {
      editor.add({
        name: "office_connect",
        description:
          "Connect OpenCode to Agent Office. Starts an approval and returns its link at once; show the user the link. When the Office admin approves, the tokens are stored, the office MCP tools (checkin, tasks, messages, journal, decisions) are added and the session is told. Already connected: reports who you are. `reconnect` starts a fresh approval. `agent_name` pre-fills a new Office agent name. When the user pastes a http://127.0.0.1:.../callback?... URL, pass it as `callback_url` to finish the approval.",
        input: {
          type: "object",
          properties: {
            reconnect: { type: "boolean", description: "force a new OAuth approval" },
            agent_name: { type: "string", description: "Office agent name to pre-fill when approving a new agent" },
            callback_url: { type: "string", description: "the full http://127.0.0.1:.../callback?... URL from the browser when it cannot reach the loopback listener" },
          },
          additionalProperties: false,
        },
        async execute(input, toolCtx) {
          claimSession(toolCtx.sessionID)
          const { reconnect, callback_url, agent_name } = input as { reconnect?: boolean; callback_url?: string; agent_name?: string }
          if (callback_url) return { content: await finishWithCallback(pastedCallback(callback_url) ?? callback_url) }
          // A session with no identity of its own is exactly who calls this, and
          // it gets an approval link rather than "not connected". ownedBearer is
          // the check, not a flag: a grant whose credentials are not there yet
          // still approves (ao-pi's inherited-pointer case).
          if (agent_name?.trim() || reconnect || !(await ownedBearer())) {
            const approvalName = agent_name?.trim() ? agent_name.trim().slice(0, 80) : clientName
            return { content: approvalText(approvalName, await startApproval(approvalName, toolCtx.sessionID)) }
          }
          return { content: await connectedSummary() }
        },
      })
      editor.add({
        name: "office_status",
        description: "Report the Agent Office plugin's poll state: last poll, last event, resolved session, pending counts, and errors.",
        input: { type: "object", properties: {}, additionalProperties: false },
        async execute(_input, toolCtx) {
          claimSession(toolCtx.sessionID)
          return {
            content: JSON.stringify(
              {
                location: ctx.location.directory,
                canonical: ctx.location.project?.canonical,
                credentialsFile,
                tokenStorage: usingCredentialsOverride ? "override" : "agent",
                officeIdentity: noIdentity()
                  ? "none (unconnected worktree: run /office here)"
                  : owns
                    ? "connected"
                    : "none (this session has no Office approval of its own: run /office here)",
                ownsIdentity: owns,
                identitySessionID: ownedSessionID,
                sessionApproval: sessionAgent(ownedSessionID),
                approvedHere: approvedSessionID,
                legacySessionID,
                refusedSessions: [...refusedSessions],
                agentPointer: pointerFile,
                repoAgentPointer: repoPointerFile,
                clientName,
                plugin: ctx.app.version,
                mode: dispatchMode,
                inbox: inboxMode,
                pollInterval,
                lastPollAt: lastPollAt ? new Date(lastPollAt).toISOString() : undefined,
                lastPollError,
                lastEventAt: lastEventAt ? new Date(lastEventAt).toISOString() : undefined,
                lastEventType,
                lastEventDirectory,
                lastEventSessionID,
                lastSessionID,
                targetSession: targetSession(),
                busy: [...busy],
                lastPromptAt: lastPromptAt ? new Date(lastPromptAt).toISOString() : undefined,
                lastPromptError,
                lastPending,
                seen: { messages: seenMessages.size, unclaimed: seenTasks.size, assigned: seenAssigned.size },
                backlog: [...backlog],
                dispatching: [...dispatching],
                inboxSessions: Object.fromEntries(inboxSessions),
              },
              null,
              2,
            ),
          }
        },
      })
    })

    void (async () => {
      try {
        for await (const event of ctx.event.subscribe({ signal: life.signal })) {
          lastEventAt = Date.now()
          lastEventType = event.type
          const directory = (event as any).location?.directory
          lastEventDirectory = directory
          const eventSessionID = typeof (event as any).data?.sessionID === "string" ? ((event as any).data.sessionID as string) : undefined
          const deleted = event.type === "session.deleted"
          if (eventSessionID ? !deleted && !(await ownsSession(eventSessionID)) : directory && directory !== ctx.location.directory) continue
          lastEventSessionID = eventSessionID
          if (lastEventSessionID && !deleted && !officeChild(lastEventSessionID)) rememberSession(lastEventSessionID)
          // Who is speaking decides whether this process may speak at all: a
          // session with no approval of its own is told so and left with no
          // tools, never handed the folder pointer's token (#491). A deleted
          // session is bookkeeping, not a speaker.
          if (!deleted) {
            claimSession(eventSessionID)
            refuseSession(eventSessionID)
          }
          void ensureConnected()
          switch (event.type) {
            case "session.status": {
              const child = officeChild(event.data.sessionID)
              if (!child) rememberSession(event.data.sessionID)
              if (event.data.status.type === "idle") {
                stop(event.data.sessionID)
              } else start(event.data.sessionID)
              break
            }
            case "session.idle": {
              const child = officeChild(event.data.sessionID)
              if (!child) rememberSession(event.data.sessionID)
              stop(event.data.sessionID)
              break
            }
            case "session.deleted": {
              ownership.delete(event.data.sessionID)
              defaultModelNoticed.delete(event.data.sessionID)
              for (const pending of creating) if (pending.parent === event.data.sessionID) pending.deleted = true
              // A deleted inbox child is recreated on the next mail.
              for (const [child, root] of inboxSessions) if (child === event.data.sessionID || root === event.data.sessionID) inboxSessions.delete(child)
              const key = sessionStoreKey()
              if (lastSessionID === event.data.sessionID) {
                lastSessionID = undefined
                if (key) void ctx.storage.remove(key).catch(() => {})
              }
              stop(event.data.sessionID)
              break
            }
          }
        }
      } catch (err) {
        if (!life.signal.aborted) warn("event subscription stopped", { error: String(err) })
      }
    })()

    return async () => {
      life.abort()
      clearInterval(heartbeat)
      clearInterval(pollTimer)
      clearTimeout(refreshTimer)
      if (busy.size === 0) return
      busy.clear()
      await report("idle")
    }
  },
})

