// Agent Office plugin for OpenCode (https://opencode.ai/docs/plugins/).
//
//   /office [name]     connects: OAuth in the browser, Azri approves, the token
//                      is stored under the agent name, and the Office MCP tools
//                      appear in OpenCode. The name pre-fills a new agent.
//   while running      `working` while any session runs, a heartbeat every
//                      90 s, `idle` once every session is quiet; `blocked` is
//                      never overridden. New mail runs in an inbox child of
//                      the chat's session. New assigned tasks go to a supervisor
//                      and its worker in a worktree, through Herdr or a child
//                      session, up to two tasks at a time.
//
//   AGENT_OFFICE_URL                   dashboard origin, default https://office.byma.tech
//   AGENT_OFFICE_TOKEN                 an agent API token instead of OAuth (optional)
//   AGENT_OFFICE_OPENCODE_CREDENTIALS  exact credentials file, overrides agent-scoped storage
//   AGENT_OFFICE_CALLBACK_PORT         loopback port for the OAuth callback, default 41997
//   AGENT_OFFICE_POLL_MS               inbox poll interval, default 60000
//   AGENT_OFFICE_AUTODISPATCH          on (default), off, herdr, or session
//   AGENT_OFFICE_CLIENT_NAME           OAuth client name; defaults to folder@host
//
// Tokens are never logged.

import { Plugin } from "@opencode/plugin"
import { spawn } from "node:child_process"
import { createHash, randomBytes } from "node:crypto"
import { existsSync } from "node:fs"
import { createServer } from "node:http"
import { hostname } from "node:os"
import { basename } from "node:path"
import { fileURLToPath } from "node:url"
import {
  type Credentials,
  acquireDispatch,
  addWorktree,
  agentCredentialsPath,
  agentStateDir,
  CREDENTIALS_OVERRIDE,
  claimNotice,
  claimTask,
  createBearer,
  createClient,
  cronLine,
  DISPATCH_STALE_MS,
  fromOthers,
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
  summary,
  supervisorPrompt,
  taskID,
  taskShort,
  tokenRequest,
  workspaceKey,
} from "./lib/office.ts"

export { agentCredentialsPath, agentPointerPath, credentialsPath } from "./lib/office.ts"

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

function jsonResult(output: string): any {
  try {
    const parsed = JSON.parse(output)
    return parsed?.result ?? parsed
  } catch {
    return undefined
  }
}

function agentState(output: string): string | undefined {
  const result = jsonResult(output)
  return result?.agent?.agent_status ?? result?.agent_status ?? result?.state
}

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

function openBrowser(url: string) {
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
    let credentialsFile = resolved.credentialsFile
    const usingCredentialsOverride = Boolean(CREDENTIALS_OVERRIDE)
    const auth = createBearer(() => credentialsFile, resolved.credentials)
    const adoptAgent = (slug?: string) => {
      if (usingCredentialsOverride || !slug) return
      const target = agentCredentialsPath(slug)
      if (auth.credentials) saveCredentials(target, auth.credentials)
      credentialsFile = target
      saveAgentPointer(pointerFile, slug)
    }
    const dispatchDir = () => agentStateDir(legacyCredentialsFile, "dispatch", ownAgent?.slug ?? locationKey)
    const noticeDir = () => agentStateDir(legacyCredentialsFile, "notices", ownAgent?.slug ?? locationKey)
    const configuredDispatch = String(ctx.options.autoDispatch ?? process.env.AGENT_OFFICE_AUTODISPATCH ?? "on").toLowerCase()
    const dispatchMode = configuredDispatch === "off" || configuredDispatch === "false" ? "off" : configuredDispatch === "herdr" ? "herdr" : configuredDispatch === "session" ? "session" : process.env.HERDR_ENV === "1" ? "herdr" : "session"
    const bearer = auth.bearer
    const { call, checkin } = createClient(bearer)

    // Hands OpenCode the office MCP server with the current access token, and
    // again whenever the token is refreshed. The transform captures the token
    // and `reload()` replays it after every refresh.
    let token: string | undefined
    await ctx.mcp.transform((editor) => {
      if (!token) return
      editor.set("office", { type: "remote", url: `${OFFICE}/mcp`, oauth: false, headers: { Authorization: `Bearer ${token}` } })
    })
    let refreshTimer: ReturnType<typeof setTimeout> | undefined
    const connectMcp = async () => {
      token = await bearer()
      if (!token) return
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
          if (!(await bearer())) return
          if (status === "idle" && (await checkin()).agent?.work_status === "blocked") return
          await checkin({ status, note })
        } catch (err) {
          warn(`could not report ${status}`, { error: String(err) })
        }
      })
      return queue
    }

    const start = (sessionID: string) => {
      if (busy.size === 0) {
        heartbeat = setInterval(() => checkin().catch(() => {}), HEARTBEAT_MS)
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
    // Children this plugin started count as busy work, but never as a target.
    const officeChild = (sessionID: string) => dispatchedSessions.has(sessionID) || inboxSessions.has(sessionID)
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
        const child = await bounded(
          ctx.session.create({
            parentID: parent,
            title,
            location: location ?? { directory: info.location.directory },
            ...(info.model ? { model: info.model } : {}),
          }),
          "session.create",
        )
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
    const deliverInbox = async (parent: string, text: string) => {
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
        await bounded(ctx.session.prompt({ sessionID, text, delivery: "queue" }), "session.prompt")
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

    const connectedSummary = async () => {
      await ensureConnected()
      const me = await checkin()
      adoptAgent(me.agent?.slug)
      ownAgent = me.agent
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
          connected = undefined
          return connectedSummary()
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
    const dispatchedSessions = new Map<string, { taskID: string; parent: string; title: string }>()

    const settleChild = (sessionID: string) => {
      const child = dispatchedSessions.get(sessionID)
      if (!child) return
      dispatchedSessions.delete(sessionID)
      releaseDispatch(dispatchDir(), child.taskID)
      dispatching.delete(child.taskID)
      notify(child.parent, `Office task ${child.taskID} supervisor session ${sessionID} is idle. Read the session; the supervisor opens the PR and updates the task.`)
      pumpDispatch()
    }

    const repo = ctx.location.project?.canonical ?? ctx.location.directory

    const dispatchSession = async (id: string, task: any, title: string, parent: string) => {
      const { path, branch } = await addWorktree(repo, id)
      const child = await createChild(parent, `Office: ${title}`.slice(0, 120), { directory: path })
      dispatchedSessions.set(child.id, { taskID: id, parent, title })
      try {
        await ctx.session.prompt({ sessionID: child.id, text: supervisorPrompt(id, task, path, branch) })
      } catch (err) {
        dispatchedSessions.delete(child.id)
        void ctx.session.remove({ sessionID: child.id }).catch(() => {})
        throw err
      }
      notify(parent, `Dispatched Office task ${id} to supervisor session ${child.id} (${branch}) at ${path}.`)
    }

    const dispatchHerdr = async (id: string, task: any, title: string, parent: string) => {
      const pseudo = `herdr:${id}`
      start(pseudo)
      try {
        const short = taskShort(id)
        const branch = `office/${short}`
        const created = await run("herdr", ["worktree", "create", "--cwd", repo, "--branch", branch, "--label", title.slice(0, 80), "--no-focus"], 120_000)
        if (created.code !== 0) throw new Error(`herdr worktree create failed: ${created.stderr || created.stdout}`)
        const info = jsonResult(created.stdout)
        const pane = info?.root_pane?.pane_id
        if (!pane) throw new Error("herdr worktree create returned no root pane")
        const name = `office-${short}`
        const started = await run("herdr", ["agent", "start", name, "--kind", "opencode", "--pane", pane, "--timeout", "120000"], 150_000)
        if (started.code !== 0) throw new Error(`herdr agent start failed: ${started.stderr || started.stdout}`)
        const worktree = info?.worktree?.path ?? "of this pane"
        const prompted = await run("herdr", ["agent", "prompt", name, supervisorPrompt(id, task, worktree, branch), "--wait", "--timeout", String(DISPATCH_STALE_MS)], DISPATCH_STALE_MS + 60_000)
        const state = agentState(prompted.stdout) ?? "unknown"
        notify(parent, `Office task ${id} dispatched to Herdr supervisor ${name} (${branch}) at ${info?.worktree?.path ?? "its worktree"}. State: ${state}.`)
      } finally {
        stop(pseudo)
      }
    }

    const dispatchTask = async (id: string) => {
      let held = false
      const parent = await targetSessionAsync()
      try {
        if (!parent) return
        const detail = await call("get_task", { task_id: id }).catch(() => undefined)
        const task = detail?.task ?? detail
        const title = String(task?.title ?? id)
        await claimTask(call, id)
        if (dispatchMode === "herdr") {
          await dispatchHerdr(id, task, title, parent)
        } else {
          await dispatchSession(id, task, title, parent)
          held = true
        }
      } catch (err) {
        warn("could not dispatch Office task", { task: id, error: String(err) })
        notify(parent, `Could not dispatch Office task ${id}: ${String(err)}`)
      } finally {
        if (!held) {
          releaseDispatch(dispatchDir(), id)
          dispatching.delete(id)
          pumpDispatch()
        }
      }
    }

    // One pump at a time, including while it looks up the remembered session.
    let pumping = false
    const pumpDispatch = () => {
      if (life.signal.aborted || dispatchMode === "off" || pumping) return
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

      if (freshMessages.length > 0) {
        const ids = freshMessages.map((item) => itemKey(item, "message"))
        if (!claimNotice(noticeDir(), ids)) {
          for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
        } else {
          const text = `${INBOX_SESSION}\n${inboxPrompt(freshMessages)}`
          try {
            await deliverInbox(sessionID, text)
            lastPromptAt = Date.now()
            lastPromptError = undefined
            for (const item of freshMessages) seenMessages.add(itemKey(item, "message"))
          } catch (err) {
            releaseNotice(noticeDir(), ids)
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
        if (!(await bearer())) return
        const me = await checkin()
        adoptAgent(me.agent?.slug)
        ownAgent = me.agent
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
    if (!onTaskBranch(ctx.location.directory) && !inHerdrWorktree(ctx.location.directory)) {
      pollTimer = setInterval(() => void poll(), pollInterval)
      pollTimer.unref()
      void poll()
    }

    await ctx.command.transform((editor) => {
      editor.add({
        name: "office",
        description: "Connect to Agent Office: register an agent, check status, or reconnect",
        execute: async ({ sessionID, prompt, delivery }) => {
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
          const { reconnect, callback_url, agent_name } = input as { reconnect?: boolean; callback_url?: string; agent_name?: string }
          if (callback_url) return { content: await finishWithCallback(pastedCallback(callback_url) ?? callback_url) }
          if (agent_name?.trim() || reconnect || !(await bearer())) {
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
        async execute() {
          return {
            content: JSON.stringify(
              {
                location: ctx.location.directory,
                canonical: ctx.location.project?.canonical,
                credentialsFile,
                tokenStorage: usingCredentialsOverride ? "override" : "agent",
                agentPointer: pointerFile,
                repoAgentPointer: repoPointerFile,
                clientName,
                plugin: ctx.app.version,
                mode: dispatchMode,
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
          void ensureConnected()
          switch (event.type) {
            case "session.status": {
              const child = officeChild(event.data.sessionID)
              if (!child) rememberSession(event.data.sessionID)
              if (event.data.status.type === "idle") {
                stop(event.data.sessionID)
                settleChild(event.data.sessionID)
              } else start(event.data.sessionID)
              break
            }
            case "session.idle": {
              const child = officeChild(event.data.sessionID)
              if (!child) rememberSession(event.data.sessionID)
              stop(event.data.sessionID)
              settleChild(event.data.sessionID)
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
              settleChild(event.data.sessionID)
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

