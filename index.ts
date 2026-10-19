// Agent Office plugin for OpenCode (https://opencode.ai/docs/plugins/).
//
//   /office            connects: OAuth in the browser, Azri approves, tokens are
//                      stored, and the Office MCP tools appear in OpenCode.
//   while running      `working` while any session runs, a heartbeat every
//                      90 s, `idle` once every session is quiet; `blocked` is
//                      never overridden.
//
//   AGENT_OFFICE_URL                   dashboard origin, default https://office.byma.tech
//   AGENT_OFFICE_TOKEN                 an agent API token instead of OAuth (optional)
//   AGENT_OFFICE_OPENCODE_CREDENTIALS  default ~/.config/agent-office/opencode.json
//   AGENT_OFFICE_CALLBACK_PORT         loopback port for the OAuth callback, default 41997
//
// Tokens are never logged.

import { tool, type Plugin } from "@opencode-ai/plugin"
import { spawn } from "node:child_process"
import { createHash, randomBytes } from "node:crypto"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { createServer } from "node:http"
import { homedir } from "node:os"
import { basename, dirname, join } from "node:path"

const OFFICE = (process.env.AGENT_OFFICE_URL ?? "https://office.byma.tech").replace(/\/+$/, "")
const CREDENTIALS = process.env.AGENT_OFFICE_OPENCODE_CREDENTIALS ?? join(homedir(), ".config", "agent-office", "opencode.json")
const CALLBACK_PORT = Number(process.env.AGENT_OFFICE_CALLBACK_PORT ?? 41997)
const CALLBACK = `http://127.0.0.1:${CALLBACK_PORT}/callback`
const SCOPES = "office:read office:write office:events"
const HEARTBEAT_MS = 90_000
const REFRESH_EARLY_MS = 5 * 60_000
const APPROVAL_TIMEOUT_MS = 5 * 60_000
const MCP_PROTOCOL = "2026-07-28"

type Credentials = { office: string; client_id: string; access_token: string; refresh_token: string; expires_at: number }

function loadCredentials(): Credentials | undefined {
  try {
    const c = JSON.parse(readFileSync(CREDENTIALS, "utf8")) as Credentials
    return c.office === OFFICE ? c : undefined
  } catch {
    return undefined
  }
}

function saveCredentials(c: Credentials) {
  mkdirSync(dirname(CREDENTIALS), { recursive: true, mode: 0o700 })
  writeFileSync(CREDENTIALS, JSON.stringify(c, null, 2) + "\n", { mode: 0o600 })
}

async function tokenRequest(client_id: string, form: Record<string, string>): Promise<Credentials> {
  const res = await fetch(`${OFFICE}/oauth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id, resource: `${OFFICE}/mcp`, ...form }),
    signal: AbortSignal.timeout(15_000),
  })
  const body = (await res.json()) as { access_token: string; refresh_token: string; expires_in: number; error?: string; error_description?: string }
  if (!res.ok) throw new Error(`Office token request failed: ${body.error_description ?? body.error ?? res.status}`)
  return { office: OFFICE, client_id, access_token: body.access_token, refresh_token: body.refresh_token, expires_at: Date.now() + body.expires_in * 1000 }
}

async function registerClient(): Promise<string> {
  const res = await fetch(`${OFFICE}/oauth/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      client_name: "OpenCode",
      redirect_uris: [CALLBACK],
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

// Waits on the loopback callback for the authorization code of this state.
function awaitCode(state: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const server = createServer((req, res) => {
      const url = new URL(req.url ?? "/", CALLBACK)
      if (url.pathname !== "/callback" || url.searchParams.get("state") !== state) {
        res.writeHead(404).end()
        return
      }
      const code = url.searchParams.get("code")
      const error = url.searchParams.get("error_description") ?? url.searchParams.get("error")
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" })
      res.end(`<!doctype html><title>Agent Office</title><p>${code ? "OpenCode is connected to Agent Office. You can close this tab." : `Not connected: ${error}`}</p>`)
      server.close()
      clearTimeout(timer)
      code ? resolve(code) : reject(new Error(`Office authorization failed: ${error}`))
    })
    const timer = setTimeout(() => {
      server.close()
      reject(new Error("timed out waiting for the approval in the browser"))
    }, APPROVAL_TIMEOUT_MS)
    server.on("error", (err) => {
      clearTimeout(timer)
      reject(new Error(`cannot listen on ${CALLBACK}: ${err.message}`))
    })
    server.listen(CALLBACK_PORT, "127.0.0.1")
  })
}

async function authorize(onUrl: (url: string) => void): Promise<Credentials> {
  const client_id = await registerClient()
  const verifier = randomBytes(32).toString("base64url")
  const state = randomBytes(16).toString("base64url")
  const url = new URL(`${OFFICE}/oauth/authorize`)
  url.search = new URLSearchParams({
    response_type: "code",
    client_id,
    redirect_uri: CALLBACK,
    scope: SCOPES,
    state,
    code_challenge: createHash("sha256").update(verifier).digest("base64url"),
    code_challenge_method: "S256",
    resource: `${OFFICE}/mcp`,
  }).toString()
  const pending = awaitCode(state)
  onUrl(url.href)
  openBrowser(url.href)
  const code = await pending
  return tokenRequest(client_id, { grant_type: "authorization_code", code, redirect_uri: CALLBACK, code_verifier: verifier })
}

export const AgentOfficePlugin: Plugin = async ({ client, worktree }) => {
  const log = async (level: "warn" | "info", message: string, extra?: Record<string, unknown>) => {
    await client.app.log({ body: { service: "agent-office", level, message, extra } }).catch(() => {})
  }

  let credentials = loadCredentials()

  // The bearer for every Office call: the API token from the environment, or
  // the OAuth access token, refreshed shortly before it expires.
  const bearer = async (): Promise<string | undefined> => {
    if (process.env.AGENT_OFFICE_TOKEN) return process.env.AGENT_OFFICE_TOKEN
    if (!credentials) return undefined
    if (credentials.expires_at - Date.now() < REFRESH_EARLY_MS) {
      credentials = await tokenRequest(credentials.client_id, { grant_type: "refresh_token", refresh_token: credentials.refresh_token })
      saveCredentials(credentials)
    }
    return credentials.access_token
  }

  const call = async (name: string, args: Record<string, unknown> = {}) => {
    const token = await bearer()
    if (!token) throw new Error("not connected to Agent Office: run /office")
    const res = await fetch(`${OFFICE}/mcp`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json, text/event-stream",
        "MCP-Protocol-Version": MCP_PROTOCOL,
        "Mcp-Method": "tools/call",
      },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "tools/call",
        params: { name, arguments: args, _meta: { "io.modelcontextprotocol/protocolVersion": MCP_PROTOCOL } },
      }),
      signal: AbortSignal.timeout(15_000),
    })
    if (!res.ok) throw new Error(`Office ${name}: HTTP ${res.status}`)
    const body = (await res.json()) as { result?: { isError?: boolean; structuredContent?: Record<string, any>; content?: { text?: string }[] }; error?: { message: string } }
    if (body.error) throw new Error(`Office ${name}: ${body.error.message}`)
    if (body.result?.isError) throw new Error(`Office ${name}: ${body.result.content?.[0]?.text ?? "tool error"}`)
    return body.result?.structuredContent ?? {}
  }

  // Hands OpenCode the office MCP server with the current access token, and
  // again whenever the token is refreshed.
  let refreshTimer: ReturnType<typeof setTimeout> | undefined
  const connectMcp = async () => {
    const token = await bearer()
    if (!token) return
    await client.mcp.add({
      body: { name: "office", config: { type: "remote", url: `${OFFICE}/mcp`, oauth: false, headers: { Authorization: `Bearer ${token}` } } },
    })
    clearTimeout(refreshTimer)
    if (credentials && !process.env.AGENT_OFFICE_TOKEN) {
      const delay = Math.max(credentials.expires_at - Date.now() - REFRESH_EARLY_MS, 10_000)
      refreshTimer = setTimeout(() => connectMcp().catch((err) => log("warn", "could not refresh the Office connection", { error: String(err) })), delay)
      refreshTimer.unref()
    }
  }
  let connected: Promise<void> | undefined
  const ensureConnected = () => (connected ??= connectMcp().catch((err) => log("warn", "could not connect the Office MCP server", { error: String(err) })))

  const note = `opencode: ${basename(worktree)}`
  const busy = new Set<string>()
  let heartbeat: ReturnType<typeof setInterval> | undefined
  let queue = Promise.resolve()

  // Reports run one after another so a short turn cannot land idle before working.
  const report = (status: "working" | "idle") => {
    queue = queue.then(async () => {
      try {
        if (!(await bearer())) return
        if (status === "idle" && (await call("checkin")).agent?.work_status === "blocked") return
        await call("checkin", { status, note })
      } catch (err) {
        await log("warn", `could not report ${status}`, { error: String(err) })
      }
    })
    return queue
  }

  const start = (sessionID: string) => {
    if (busy.size === 0) {
      heartbeat = setInterval(() => call("checkin").catch(() => {}), HEARTBEAT_MS)
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

  return {
    config: async (cfg) => {
      ;(cfg.command ??= {}).office ??= {
        description: "Connect to Agent Office",
        template: "Connect this OpenCode to Agent Office with the office_connect tool and tell me what it returned. $ARGUMENTS",
      }
    },
    tool: {
      office_connect: tool({
        description:
          "Connect OpenCode to Agent Office. Opens the browser for the Office admin to approve, stores the tokens and adds the office MCP tools (checkin, tasks, messages, journal, decisions). Already connected: reports who you are. `reconnect` starts a fresh approval.",
        args: { reconnect: tool.schema.boolean().optional().describe("force a new OAuth approval") },
        async execute(args, ctx) {
          if (args.reconnect || !(await bearer())) {
            credentials = await authorize((url) => ctx.metadata({ title: "Approve OpenCode in the browser", metadata: { url } }))
            saveCredentials(credentials)
            connected = undefined
          }
          await ensureConnected()
          const me = await call("checkin")
          const pending = me.pending ?? {}
          return `Connected to Agent Office as @${me.agent?.slug} (${me.agent?.name}). Open tasks: ${pending.open_tasks?.length ?? 0}, unclaimed team tasks: ${pending.unclaimed_team_tasks?.length ?? 0}, unread messages: ${pending.unread_messages?.length ?? 0}. The office MCP tools are available; call checkin first.`
        },
      }),
    },
    event: async ({ event }) => {
      void ensureConnected()
      switch (event.type) {
        case "session.status":
          if (event.properties.status.type === "idle") stop(event.properties.sessionID)
          else start(event.properties.sessionID)
          break
        case "session.idle":
          stop(event.properties.sessionID)
          break
        case "session.deleted":
          stop(event.properties.info.id)
          break
      }
    },
    dispose: async () => {
      clearInterval(heartbeat)
      clearTimeout(refreshTimer)
      if (busy.size === 0) return
      busy.clear()
      await report("idle")
    },
  }
}
