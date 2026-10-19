// Agent Office plugin for OpenCode (https://opencode.ai/docs/plugins/).
// Keeps this OpenCode's Office status honest: `working` while any session
// runs, `idle` once every session is quiet, with a heartbeat in between so
// the dashboard shows it online. It never overrides `blocked`. The Office
// tools themselves come from the Office MCP connector (see README).
//
//   AGENT_OFFICE_URL         dashboard origin, default https://office.byma.tech
//   AGENT_OFFICE_TOKEN       API token (optional when the token file exists)
//   AGENT_OFFICE_TOKEN_FILE  default ~/.config/agent-office/token
//
// The token is read once at startup and never logged.

import type { Plugin } from "@opencode-ai/plugin"
import { readFileSync } from "node:fs"
import { homedir } from "node:os"
import { basename, join } from "node:path"

const HEARTBEAT_MS = 90_000

function readToken(): string | undefined {
  if (process.env.AGENT_OFFICE_TOKEN) return process.env.AGENT_OFFICE_TOKEN
  const file = process.env.AGENT_OFFICE_TOKEN_FILE ?? join(homedir(), ".config", "agent-office", "token")
  try {
    return readFileSync(file, "utf8").trim() || undefined
  } catch {
    return undefined
  }
}

export const AgentOfficePlugin: Plugin = async ({ client, worktree }) => {
  const base = (process.env.AGENT_OFFICE_URL ?? "https://office.byma.tech").replace(/\/+$/, "")
  const token = readToken()
  const log = (level: "warn" | "info", message: string, extra?: Record<string, unknown>) =>
    client.app.log({ body: { service: "agent-office", level, message, extra } }).catch(() => {})

  if (!token) {
    await log("warn", "no token: set AGENT_OFFICE_TOKEN or write the token file; status is not reported")
    return {}
  }

  const api = async (method: "GET" | "PUT", path: string, body?: unknown) => {
    const res = await fetch(`${base}/api/v1/agent${path}`, {
      method,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: AbortSignal.timeout(10_000),
    })
    if (!res.ok) throw new Error(`${method} ${path}: HTTP ${res.status}`)
    return res.json()
  }

  const note = `opencode: ${basename(worktree)}`
  const busy = new Set<string>()
  let heartbeat: ReturnType<typeof setInterval> | undefined
  let queue = Promise.resolve()

  // Reports run one after another so a short turn cannot land idle before working.
  const report = (status: "working" | "idle") => {
    queue = queue.then(async () => {
      try {
        if (status === "idle" && (await api("GET", "/me")).work_status === "blocked") return
        await api("PUT", "/status", { status, note })
      } catch (err) {
        await log("warn", `could not report ${status}`, { error: String(err) })
      }
    })
    return queue
  }

  const start = (sessionID: string) => {
    if (busy.size === 0) {
      heartbeat = setInterval(() => api("GET", "/me").catch(() => {}), HEARTBEAT_MS)
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
    event: async ({ event }) => {
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
      if (busy.size === 0) return
      busy.clear()
      await report("idle")
    },
  }
}
