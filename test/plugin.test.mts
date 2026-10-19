import assert from "node:assert/strict"
import { test } from "node:test"

process.env.AGENT_OFFICE_TOKEN = "secret-token"
process.env.AGENT_OFFICE_URL = "https://office.example/"

const calls: string[] = []
const logs: { level: string; message: string }[] = []
let me = { work_status: "working" }
globalThis.fetch = (async (url: string, init: RequestInit) => {
  assert.equal((init.headers as Record<string, string>).Authorization, "Bearer secret-token")
  calls.push(`${init.method} ${url}${init.body ? " " + init.body : ""}`)
  return { ok: true, status: 200, json: async () => me }
}) as unknown as typeof fetch
const client = { app: { log: async (x: { body: { level: string; message: string } }) => void logs.push(x.body) } }
const tick = () => new Promise((r) => setTimeout(r, 5))

const { AgentOfficePlugin } = await import("../index.ts")
const hooks = (await AgentOfficePlugin({ client, worktree: "/home/x/agent-office" } as never)) as {
  event: (input: { event: unknown }) => Promise<void>
  dispose: () => Promise<void>
}
const ev = (type: string, properties: unknown) => hooks.event({ event: { type, properties } })
const working = 'PUT https://office.example/api/v1/agent/status {"status":"working","note":"opencode: agent-office"}'
const idle = 'PUT https://office.example/api/v1/agent/status {"status":"idle","note":"opencode: agent-office"}'
const meCall = "GET https://office.example/api/v1/agent/me"

test("working once while any session is busy, idle when all are quiet", async () => {
  await ev("session.status", { sessionID: "a", status: { type: "busy" } })
  await ev("session.status", { sessionID: "a", status: { type: "busy" } })
  await ev("session.status", { sessionID: "b", status: { type: "retry", attempt: 1 } })
  await tick()
  assert.deepEqual(calls, [working])
  await ev("session.idle", { sessionID: "a" })
  await tick()
  assert.deepEqual(calls, [working], "b still busy")
  await ev("session.deleted", { info: { id: "b" } })
  await tick()
  assert.deepEqual(calls, [working, meCall, idle])
})

test("never overrides blocked", async () => {
  calls.length = 0
  me = { work_status: "blocked" }
  await ev("session.status", { sessionID: "c", status: { type: "busy" } })
  await ev("session.status", { sessionID: "c", status: { type: "idle" } })
  await tick()
  assert.deepEqual(calls, [working, meCall])
  me = { work_status: "working" }
})

test("ignores sessions it never saw", async () => {
  calls.length = 0
  await ev("session.idle", { sessionID: "unknown" })
  await tick()
  assert.deepEqual(calls, [])
})

test("logs a failure without the token and keeps going", async () => {
  const real = globalThis.fetch
  globalThis.fetch = (async () => ({ ok: false, status: 401, json: async () => ({}) })) as unknown as typeof fetch
  await ev("session.status", { sessionID: "d", status: { type: "busy" } })
  await tick()
  assert.equal(logs.at(-1)?.level, "warn")
  assert.match(logs.at(-1)?.message ?? "", /working/)
  assert.ok(!JSON.stringify(logs).includes("secret-token"))
  globalThis.fetch = real
  await hooks.dispose()
})

test("no token: no hooks, one warning", async () => {
  delete process.env.AGENT_OFFICE_TOKEN
  process.env.AGENT_OFFICE_TOKEN_FILE = "/nonexistent"
  assert.deepEqual(await AgentOfficePlugin({ client, worktree: "/x" } as never), {})
  assert.match(logs.at(-1)?.message ?? "", /no token/)
})
