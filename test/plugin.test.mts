import assert from "node:assert/strict"
import { mkdtempSync, readFileSync, statSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { after, before, test } from "node:test"
import { fakeOffice } from "./fake-office.mts"

const office = fakeOffice()
const dir = mkdtempSync(join(tmpdir(), "ao-opencode-"))
const credentials = join(dir, "opencode.json")
process.env.AGENT_OFFICE_OPENCODE_CREDENTIALS = credentials
process.env.AGENT_OFFICE_CALLBACK_PORT = "45123"
delete process.env.AGENT_OFFICE_TOKEN

const added: any[] = []
const logs: { level: string; message: string }[] = []
const client = {
  app: { log: async (x: { body: { level: string; message: string } }) => void logs.push(x.body) },
  mcp: { add: async (x: { body: any }) => void added.push(x.body) },
}
const tick = () => new Promise((r) => setTimeout(r, 20))
const mcpCalls = () => office.calls.filter((c) => c.path === "/mcp").map((c) => `${c.auth} ${c.body.params.name} ${JSON.stringify(c.body.params.arguments)}`)

let hooks: any
before(async () => {
  process.env.AGENT_OFFICE_URL = (await office.listen()) + "/"
  const { AgentOfficePlugin } = await import("../index.ts")
  hooks = await AgentOfficePlugin({ client, worktree: "/home/x/agent-office" } as never)
})
after(() => office.close())

test("registers the /office command", async () => {
  const cfg: any = {}
  await hooks.config(cfg)
  assert.match(cfg.command.office.template, /office_connect/)
})

test("not connected: events report nothing and the tools are not added", async () => {
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "a", status: { type: "busy" } } } })
  await hooks.event({ event: { type: "session.idle", properties: { sessionID: "a" } } })
  await tick()
  assert.deepEqual(added, [])
  assert.deepEqual(office.calls, [])
})

test("/office: OAuth with PKCE, stored tokens, MCP server added", async () => {
  let authorizeUrl = ""
  const result = hooks.tool.office_connect.execute({}, { metadata: (m: any) => (authorizeUrl = m.metadata.url) })
  while (!authorizeUrl) await tick()
  // Azri's browser: the Office approves and sends it back to the plugin's callback.
  const back = await fetch(authorizeUrl, { redirect: "manual" })
  assert.equal(back.status, 302)
  const landing = await fetch(back.headers.get("location")!)
  assert.match(await landing.text(), /connected/)
  assert.match(await result, /^Connected to Agent Office as @opencode \(OpenCode\)\. Open tasks: 1, unclaimed team tasks: 0, unread messages: 2\./)

  const register = office.calls.find((c) => c.path === "/oauth/register")!
  const token = office.calls.find((c) => c.path === "/oauth/token")!
  assert.deepEqual(register.body.redirect_uris, ["http://127.0.0.1:45123/callback"])
  assert.equal(token.body.grant_type, "authorization_code")
  assert.equal(token.body.code, "code-1")
  assert.equal(token.body.client_id, "client-1")
  assert.ok(token.body.code_verifier.length >= 43)
  assert.equal(token.body.resource, process.env.AGENT_OFFICE_URL!.replace(/\/$/, "") + "/mcp")

  const saved = JSON.parse(readFileSync(credentials, "utf8"))
  assert.equal(saved.access_token, "access-1")
  assert.equal(statSync(credentials).mode & 0o777, 0o600)
  assert.deepEqual(added, [{ name: "office", config: { type: "remote", url: saved.office + "/mcp", oauth: false, headers: { Authorization: "Bearer access-1" } } }])
  assert.deepEqual(mcpCalls(), ["Bearer access-1 checkin {}"])
  office.calls.length = 0
})

test("working while any session is busy, idle when all are quiet", async () => {
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "a", status: { type: "busy" } } } })
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "b", status: { type: "retry", attempt: 1 } } } })
  await hooks.event({ event: { type: "session.idle", properties: { sessionID: "a" } } })
  await tick()
  assert.deepEqual(mcpCalls(), ['Bearer access-1 checkin {"status":"working","note":"opencode: agent-office"}'])
  await hooks.event({ event: { type: "session.deleted", properties: { info: { id: "b" } } } })
  await tick()
  assert.deepEqual(mcpCalls().slice(1), ["Bearer access-1 checkin {}", 'Bearer access-1 checkin {"status":"idle","note":"opencode: agent-office"}'])
  office.calls.length = 0
})

test("never overrides blocked", async () => {
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "c", status: { type: "busy" } } } })
  await tick()
  office.state.workStatus = "blocked" // the model asked Azri something through `ask`
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "c", status: { type: "idle" } } } })
  await tick()
  assert.deepEqual(mcpCalls(), ['Bearer access-1 checkin {"status":"working","note":"opencode: agent-office"}', "Bearer access-1 checkin {}"])
  office.state.workStatus = "idle"
  office.calls.length = 0
  await hooks.dispose()
})

test("a stored token near expiry is refreshed before use and handed to OpenCode again", async () => {
  const saved = JSON.parse(readFileSync(credentials, "utf8"))
  writeFileSync(credentials, JSON.stringify({ ...saved, expires_at: Date.now() + 60_000 }))
  added.length = 0
  const { AgentOfficePlugin } = await import("../index.ts")
  const again: any = await AgentOfficePlugin({ client, worktree: "/x" } as never)
  await again.event({ event: { type: "session.created", properties: { info: { id: "d" } } } })
  await tick()
  const refresh = office.calls.find((c) => c.path === "/oauth/token")!
  assert.equal(refresh.body.grant_type, "refresh_token")
  assert.equal(refresh.body.refresh_token, "refresh-1")
  assert.equal(added[0].config.headers.Authorization, "Bearer access-2")
  assert.equal(JSON.parse(readFileSync(credentials, "utf8")).refresh_token, "refresh-2")
  await again.dispose()
})

test("logs failures without the token", async () => {
  await office.close()
  await hooks.event({ event: { type: "session.status", properties: { sessionID: "e", status: { type: "busy" } } } })
  for (let i = 0; i < 100 && logs.at(-1)?.level !== "warn"; i++) await tick()
  assert.equal(logs.at(-1)?.level, "warn")
  assert.match(logs.at(-1)?.message ?? "", /working/)
  assert.ok(!JSON.stringify(logs).includes("access-"))
  await hooks.dispose()
})
