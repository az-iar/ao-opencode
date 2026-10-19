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

const tick = () => new Promise((r) => setTimeout(r, 20))
const until = async (check: () => boolean, what: string) => {
  for (let i = 0; i < 100; i++) {
    if (check()) return
    await tick()
  }
  assert.fail(`timed out waiting for ${what}`)
}
const mcpCalls = () => office.calls.filter((c) => c.path === "/mcp").map((c) => `${c.auth} ${c.body.params.name} ${JSON.stringify(c.body.params.arguments)}`)

// The V2 plugin context, reduced to the domains this plugin uses. Transforms
// run immediately and again on reload, like the real registries.
function makeContext(directory = "/home/x/agent-office", options: Record<string, unknown> = {}) {
  const tools: any[] = []
  const commands: any[] = []
  const prompts: any[] = []
  const synthetics: any[] = []
  const servers = new Map<string, any>()
  const transforms: ((editor: any) => void)[] = []
  const queue: any[] = []
  let wake: (() => void) | undefined

  const mcpEditor = {
    list: () => [...servers.entries()],
    get: (name: string) => servers.get(name),
    set: (name: string, config: any) => servers.set(name, config),
    update: (name: string, update: (config: any) => void) => update(servers.get(name)),
    remove: (name: string) => servers.delete(name),
  }

  const ctx: any = {
    app: { name: "opencode", version: "2.0.11", channel: "latest" },
    location: { directory, project: { id: "project", directory, canonical: directory } },
    options,
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
        callback(mcpEditor)
        return { dispose: async () => {} }
      },
      reload: async () => {
        for (const callback of transforms) callback(mcpEditor)
      },
    },
    session: {
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
    tools,
    commands,
    prompts,
    synthetics,
    servers,
    push: (event: any) => {
      queue.push(event)
      wake?.()
      wake = undefined
    },
  }
}

let plugin: any
let harness: ReturnType<typeof makeContext>
let cleanupFirst: (() => Promise<void>) | undefined
const warnings: string[] = []
const originalWarn = console.warn

before(async () => {
  process.env.AGENT_OFFICE_URL = (await office.listen()) + "/"
  console.warn = (...args: unknown[]) => void warnings.push(args.join(" "))
  const { default: definition } = await import("../index.ts")
  plugin = definition
  harness = makeContext()
  cleanupFirst = await plugin.setup(harness.ctx)
})
after(async () => {
  await cleanupFirst?.()
  console.warn = originalWarn
  office.close()
})

test("defines the agent-office plugin", () => {
  assert.equal(plugin.id, "agent-office")
})

test("registers the /office command and the office_connect tool", async () => {
  const command = harness.commands.find((c) => c.name === "office")
  assert.ok(command)
  await command.execute({ sessionID: "s1", prompt: { text: "please" }, delivery: "steer" })
  assert.equal(harness.prompts.at(-1).sessionID, "s1")
  assert.match(harness.prompts.at(-1).text, /^Connect this OpenCode to Agent Office with the office_connect tool and tell me what it returned\. please$/)
  assert.equal(harness.tools[0].name, "office_connect")
  assert.equal(harness.tools[0].input.properties.reconnect.type, "boolean")
})

test("not connected: events report nothing and no MCP server is added", async () => {
  harness.push({ type: "session.status", data: { sessionID: "a", status: { type: "busy" } } })
  harness.push({ type: "session.idle", data: { sessionID: "a" } })
  await tick()
  assert.equal(harness.servers.size, 0)
  assert.deepEqual(office.calls, [])
})

test("/office: OAuth with PKCE, stored tokens, MCP server added", async () => {
  let authorizeUrl = ""
  const result = harness.tools[0].execute({}, { progress: async (m: any) => (authorizeUrl = m.metadata.url) })
  await until(() => authorizeUrl !== "", "the authorization URL")
  // Azri's browser: the Office approves and sends it back to the plugin's callback.
  const back = await fetch(authorizeUrl, { redirect: "manual" })
  assert.equal(back.status, 302)
  const landing = await fetch(back.headers.get("location")!)
  assert.match(await landing.text(), /connected/)
  const response = await result
  assert.match(response.content, /^Connected to Agent Office as @opencode \(OpenCode\)\. Open tasks: 1, unclaimed team tasks: 0, unread messages: 2\./)

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
  assert.deepEqual(harness.servers.get("office"), { type: "remote", url: saved.office + "/mcp", oauth: false, headers: { Authorization: "Bearer access-1" } })
  assert.deepEqual(mcpCalls(), ["Bearer access-1 checkin {}"])
  office.calls.length = 0
})

test("working while any session is busy, idle when all are quiet", async () => {
  harness.push({ type: "session.status", data: { sessionID: "a", status: { type: "busy" } } })
  harness.push({ type: "session.status", data: { sessionID: "b", status: { type: "retry", attempt: 1 } } })
  harness.push({ type: "session.idle", data: { sessionID: "a" } })
  await until(() => mcpCalls().includes('Bearer access-1 checkin {"status":"working","note":"opencode: agent-office","rules_etag":"etag-1"}'), "the working report")
  harness.push({ type: "session.deleted", data: { sessionID: "b" } })
  await until(() => mcpCalls().some((call) => call.endsWith('{"status":"idle","note":"opencode: agent-office","rules_etag":"etag-1"}')), "the idle report")
  assert.deepEqual(mcpCalls(), ['Bearer access-1 checkin {"status":"working","note":"opencode: agent-office","rules_etag":"etag-1"}', 'Bearer access-1 checkin {"rules_etag":"etag-1"}', 'Bearer access-1 checkin {"status":"idle","note":"opencode: agent-office","rules_etag":"etag-1"}'])
  office.calls.length = 0
})

test("never overrides blocked", async () => {
  harness.push({ type: "session.status", data: { sessionID: "c", status: { type: "busy" } } })
  await until(() => office.calls.some((c) => c.path === "/mcp"), "the working report")
  office.state.workStatus = "blocked" // the model asked Azri something through `ask`
  harness.push({ type: "session.status", data: { sessionID: "c", status: { type: "idle" } } })
  await until(() => mcpCalls().includes('Bearer access-1 checkin {"rules_etag":"etag-1"}'), "the blocked check")
  await tick()
  assert.deepEqual(mcpCalls(), ['Bearer access-1 checkin {"status":"working","note":"opencode: agent-office","rules_etag":"etag-1"}', 'Bearer access-1 checkin {"rules_etag":"etag-1"}'])
  office.state.workStatus = "idle"
  office.calls.length = 0
})

test("polls and notifies the active session about new Office items", async () => {
  office.calls.length = 0
  const polled = makeContext("/poll", { pollInterval: 1_000 })
  const cleanup = await plugin.setup(polled.ctx)
  await until(() => mcpCalls().length > 0, "the baseline poll")
  office.calls.length = 0
  polled.push({ type: "session.status", data: { sessionID: "p", status: { type: "busy" } } })
  await until(() => office.calls.some((c) => c.path === "/mcp"), "the working report")
  office.state.unreadMessages = [...office.state.unreadMessages, { id: "message-new", from: "azri", subject: "New work" }]
  await until(() => polled.synthetics.length > 0, "the poll notification")
  assert.equal(polled.synthetics[0].sessionID, "p")
  assert.match(polled.synthetics[0].text, /1 new unread message/)
  assert.match(polled.synthetics[0].text, /message:message-new/)
  await cleanup()
  office.calls.length = 0
})

test("cleanup reports idle while a session is still busy", async () => {
  const fresh = makeContext("/tmp/x")
  const cleanup = await plugin.setup(fresh.ctx)
  await until(() => mcpCalls().length > 0, "the baseline poll")
  office.calls.length = 0
  fresh.push({ type: "session.status", data: { sessionID: "z", status: { type: "busy" } } })
  await until(() => office.calls.some((c) => c.path === "/mcp"), "the working report")
  office.calls.length = 0
  await cleanup()
  assert.deepEqual(mcpCalls(), ['Bearer access-1 checkin {"rules_etag":"etag-1"}', 'Bearer access-1 checkin {"status":"idle","note":"opencode: x","rules_etag":"etag-1"}'])
  office.calls.length = 0
})

test("a stored token near expiry is refreshed before use and handed to OpenCode again", async () => {
  const saved = JSON.parse(readFileSync(credentials, "utf8"))
  writeFileSync(credentials, JSON.stringify({ ...saved, expires_at: Date.now() + 60_000 }))
  const second = makeContext("/x")
  await plugin.setup(second.ctx)
  second.push({ type: "session.created", data: { info: { id: "d" } } })
  await until(() => second.servers.get("office")?.headers?.Authorization === "Bearer access-2", "the refreshed MCP server")
  const refresh = office.calls.find((c) => c.path === "/oauth/token")!
  assert.equal(refresh.body.grant_type, "refresh_token")
  assert.equal(refresh.body.refresh_token, "refresh-1")
  assert.equal(JSON.parse(readFileSync(credentials, "utf8")).refresh_token, "refresh-2")
  second.push({ type: "session.status", data: { sessionID: "d", status: { type: "busy" } } })
  await tick()
})

test("logs failures without the token", async () => {
  await office.close()
  harness.push({ type: "session.status", data: { sessionID: "e", status: { type: "busy" } } })
  await until(() => warnings.some((warning) => /could not report working/.test(warning)), "the failure warning")
  assert.ok(warnings.some((warning) => /could not report working/.test(warning)))
  assert.ok(!warnings.join("\n").includes("access-"))
})
