// A stand-in for the Office: its OAuth endpoints (approval is automatic) and
// an MCP endpoint with a `checkin` tool. Records every call for assertions.
import { createServer, type Server } from "node:http"

export type Call = { method: string; path: string; auth?: string; body: any }

export function fakeOffice() {
  const calls: Call[] = []
  const state = { workStatus: "working", tokens: 0 }
  const server: Server = createServer(async (req, res) => {
    const chunks: Buffer[] = []
    for await (const chunk of req) chunks.push(chunk as Buffer)
    const raw = Buffer.concat(chunks).toString()
    const url = new URL(req.url ?? "/", "http://x")
    const json = (status: number, body: unknown) => {
      res.writeHead(status, { "Content-Type": "application/json" })
      res.end(JSON.stringify(body))
    }
    const body = req.headers["content-type"]?.includes("json") ? JSON.parse(raw || "{}") : Object.fromEntries(new URLSearchParams(raw))
    calls.push({ method: req.method ?? "", path: url.pathname, auth: req.headers.authorization, body })

    if (url.pathname === "/oauth/register") return json(201, { client_id: "client-1", redirect_uris: body.redirect_uris })
    if (url.pathname === "/oauth/authorize") {
      // Azri approves at once: back to the client with a code.
      const back = new URL(url.searchParams.get("redirect_uri")!)
      back.searchParams.set("code", "code-1")
      back.searchParams.set("state", url.searchParams.get("state")!)
      res.writeHead(302, { Location: back.href }).end()
      return
    }
    if (url.pathname === "/oauth/token") {
      state.tokens++
      return json(200, { access_token: `access-${state.tokens}`, refresh_token: `refresh-${state.tokens}`, expires_in: 3600, token_type: "Bearer", scope: "office:read office:write office:events" })
    }
    if (url.pathname === "/mcp") {
      if (!req.headers.authorization?.startsWith("Bearer ")) return json(401, { error: "unauthorized" })
      const result = (r: unknown) => json(200, { jsonrpc: "2.0", id: body.id, result: r })
      if (body.method === "initialize") return result({ protocolVersion: "2025-06-18", capabilities: { tools: {} }, serverInfo: { name: "fake office", version: "0" } })
      if (body.method?.startsWith("notifications/")) return res.writeHead(202).end()
      if (body.method === "tools/list") return result({ tools: [{ name: "checkin", description: "Check in", inputSchema: { type: "object", properties: {} } }] })
      if (body.method === "tools/call") {
        if (body.params.arguments?.status) state.workStatus = body.params.arguments.status
        const data = { agent: { slug: "opencode", name: "OpenCode", work_status: state.workStatus }, pending: { open_tasks: [{}], unclaimed_team_tasks: [], unread_messages: [{}, {}] } }
        return result({ content: [{ type: "text", text: JSON.stringify(data) }], structuredContent: data })
      }
      return json(200, { jsonrpc: "2.0", id: body.id, error: { code: -32601, message: "no such method" } })
    }
    json(404, { error: "not found" })
  })
  return {
    calls,
    state,
    listen: () => new Promise<string>((r) => server.listen(0, "127.0.0.1", () => r(`http://127.0.0.1:${(server.address() as { port: number }).port}`))),
    close: () => new Promise<void>((r) => (server.closeAllConnections(), server.close(() => r()))),
  }
}

// `node test/fake-office.mts` runs it standalone for the OpenCode integration check.
if (process.argv[1]?.endsWith("fake-office.mts")) {
  const office = fakeOffice()
  console.log(await office.listen())
  setInterval(() => {
    const call = office.calls.shift()
    if (call) console.log(JSON.stringify(call))
  }, 100)
}
