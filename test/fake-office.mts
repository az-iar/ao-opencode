// A stand-in for the Office: its OAuth endpoints (approval is automatic) and
// an MCP endpoint with a `checkin` tool. Records every call for assertions.
import { createServer, type Server } from "node:http"

export type Call = { method: string; path: string; auth?: string; headers?: Record<string, string | undefined>; body: any }

export function fakeOffice() {
  const calls: Call[] = []
  // Requests the Office is still answering. `idle()` resolves when none is, which
  // is the fact a test's finally can wait on: a request a plugin left in flight
  // after its cleanup() is answered before the log is cleared, so it cannot be
  // counted by the test that runs next. A count, never a quiet period.
  let inFlight = 0
  let answered: (() => void)[] = []
  const state = {
    // The access token the token endpoint issues. A test pins its own, so the
    // calls the approving plugin makes are attributable to that test and no
    // other test's are.
    issuedToken: undefined as string | undefined,
    workStatus: "working",
    // Who this Office says the caller is; a test sets another agent. The
    // delivery mode is what the poller reads on every checkin (Bridge M3):
    // undefined is an Office older than the field.
    agent: { slug: "opencode", name: "OpenCode", delivery_mode: undefined as string | undefined },
    tokens: 0,
    openTasks: [{ id: "task-open-1", title: "Open task" }],
    unclaimedTasks: [] as any[],
    // Tool names that answer with a tool error, as a refused claim would.
    failing: new Set<string>(),
    // A narrower refusal, by tool name and arguments.
    refuse: undefined as undefined | ((name: string, args: any) => boolean),
    // An HTTP status to answer with instead, as an Office deploy would.
    httpError: undefined as undefined | ((name: string, args: any) => number | undefined),
    // What get_task returns for a task id; else the open task of that id;
    // other ids get the checkin data.
    taskDetails: new Map<string, any>(),
    unreadMessages: [
      { id: "message-old-1", from: "azri", subject: "Old one" },
      { id: "message-old-2", from: "azri", subject: "Old two" },
    ],
  }
  const server: Server = createServer(async (req, res) => {
    inFlight++
    // 'close' fires when the response is finished or the client goes away, so
    // this is "this request is done", not "the handler returned".
    res.on("close", () => {
      if (--inFlight > 0) return
      for (const done of answered.splice(0)) done()
    })
    const chunks: Buffer[] = []
    for await (const chunk of req) chunks.push(chunk as Buffer)
    const raw = Buffer.concat(chunks).toString()
    const url = new URL(req.url ?? "/", "http://x")
    const json = (status: number, body: unknown) => {
      res.writeHead(status, { "Content-Type": "application/json" })
      res.end(JSON.stringify(body))
    }
    const body = req.headers["content-type"]?.includes("json") ? JSON.parse(raw || "{}") : Object.fromEntries(new URLSearchParams(raw))
    // `headers` keeps what a test needs to assert about the identity headers
    // GH-476 added; `auth` stays as it was for the tests that read only that.
    // `auth` is recorded on every call, including the ones that carried no
    // Authorization header at all (undefined), so a test can tell a call made
    // without a token from one made with another test's.
    calls.push({
      method: req.method ?? "",
      path: url.pathname,
      auth: req.headers.authorization,
      headers: req.headers as Record<string, string | undefined>,
      body,
    })

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
      return json(200, { access_token: state.issuedToken ?? `access-${state.tokens}`, refresh_token: `refresh-${state.tokens}`, expires_in: 3600, token_type: "Bearer", scope: "office:read office:write office:events" })
    }
    if (url.pathname === "/mcp") {
      if (!req.headers.authorization?.startsWith("Bearer ")) return json(401, { error: "unauthorized" })
      const result = (r: unknown) => json(200, { jsonrpc: "2.0", id: body.id, result: r })
      if (body.method === "initialize") return result({ protocolVersion: "2025-06-18", capabilities: { tools: {} }, serverInfo: { name: "fake office", version: "0" } })
      if (body.method?.startsWith("notifications/")) return res.writeHead(202).end()
      if (body.method === "tools/list") return result({ tools: [{ name: "checkin", description: "Check in", inputSchema: { type: "object", properties: {} } }] })
      if (body.method === "tools/call") {
        const status = state.httpError?.(body.params.name, body.params.arguments)
        if (status) return json(status, { error: "bad gateway" })
        if (state.failing.has(body.params.name) || state.refuse?.(body.params.name, body.params.arguments)) return result({ isError: true, content: [{ type: "text", text: `${body.params.name} refused` }] })
        if (body.params.name === "get_task") {
          // A detail set for the id wins; else the open task of that id, as
          // the real get_task returns the task with its current status.
          const id = body.params.arguments?.task_id
          const open = state.openTasks.find((t: any) => t.id === id)
          if (state.taskDetails.has(id) || open) {
            const detail = { task: state.taskDetails.has(id) ? state.taskDetails.get(id) : open }
            return result({ content: [{ type: "text", text: JSON.stringify(detail) }], structuredContent: detail })
          }
        }
        if (body.params.arguments?.status) state.workStatus = body.params.arguments.status
        const data = {
          agent: { ...state.agent, work_status: state.workStatus },
          pending: { open_tasks: state.openTasks, unclaimed_team_tasks: state.unclaimedTasks, unread_messages: state.unreadMessages },
          rules: { etag: "etag-1", effective: [], always: [], projects: [], layers: [] },
        }
        return result({ content: [{ type: "text", text: JSON.stringify(data) }], structuredContent: data })
      }
      return json(200, { jsonrpc: "2.0", id: body.id, error: { code: -32601, message: "no such method" } })
    }
    json(404, { error: "not found" })
  })
  return {
    calls,
    // The calls that arrived with `token`'s bearer, in arrival order: the
    // attribution a test asserts on. A straggler from an earlier test stays in
    // `calls` under that test's token instead of being dropped.
    byToken: (token: string) => calls.filter((c) => c.auth === `Bearer ${token}`),
    state,
    listen: () => new Promise<string>((r) => server.listen(0, "127.0.0.1", () => r(`http://127.0.0.1:${(server.address() as { port: number }).port}`))),
    // Resolves when this Office is answering nothing. Resolves at once when
    // there is nothing in flight, so a caller that has none is not delayed.
    idle: () => (inFlight === 0 ? Promise.resolve() : new Promise<void>((done) => answered.push(done))),
    inFlight: () => inFlight,
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
