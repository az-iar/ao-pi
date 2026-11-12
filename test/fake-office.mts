// A stand-in for the Office: its OAuth endpoints (approval is automatic) and
// an MCP endpoint with a `checkin` tool. Records every call for assertions.
import { createServer, type Server } from "node:http"

export type Call = { method: string; path: string; auth?: string; headers?: Record<string, string | undefined>; body: any }

// A request a test is waiting for: `seen` settles it from inside the request
// handler, so the wait completes when the request actually arrives rather than
// when the code that sent it claims it did (ao-pi#41).
type Waiter = { match: (call: Call) => boolean; seen: (call: Call) => void; fail: (err: Error) => void }

export function fakeOffice() {
  const calls: Call[] = []
  const waiters: Waiter[] = []
  const state = {
    workStatus: "working",
    // The agent this Office answers checkins as; tests set it to the agent they
    // connected the workspace as.
    agentSlug: "pi",
    agentName: "Pi",
    // The presence delivery_mode Office sends on checkin: poll, webhook,
    // agent37, bridge, or undefined for a server that predates the field.
    deliveryMode: "poll" as string | undefined,
    tokens: 0,
    // How many upcoming token requests a gateway in front of the Office answers
    // with an HTML 502 instead of the token: the failure shape that used to
    // surface as "SyntaxError: Unexpected token '<'" (ao-pi MCP refresh).
    tokenFailures: 0,
    // How many upcoming token requests the Office itself refuses with a 400
    // invalid_grant: the grant is dead, and the same request would be refused
    // again, which is what a retry has to stop on.
    tokenRefusals: 0,
    // How many upcoming token requests are dropped without a reply at all: an
    // unreachable Office, which fetch reports as a network error rather than a
    // status, and which the retry has to keep trying through.
    tokenDrops: 0,
    openTasks: [{ id: "task-open-1", title: "Open task" }],
    unclaimedTasks: [] as any[],
    // The spend rows a worker's report_usage stored, one per idempotency key, the
    // way Office stores them (agent-office#843): a repeat with the key already
    // there answers 200 and stores nothing.
    usageEvents: [] as any[],
    // How many upcoming report_usage calls a gateway refuses with a 500: the
    // transient failure a retry has to survive without counting twice.
    usageFailures: 0,
    // What get_task answers, per task id. A dispatch re-reads the task before it
    // claims it, so a test can show the poller's snapshot and the current task
    // disagreeing (a task cancelled between the two). Absent, get_task falls back
    // to the copy in openTasks or unclaimedTasks.
    tasks: {} as Record<string, any>,
    unreadMessages: [
      { id: "message-old-1", from: "azri", subject: "Old one" },
      { id: "message-old-2", from: "azri", subject: "Old two" },
    ],
  }
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
    // `headers` keeps what a test needs to assert about the identity headers
    // GH-476 added; `auth` stays as it was for the tests that read only that.
    calls.push({
      method: req.method ?? "",
      path: url.pathname,
      auth: req.headers.authorization,
      headers: req.headers as Record<string, string | undefined>,
      body,
    })
    for (const waiter of [...waiters]) {
      if (!waiter.match(calls[calls.length - 1])) continue
      waiters.splice(waiters.indexOf(waiter), 1)
      waiter.seen(calls[calls.length - 1])
    }

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
      if (state.tokenRefusals > 0) {
        state.tokenRefusals--
        return json(400, { error: "invalid_grant", error_description: "the refresh token is expired or revoked" })
      }
      if (state.tokenDrops > 0) {
        state.tokenDrops--
        // No status and no body: the connection dies, so the caller sees a
        // network failure rather than an HTTP one.
        res.socket?.destroy()
        return
      }
      if (state.tokenFailures > 0) {
        state.tokenFailures--
        res.writeHead(502, { "Content-Type": "text/html" })
        res.end("<!doctype html><html><head><title>502 Bad Gateway</title></head><body><h1>502 Bad Gateway</h1><hr><center>nginx</center></body></html>")
        return
      }
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
        const tool = String(body.params?.name ?? "")
        const args = body.params?.arguments ?? {}
        const structured = (value: unknown) => result({ content: [{ type: "text", text: JSON.stringify(value) }], structuredContent: value })
        if (tool === "report_usage") {
          if (state.usageFailures > 0) {
            state.usageFailures--
            res.writeHead(500, { "Content-Type": "application/json" })
            return res.end(JSON.stringify({ error: "gateway" }))
          }
          const key = String(args.idempotency_key ?? "")
          const stored = state.usageEvents.find((event: any) => event.idempotency_key === key)
          if (stored) return structured({ event: stored, created: false })
          const event = { ...args, id: `usage-${state.usageEvents.length + 1}` }
          state.usageEvents.push(event)
          return structured({ event, created: true })
        }
        // The task tools a dispatch drives: get_task answers the current state,
        // and update_task moves it, so a second read agrees with the write.
        if (tool === "get_task") {
          const id = String(args.task_id ?? "")
          const found = state.tasks[id] ?? [...state.openTasks, ...state.unclaimedTasks].find((item: any) => String(item?.id ?? item?.task_id) === id)
          return structured({ task: found ?? { id, status: "todo" } })
        }
        if (tool === "update_task") {
          const task = state.tasks[String(args.task_id ?? "")]
          if (task && typeof args.status === "string") task.status = args.status
          return structured({ ok: true, task_id: args.task_id, status: args.status })
        }
        if (tool === "ack_task" || tool === "accept_task") return structured({ ok: true, task_id: args.task_id })
        if (args.status) state.workStatus = args.status
        const data = {
          agent: {
            slug: state.agentSlug,
            name: state.agentName,
            work_status: state.workStatus,
            ...(state.deliveryMode === undefined ? {} : { delivery_mode: state.deliveryMode }),
          },
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
    state,
    listen: () => new Promise<string>((r) => server.listen(0, "127.0.0.1", () => r(`http://127.0.0.1:${(server.address() as { port: number }).port}`))),
    // Resolves with the FIRST request that matches, including one already served,
    // and rejects with `message` when none arrives: a regression has to fail the
    // test in seconds rather than leave it pending. The timeout is a failure guard,
    // not the wait - the wait itself is the request. Because it is the first match,
    // the match has to identify the request, not merely its path: a path alone can
    // be satisfied by another session's call, and the test then asserts against a
    // request it did not mean (ao-pi#45).
    waitForCall: (match: (call: Call) => boolean, message: string, timeoutMs = 5_000) =>
      new Promise<Call>((resolve, reject) => {
        const served = calls.find(match)
        if (served) return resolve(served)
        let timer: ReturnType<typeof setTimeout>
        const waiter: Waiter = {
          match,
          seen: (call) => {
            clearTimeout(timer)
            resolve(call)
          },
          fail: (err) => {
            clearTimeout(timer)
            reject(err)
          },
        }
        timer = setTimeout(() => {
          waiters.splice(waiters.indexOf(waiter), 1)
          waiter.fail(new Error(message))
        }, timeoutMs)
        waiters.push(waiter)
      }),
    close: () =>
      new Promise<void>((r) => {
        // Nothing can arrive once the server is closing, so nobody may be left
        // waiting on it.
        for (const waiter of waiters.splice(0)) waiter.fail(new Error("the fake office closed before the expected request arrived"))
        server.closeAllConnections()
        server.close(() => r())
      }),
  }
}

// `node test/fake-office.mts` runs it standalone for an integration check.
if (process.argv[1]?.endsWith("fake-office.mts")) {
  const office = fakeOffice()
  console.log(await office.listen())
  setInterval(() => {
    const call = office.calls.shift()
    if (call) console.log(JSON.stringify(call))
  }, 100)
}
