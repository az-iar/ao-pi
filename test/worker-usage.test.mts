// A dispatched worker reports what its own run cost (Office 1b8f755a).
//
// The counts are Pi's own record of the run, so the worker is where they are
// reported from: its session entries are the source, `report_usage` is the route
// (POST /api/v1/agent/usage, agent-office#843), and the task it paid for and the
// preset it ran with are handed down by the dispatcher in the environment.
//
// This drives the real extension in child mode - the way a worker Pi runs it -
// through agent_settled, and asserts what the fake office received. No worker is
// spawned, no network leaves the process, and no token or credential is printed:
// the credentials file below is the fixture the existing tests use.
import assert from "node:assert/strict"
import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { after, test } from "node:test"
import { type Call, fakeOffice } from "./fake-office.mts"
import { tempDir } from "./tmp.mts"

const office = fakeOffice()
office.state.unreadMessages = []
office.state.openTasks = []
office.state.unclaimedTasks = []
const url = await office.listen()
after(() => office.close())

const root = tempDir("worker-usage")
const home = join(root, "home")
const repo = join(root, "repo")
const store = join(root, "store")
for (const dir of [home, repo, store]) mkdirSync(dir, { recursive: true })
const credentials = join(store, "pi.json")
writeFileSync(
  credentials,
  JSON.stringify({ office: url, client_id: "client-1", access_token: "access-worker", refresh_token: "refresh-worker", expires_at: Date.now() + 3_600_000 }),
)

// The task this worker was dispatched for, as Office names it: a uuid, because
// the spend route refuses a task_id that is not one.
const TASK = "3f2a1a4e-9c1d-4f2b-8a55-0f3b7c9d1e21"
process.chdir(repo)
process.env.HOME = home
process.env.AGENT_OFFICE_URL = url
process.env.AGENT_OFFICE_POLL_MS = "60000"
process.env.AGENT_OFFICE_CHILD = "1"
process.env.AGENT_OFFICE_PI_CREDENTIALS = credentials
process.env.AGENT_OFFICE_TASK_ID = TASK
// What the dispatcher hands a worker whose preset mapped to both tiers.
process.env.AGENT_OFFICE_PRESET_RAN = JSON.stringify({ model: "strong", thinking: "high" })
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_AUTODISPATCH
delete process.env.OPTCHAT_PROFILE

// The extension's own logging goes through console.warn; a test that asserts the
// "nothing reported" line has to see it rather than print it.
const warns: string[] = []
console.warn = (...args: unknown[]) => {
  warns.push(args.map((arg) => (typeof arg === "string" ? arg : JSON.stringify(arg))).join(" "))
}

const { default: extension } = await import("../index.ts")

const handlers = new Map<string, (event: unknown, ctx: unknown) => unknown>()
const notices: string[] = []
extension({
  on: (name: string, handler: (event: unknown, ctx: unknown) => unknown) => handlers.set(name, handler),
  registerCommand() {},
  registerTool() {},
  registerMcpServer() {},
  unregisterMcpServer() {},
  sendUserMessage() {},
} as any)

type Entry = { type: string; message?: { role: string; model?: string; usage?: Record<string, number> }; usage?: Record<string, number> }

const usage = (input: number, output: number, cacheRead: number, cacheWrite: number) => ({ input, output, cacheRead, cacheWrite, totalTokens: input + output + cacheRead + cacheWrite, cost: { total: 0 } })

// An assistant message Pi billed: the shape a session entry carries (usage on the
// message, model named beside it).
const assistant = (model: string, tokens: Record<string, number>): Entry => ({ type: "message", message: { role: "assistant", model, usage: tokens } })

// start runs the extension's session_start for one worker session, the way a
// worker Pi does when it boots with its dispatch grant.
async function start(run: string, entries: Entry[], startedAt = "2026-11-12T08:30:00.000Z") {
  const ctx = {
    sessionManager: {
      getSessionId: () => run,
      getEntries: () => entries,
      getHeader: () => ({ type: "session", id: run, timestamp: startedAt, cwd: repo }),
    },
    ui: { notify: (text: string) => notices.push(text) },
    isIdle: () => true,
  }
  await handlers.get("session_start")!({}, ctx)
  return ctx
}

// settle is the worker finishing: Pi's last awaited event before the process exits.
async function settle(ctx: unknown) {
  await handlers.get("agent_settled")!({ aborted: false }, ctx)
}

const toolNamed = (name: string, from: Call[] = office.calls): Call[] => from.filter((call) => call.path === "/mcp" && call.body?.method === "tools/call" && call.body?.params?.name === name)
const argumentsOf = (call: Call) => call.body?.params?.arguments ?? {}
// The fake stores one row per key, the way Office does; a test asks for the rows
// this run owns rather than every row the file has ever reported.
const storedFor = (run: string) => office.state.usageEvents.filter((event: any) => event.idempotency_key === `pi:${TASK}:${run}`)

test("a finished run posts its token counts once, with the task and the run's key", async () => {
  office.calls.length = 0
  const run = "session-usage-finished"
  const ctx = await start(run, [
    assistant("claude-sonnet-5-5", usage(1200, 300, 4000, 50)),
    assistant("claude-sonnet-5-5", usage(100, 20, 0, 0)),
  ])
  await settle(ctx)

  const reports = toolNamed("report_usage")
  assert.equal(reports.length, 1, "one report per run")
  const body = argumentsOf(reports[0])
  assert.equal(body.source, "harness:pi", "the harness this is: Pi")
  assert.equal(body.task_id, TASK, "the task the run paid for")
  assert.equal(body.model, "claude-sonnet-5-5", "the model that did the work")
  assert.deepEqual(body.tokens, { input: 1300, output: 320, cache_read: 4000, cache_write: 50 }, "the sum of Pi's own counts")
  assert.equal(body.day, "2026-11-12", "the UTC day the run started")
  assert.equal(body.idempotency_key, `pi:${TASK}:${run}`, "the key is the task and the run")
  assert.equal(storedFor(run).length, 1, "stored once")

  const updates = toolNamed("update_task")
  assert.equal(updates.length, 1, "the preset that ran is reported once")
  assert.deepEqual(argumentsOf(updates[0]).preset_ran, { model: "strong", thinking: "high" })
  assert.equal(argumentsOf(updates[0]).status, undefined, "reporting the preset does not move the task")
})

test("a retry of the same run reuses the key, so it does not count twice", async () => {
  office.calls.length = 0
  // The first report is refused by a gateway; the second has to carry the same
  // key, or Office would store the run's tokens twice.
  office.state.usageFailures = 1
  const run = "session-usage-retried"
  const ctx = await start(run, [assistant("claude-haiku-5-5", usage(500, 40, 900, 0))])
  await settle(ctx)

  const reports = toolNamed("report_usage")
  assert.equal(reports.length, 2, "the failed attempt was retried")
  assert.equal(argumentsOf(reports[0]).idempotency_key, `pi:${TASK}:${run}`)
  assert.equal(argumentsOf(reports[1]).idempotency_key, argumentsOf(reports[0]).idempotency_key, "the retry reuses the run's key")
  assert.equal(storedFor(run).length, 1, "Office stored the run once")
})

test("a run Pi recorded no counts for posts nothing, and says so once", async () => {
  office.calls.length = 0
  const before = warns.filter((line) => line.includes("no token counts")).length
  const ctx = await start("session-usage-no-counts", [{ type: "message", message: { role: "user", content: "hi" } }])
  await settle(ctx)

  assert.deepEqual(toolNamed("report_usage"), [], "nothing is posted from an estimate")
  assert.equal(storedFor("session-usage-no-counts").length, 0, "no spend row")
  assert.equal(warns.filter((line) => line.includes("no token counts")).length, before + 1, "logged once")
  // The preset did run, so it is still reported: the two facts are independent.
  assert.equal(toolNamed("update_task").length, 1)

  const again = await start("session-usage-no-counts-2", [{ type: "message", message: { role: "user", content: "hi" } }])
  await settle(again)
  assert.equal(warns.filter((line) => line.includes("no token counts")).length, before + 1, "a later run with no counts does not log again")
})

test("a report that keeps failing is logged and changes nothing", async () => {
  office.calls.length = 0
  const before = warns.filter((line) => line.includes("could not report this run's token usage")).length
  office.state.usageFailures = 5
  const run = "session-usage-unreachable"
  const ctx = await start(run, [assistant("claude-opus-5-5", usage(10, 5, 0, 0))])
  // The settle handler must not throw: a failed post is not a failed run.
  await settle(ctx)

  assert.equal(toolNamed("report_usage").length, 2, "one attempt and one retry, then it stops")
  assert.equal(storedFor(run).length, 0, "nothing stored")
  assert.equal(warns.filter((line) => line.includes("could not report this run's token usage")).length, before + 1, "logged, not thrown")
  assert.equal(toolNamed("update_task").length, 1, "the other report still went out")
})
