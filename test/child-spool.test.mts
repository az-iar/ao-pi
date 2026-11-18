// A child's token usage, queued for the parent that started it (Office dd185d2b).
//
// A Pi started by `hc run` under a Claude Code session is a child
// (AGENT_OFFICE_CHILD=1) with no Office identity of its own, so it must not post
// its spend. With AGENT_OFFICE_USAGE_SPOOL it queues one file per settle instead,
// and the parent's ao-claude plugin posts them as the agent that started it.
//
// This drives the real extension in child mode through agent_settled and
// session_shutdown and asserts the files, then asserts what the fake office did
// not receive. The child is given credentials of its own (as a dispatched worker
// is), so "it posted nothing" means it chose not to, not that it could not. No
// network leaves the process and no token is printed. The pure range/key/body
// rules are pinned in test/spool.test.mts.
import assert from "node:assert/strict"
import { existsSync, mkdirSync, readFileSync, readdirSync, symlinkSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { after, test } from "node:test"
import { type Call, fakeOffice } from "./fake-office.mts"
import { cursorFileName, spoolFileName } from "../lib/spool.mts"
import { tempDir } from "./tmp.mts"

type Entry = { id?: string; type?: string; message?: { role?: string; model?: string; usage?: Record<string, number> }; usage?: Record<string, number> }
const usage = (input: number, output: number, cacheRead: number, cacheWrite: number) => ({ input, output, cacheRead, cacheWrite, totalTokens: input + output + cacheRead + cacheWrite, cost: { total: 0 } })
const assistant = (id: string, model: string, tokens: Record<string, number>): Entry => ({ id, type: "message", message: { role: "assistant", model, usage: tokens } })
const spoolFiles = (dir: string) => (existsSync(dir) ? readdirSync(dir).filter((name) => name.endsWith(".json")) : [])
const bodyOf = (path: string) => JSON.parse(readFileSync(path, "utf8"))

const office = fakeOffice()
office.state.unreadMessages = []
office.state.openTasks = []
office.state.unclaimedTasks = []
const url = await office.listen()
after(() => office.close())

const root = tempDir("child-spool")
const home = join(root, "home")
const repo = join(root, "repo")
const store = join(root, "store")
const spool = join(root, "spool")
for (const dir of [home, repo, store]) mkdirSync(dir, { recursive: true })
const credentials = join(store, "pi.json")
writeFileSync(
  credentials,
  JSON.stringify({ office: url, client_id: "client-1", access_token: "access-child", refresh_token: "refresh-child", expires_at: Date.now() + 3_600_000 }),
)

const TASK = "3f2a1a4e-9c1d-4f2b-8a55-0f3b7c9d1e21"
process.chdir(repo)
process.env.HOME = home
process.env.AGENT_OFFICE_URL = url
process.env.AGENT_OFFICE_POLL_MS = "60000"
process.env.AGENT_OFFICE_CHILD = "1"
process.env.AGENT_OFFICE_PI_CREDENTIALS = credentials
process.env.AGENT_OFFICE_TASK_ID = TASK
process.env.AGENT_OFFICE_USAGE_SPOOL = spool
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_AUTODISPATCH
delete process.env.AGENT_OFFICE_PRESET_RAN
delete process.env.OPTCHAT_PROFILE

const warns: string[] = []
console.warn = (...args: unknown[]) => {
  warns.push(args.map((arg) => (typeof arg === "string" ? arg : JSON.stringify(arg))).join(" "))
}

const { default: extension } = await import("../index.ts")

const handlers = new Map<string, (event: unknown, ctx: unknown) => unknown>()
extension({
  on: (name: string, handler: (event: unknown, ctx: unknown) => unknown) => handlers.set(name, handler),
  registerCommand() {},
  registerTool() {},
  registerMcpServer() {},
  unregisterMcpServer() {},
  sendUserMessage() {},
} as any)

// start runs the extension's session_start for one child session, as a child Pi
// does when it boots under `hc run`.
async function start(run: string, entries: Entry[]) {
  const ctx = {
    sessionManager: {
      getSessionId: () => run,
      getEntries: () => entries,
      getHeader: () => ({ type: "session", id: run, timestamp: "2026-11-18T08:30:00.000Z", cwd: repo }),
    },
    ui: { notify: () => {} },
    isIdle: () => true,
  }
  await handlers.get("session_start")!({}, ctx)
  return ctx
}

const settle = (ctx: unknown) => handlers.get("agent_settled")!({ aborted: false }, ctx)
const shutdown = (ctx: unknown) => handlers.get("session_shutdown")!({ type: "session_shutdown", reason: "quit" }, ctx)
const toolNamed = (name: string) => office.calls.filter((call: Call) => call.path === "/mcp" && call.body?.method === "tools/call" && call.body?.params?.name === name)

test("a child settling never posts: it queues one file per settle with only that settle's usage", async () => {
  office.calls.length = 0
  const run = "session-spool-two-settles"
  const entries = [assistant("c1", "claude-sonnet-5-5", usage(1200, 300, 4000, 50))]
  const ctx = await start(run, entries)
  await settle(ctx)

  let files = spoolFiles(spool)
  assert.equal(files.length, 1, "one settle, one report")
  const first = bodyOf(join(spool, files[0]))
  assert.equal(first.idempotency_key, `pi:${run}:c1`, "the key names the session and the entry it ends at")
  assert.deepEqual(first.tokens, { input: 1200, output: 300, cache_read: 4000, cache_write: 50 })
  assert.equal(first.day, new Date().toISOString().slice(0, 10), "the UTC day of the settle")
  assert.equal(first.task_id, TASK)
  assert.equal(first.source, "harness:pi")
  assert.equal("amount" in first, false, "Office prices the tokens")
  assert.deepEqual(toolNamed("report_usage"), [], "the child posted nothing: it has no identity to post with")

  // The same long-lived child does more work and settles again.
  entries.push(assistant("c2", "claude-sonnet-5-5", usage(100, 20, 0, 0)))
  await settle(ctx)
  files = spoolFiles(spool)
  assert.equal(files.length, 2, "the second settle queued its own report")
  const second = bodyOf(join(spool, files.find((name) => name !== spoolFileName(first.idempotency_key))!))
  assert.equal(second.idempotency_key, `pi:${run}:c2`)
  assert.deepEqual(second.tokens, { input: 100, output: 20, cache_read: 0, cache_write: 0 }, "only the usage this settle added")
  assert.deepEqual(toolNamed("report_usage"), [], "still nothing posted")

  // A settle with nothing new to say writes nothing.
  await settle(ctx)
  assert.equal(spoolFiles(spool).length, 2)
  assert.deepEqual(readdirSync(spool).filter((name) => name.endsWith(".tmp")), [], "no temporary file is left behind")
})

test("a later settle with nothing new queues nothing", async () => {
  // A long-lived child settles over and over; only what it added since the last
  // report is new spend. (A replayed range is the pure test's case.)
  const run = "session-spool-quiet"
  const ctx = await start(run, [assistant("r1", "claude-haiku-5-5", usage(500, 40, 900, 0))])
  await settle(ctx)
  const before = spoolFiles(spool).length
  await settle(ctx)
  assert.equal(spoolFiles(spool).length, before, "nothing new was recorded, so nothing new was queued")
})

test("a child with no task id writes a body with no task_id", async () => {
  office.calls.length = 0
  delete process.env.AGENT_OFFICE_TASK_ID
  try {
    const run = "session-spool-no-task"
    const ctx = await start(run, [assistant("n1", "claude-haiku-5-5", usage(70, 7, 0, 0))])
    await settle(ctx)
    const file = spoolFiles(spool)
      .map((name) => join(spool, name))
      .find((path) => bodyOf(path).idempotency_key === `pi:${run}:n1`)
    assert.ok(file, "the report was queued")
    const body = bodyOf(file!)
    assert.equal("task_id" in body, false, "a child started for no task attributes its spend to none")
    assert.deepEqual(Object.keys(body).sort(), ["day", "idempotency_key", "model", "source", "tokens"])
    assert.deepEqual(toolNamed("report_usage"), [], "the child still posts nothing")
  } finally {
    process.env.AGENT_OFFICE_TASK_ID = TASK
  }
})

test("a child that settles through shutdown queues what the last turn recorded", async () => {
  const run = "session-spool-shutdown"
  const ctx = await start(run, [assistant("s1", "claude-opus-5-5", usage(30, 3, 0, 0))])
  await shutdown(ctx)
  const files = spoolFiles(spool).map((name) => join(spool, name))
  const queued = files.map(bodyOf).find((body) => body.idempotency_key === `pi:${run}:s1`)
  assert.deepEqual(queued?.tokens, { input: 30, output: 3, cache_read: 0, cache_write: 0 }, "print mode exits through shutdown, and its last usage is queued")
})

test("with no spool directory the child behaves exactly as before and posts its usage", async () => {
  office.calls.length = 0
  delete process.env.AGENT_OFFICE_USAGE_SPOOL
  try {
    const run = "session-spool-off"
    const before = spoolFiles(spool).length
    const ctx = await start(run, [assistant("o1", "claude-sonnet-5-5", usage(400, 40, 0, 0))])
    await settle(ctx)
    const reports = toolNamed("report_usage")
    assert.equal(reports.length, 1, "today's behaviour: the worker posts its own run")
    assert.equal(reports[0].body.params.arguments.idempotency_key, `pi:${TASK}:${run}`, "with the run's key")
    assert.equal(spoolFiles(spool).length, before, "and writes nothing into the spool")

    // A relative path is no directory to spool into: the child cannot know where
    // its parent looks, and spooling into its own cwd would queue nothing anyone
    // reads. It posts as before and creates nothing.
    process.env.AGENT_OFFICE_USAGE_SPOOL = "relative/spool"
    office.calls.length = 0
    const relative = await start("session-spool-relative", [assistant("p1", "claude-sonnet-5-5", usage(1, 1, 0, 0))])
    await settle(relative)
    assert.equal(toolNamed("report_usage").length, 1, "a relative path is off, so the child posts as before")
    assert.equal(existsSync(join(repo, "relative")), false, "and creates nothing in its working directory")
    assert.equal(spoolFiles(spool).length, before)
  } finally {
    process.env.AGENT_OFFICE_USAGE_SPOOL = spool
  }
})

test("a cursor the child cannot read is logged once, and spooling starts over", async () => {
  // A resumed session whose cursor the spool cannot read: a symlink to itself
  // fails whoever the process runs as, and the child has to start from the
  // beginning of the session. That is said once, because repeating it every
  // settle would bury the run's own output.
  const run = "session-spool-bad-cursor"
  mkdirSync(spool, { recursive: true })
  symlinkSync(cursorFileName(run), join(spool, cursorFileName(run)))
  const before = warns.filter((line) => line.includes("could not read this child's spool cursor")).length

  const entries = [assistant("k1", "claude-sonnet-5-5", usage(10, 1, 0, 0))]
  const ctx = await start(run, entries)
  await settle(ctx)
  assert.equal(warns.filter((line) => line.includes("could not read this child's spool cursor")).length, before + 1, "logged once")
  assert.equal(readFileSync(join(spool, cursorFileName(run)), "utf8").trim(), "k1", "a usable cursor replaced the unreadable one")

  // The next settle is back to a delta, and does not log again.
  entries.push(assistant("k2", "claude-sonnet-5-5", usage(300, 30, 0, 0)))
  await settle(ctx)
  assert.equal(warns.filter((line) => line.includes("could not read this child's spool cursor")).length, before + 1, "not said twice")
  const bodies = spoolFiles(spool).map((name) => bodyOf(join(spool, name)))
  const last = bodies.find((body: any) => body.idempotency_key === `pi:${run}:k2`)
  assert.deepEqual(last?.tokens, { input: 300, output: 30, cache_read: 0, cache_write: 0 }, "only what the last settle added")
})

test("a spool that cannot be written is logged and changes nothing about the run", async () => {
  // A file where the directory should be: mkdir fails every time.
  const blocked = join(root, "blocked-parent")
  writeFileSync(blocked, "not a directory")
  const before = warns.filter((line) => line.includes("could not spool this run's token usage")).length
  process.env.AGENT_OFFICE_USAGE_SPOOL = join(blocked, "spool")
  try {
    const ctx = await start("session-spool-blocked", [assistant("b1", "claude-sonnet-5-5", usage(9, 9, 0, 0))])
    // The handler returns instead of throwing: the run ends as it would have.
    await settle(ctx)
    assert.equal(warns.filter((line) => line.includes("could not spool this run's token usage")).length, before + 1, "logged once")
    assert.ok(!warns.some((line) => line.includes("access-child")), "the log carries no token")
    assert.equal(existsSync(join(blocked, "spool")), false, "and nothing half-written was left behind")
  } finally {
    process.env.AGENT_OFFICE_USAGE_SPOOL = spool
  }
})
