// ao-pi#4 / #32 / #33: what dispatchTask *does*, not what its source text says.
//
// test/dispatch-guards.test.mts pins the guards by reading index.ts, and two real
// defects slipped past exactly that style: the sandbox refusal ran after the
// claim, and the repo-less scratch path never set workerStarted, so a worker that
// finished was handed back to todo and re-dispatched on the next poll. Both were
// wiring, not logic, so they are only visible from outside: drive the real
// dispatch through the poller, point the worker binary at a stub, and assert the
// Office calls the fake records and whether a worktree appeared.
//
// The last case is #33's: a task cancelled while its worker runs stops that
// worker, and nothing is unclaimed afterwards.
//
// Nothing here reaches outside the process: AGENT_OFFICE_OPEN_BROWSER=off plus an
// xdg-open stub on PATH, and AGENT_OFFICE_PI_BIN is a stub script on disk.
import assert from "node:assert/strict"
import { execFileSync } from "node:child_process"
import { chmodSync, existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, rmSync, symlinkSync, unlinkSync, writeFileSync } from "node:fs"
import { createRequire } from "node:module"
import { basename, dirname, join } from "node:path"
import { after, before, test } from "node:test"
import { type Call, fakeOffice } from "./fake-office.mts"
import { tempDir } from "./tmp.mts"

// index.ts resolves the repo root from `require`, which Pi's loader provides and
// plain node ESM does not (see worktree.test.mts).
;(globalThis as { require?: NodeRequire }).require ??= createRequire(import.meta.url)

process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

const office = fakeOffice()
const url = await office.listen()
after(() => office.close())

// The task's repository, matched by the checkout's origin rather than by its
// directory name (agent-office#540).
const FIXTURE_REPO = "az-iar/fixture-repo"

const root = tempDir("dispatch-behaviour")
const home = join(root, "home")
const store = join(root, "store")
const bin = join(root, "bin")
const runs = join(root, "runs")
// Where the stub says it ran to completion. A worker that was stopped is in
// `runs` and not here, which is how the cancellation case tells the two apart.
const finished = join(root, "finished")
const repo = join(root, "repo")
for (const dir of [home, store, bin, runs, finished]) mkdirSync(dir, { recursive: true })

// The opener is a no-op that exits at once, so no run can launch a browser even
// if a test forgets the flag above.
writeFileSync(join(bin, "xdg-open"), "#!/bin/sh\nexit 0\n", { mode: 0o755 })
// A bwrap that cannot work, for the host where the real one can: prepended to
// PATH by the sandbox case only, so `sandboxAvailable()` answers no there.
const noBwrap = join(root, "bin-no-bwrap")
mkdirSync(noBwrap)
writeFileSync(join(noBwrap, "bwrap"), "#!/bin/sh\nexit 1\n", { mode: 0o755 })

// The worker stub records where it ran and how it was called, then exits with
// STUB_EXIT (0 by default). PI_BIN is read once when the extension loads, so a
// test that needs the launch itself to fail repoints the `pig` symlink at a path
// that does not exist: spawn answers ENOENT, which is the missing-binary case.
// STUB_SLEEP_MS holds a worker open - long enough for a work scope to still be
// held by a second copy, or for the dispatch's task watch to see a cancellation.
// The exit marker is written only on a normal exit, so a worker stopped by
// SIGTERM leaves its start record and nothing else.
const worker = join(bin, "worker.js")
writeFileSync(
  worker,
  `#!${process.execPath}\n` +
    `const fs = require("fs"), path = require("path")\n` +
    `const stamp = Date.now() + "-" + process.pid\n` +
    `fs.writeFileSync(path.join(${JSON.stringify(runs)}, stamp + ".json"), JSON.stringify({ cwd: process.cwd(), argv: process.argv.slice(2), child: process.env.AGENT_OFFICE_CHILD, credentials: process.env.AGENT_OFFICE_PI_CREDENTIALS }))\n` +
    `const done = () => { fs.writeFileSync(path.join(${JSON.stringify(finished)}, stamp + ".json"), JSON.stringify({ stamp })); process.exit(Number(process.env.STUB_EXIT ?? 0)) }\n` +
    `if (process.env.STUB_SLEEP_MS) setTimeout(done, Number(process.env.STUB_SLEEP_MS))\n` +
    `else done()\n`,
)
chmodSync(worker, 0o755)
const pig = join(bin, "pig")
const pointPig = (target: string) => {
  try {
    unlinkSync(pig)
  } catch {}
  symlinkSync(target, pig)
}
pointPig(worker)

// The session's own checkout, with an origin the fixture task's repo matches.
execFileSync("git", ["init", "-q", "-b", "main", repo])
const git = (...args: string[]) => execFileSync("git", ["-C", repo, ...args], { encoding: "utf8" })
git("-c", "user.email=t@example.com", "-c", "user.name=t", "commit", "-q", "--allow-empty", "-m", "init")
git("remote", "add", "origin", `https://github.com/${FIXTURE_REPO}.git`)

const credentials = join(store, "pi.json")
writeFileSync(
  credentials,
  JSON.stringify({ office: url, client_id: "client-1", access_token: "access-dispatch", refresh_token: "refresh-dispatch", expires_at: Date.now() + 3_600_000 }),
)
process.env.HOME = home
process.env.PATH = `${bin}:${process.env.PATH}`
process.env.AGENT_OFFICE_URL = url
process.env.AGENT_OFFICE_PI_CREDENTIALS = credentials
process.env.AGENT_OFFICE_PI_BIN = pig
process.env.AGENT_OFFICE_POLL_MS = "1000"
// #33 re-reads a running worker's task on a timer; a minute is right in a
// session and far too slow for a test, so the watch ticks here instead.
process.env.AGENT_OFFICE_WORKER_STATUS_POLL_MS = "100"
process.env.AGENT_OFFICE_SANDBOX = "0"
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_CHILD
delete process.env.AGENT_OFFICE_AUTODISPATCH
delete process.env.AGENT_OFFICE_TOKEN_SCOPE
process.chdir(repo)

// Loaded after the environment above: lib/office.ts resolves ~/.config from HOME
// and index.ts reads AGENT_OFFICE_PI_BIN when it loads.
const lib = await import("../lib/office.ts")
const { default: agentOffice } = await import("../index.ts")

type Notice = { text: string; level?: string }
const notices: Notice[] = []
const handlers = new Map<string, ((event: unknown, ctx: unknown) => unknown)[]>()
const tools = new Map<string, { execute: () => Promise<{ content: { text: string }[] }> }>()
const api = {
  on(name: string, handler: (event: unknown, ctx: unknown) => unknown) {
    handlers.set(name, [...(handlers.get(name) ?? []), handler])
  },
  registerCommand() {},
  registerTool(tool: { name: string; execute: () => Promise<{ content: { text: string }[] }> }) {
    tools.set(tool.name, tool)
  },
  registerMcpServer() {},
  unregisterMcpServer() {},
  sendUserMessage() {},
}
agentOffice(api as any)
const ctx = {
  cwd: repo,
  isIdle: () => true,
  ui: { notify: (text: string, level?: string) => notices.push({ text, level }) },
  sessionManager: { getSessionId: () => "session-dispatch-harness" },
}

const officeCalls = (): Call[] => office.calls.filter((call) => call.path === "/mcp" && call.body?.method === "tools/call")
const toolNamed = (name: string, from: Call[] = officeCalls()): Call[] => from.filter((call) => call.body?.params?.name === name)
const argumentsOf = (call: Call) => call.body?.params?.arguments ?? {}
const claimCalls = (calls: Call[]) => calls.filter((call) => ["ack_task", "accept_task"].includes(call.body.params.name))
const updateCalls = (calls: Call[]) => toolNamed("update_task", calls)
const withStatus = (calls: Call[], status: string) => calls.filter((call) => argumentsOf(call).status === status)
const recorded = () =>
  readdirSync(runs)
    .sort()
    .map((file) => JSON.parse(readFileSync(join(runs, file), "utf8")) as { cwd: string; argv: string[]; child?: string; credentials?: string })
const clearRuns = () => {
  for (const file of readdirSync(runs)) unlinkSync(join(runs, file))
  for (const file of readdirSync(finished)) unlinkSync(join(finished, file))
}
const completed = () => readdirSync(finished).length
// The worktree a dispatch of this task would make beside the session's checkout.
const worktreePath = (id: string) => join(dirname(repo), `${basename(repo)}-office-pi-${lib.taskShort(id)}`)
const samePath = (a: string, b: string) => realpathSync(a) === realpathSync(b)

async function waitFor(check: () => boolean | Promise<boolean>, what: string, timeoutMs = 15_000) {
  const start = Date.now()
  while (!(await check())) {
    if (Date.now() - start > timeoutMs) throw new Error(`timed out waiting for ${what}`)
    await new Promise((resolve) => setTimeout(resolve, 10))
  }
}

async function toolStatus() {
  return JSON.parse((await tools.get("office_status")!.execute()).content[0].text)
}

// Hands tasks to the running extension the way the Office does — through the
// poller's own task list — and waits for the dispatch to finish. `read` is what
// get_task answers, so a test can show the snapshot and the current task
// disagreeing.
async function dispatch(snapshot: any[], options: { read?: Record<string, any>; sandbox?: string; duringRun?: () => void | Promise<void> } = {}) {
  office.calls.length = 0
  notices.length = 0
  clearRuns()
  for (const dir of [join(store, "dispatch"), join(store, "notices")]) rmSync(dir, { recursive: true, force: true })
  office.state.unclaimedTasks = []
  office.state.unreadMessages = []
  office.state.tasks = {}
  const previousSandbox = process.env.AGENT_OFFICE_SANDBOX
  if (options.sandbox !== undefined) process.env.AGENT_OFFICE_SANDBOX = options.sandbox
  try {
    for (const task of snapshot) office.state.tasks[task.id] = { ...(options.read?.[task.id] ?? task) }
    office.state.openTasks = snapshot.map((task) => ({ ...task }))
    // The poller picks the task up on its next tick, and dispatchTask's first
    // Office call is get_task: that is the signal the dispatch has started.
    await waitFor(
      () => snapshot.every((task) => toolNamed("get_task").some((call) => argumentsOf(call).task_id === task.id)),
      `the poll to hand ${snapshot.map((task) => task.id).join(", ")} to a dispatch`,
    )
    // Something has to happen while the worker is up (a cancellation, in the
    // #33 case), and `runs` gets its first record as the stub starts.
    if (options.duringRun) {
      await waitFor(() => recorded().length > 0, "the worker to start")
      await options.duringRun()
    }
    // dispatchTask drops its id from `dispatching` only after its finally has
    // run, and the finally is where an unclaim is sent: an empty set means the
    // dispatch is over, Office writes included.
    await waitFor(async () => {
      const status = await toolStatus()
      return status.dispatching.length === 0 && status.backlog.length === 0
    }, "the dispatch to finish")
    return { calls: officeCalls().slice(), notices: notices.slice(), runs: recorded() }
  } finally {
    office.state.openTasks = []
    if (options.sandbox !== undefined) {
      if (previousSandbox === undefined) delete process.env.AGENT_OFFICE_SANDBOX
      else process.env.AGENT_OFFICE_SANDBOX = previousSandbox
    }
  }
}

before(async () => {
  for (const handler of handlers.get("session_start") ?? []) await handler({}, ctx)
  // The first poll is a baseline: it has to have happened before a task is
  // handed over, or the task would be baselined and never dispatched.
  await waitFor(() => toolNamed("checkin").length > 0, "the extension's first poll")
})

after(async () => {
  for (const handler of handlers.get("session_shutdown") ?? []) await handler({}, ctx)
  await office.close()
  rmSync(root, { recursive: true, force: true })
})

// The #32 defect: the scratch path launched the worker without setting
// workerStarted, so a worker that ran and exited 0 was treated as one that never
// started — the task was handed back to todo and dispatched again on the next
// poll, for ever.
test("a repo-less task whose worker runs is not unclaimed to todo", async () => {
  const task = { id: "task-scratch-ok", title: "Scratch work that finishes", repo: "", status: "todo" }
  const { calls, runs: ran } = await dispatch([task])
  assert.ok(claimCalls(calls).length > 0, "the task was claimed")
  const updates = updateCalls(calls)
  assert.equal(withStatus(updates, "in_progress").length, 1, "claimed once")
  assert.deepEqual(withStatus(updates, "todo"), [], "a worker that ran is not a worker that never started: no unclaim")
  assert.equal(ran.length, 1, "one worker ran")
  assert.ok(samePath(ran[0].cwd, lib.scratchTaskDir(task.id)), "in the scratch directory")
  assert.equal(existsSync(worktreePath(task.id)), false, "no worktree for a task that names no repo")
})

test("a repo-less task whose worker cannot launch is unclaimed with the gave-up comment", async () => {
  const task = { id: "task-scratch-broken", title: "Scratch work whose worker cannot start", repo: "", status: "todo" }
  // A symlink to a file that is not there: spawn answers ENOENT, which is the
  // missing-binary case, not a worker that ran and failed.
  pointPig(join(bin, "worker-that-is-not-there"))
  const result = await dispatch([task]).finally(() => pointPig(worker))
  const updates = updateCalls(result.calls)
  const claim = updates.findIndex((call) => argumentsOf(call).status === "in_progress")
  const unclaim = updates.findIndex((call) => argumentsOf(call).status === "todo")
  assert.ok(claim >= 0, "the task was claimed")
  assert.ok(unclaim > claim, "and handed back to todo after the claim, so a retry is possible")
  assert.match(String(argumentsOf(updates[unclaim]).comment), /^Dispatch gave up before starting a worker: no worker started/)
  assert.deepEqual(result.runs, [], "no worker ran")
  assert.equal(existsSync(worktreePath(task.id)), false, "and no worktree was made")
})

// The #32 defect the ordering fix closed: with the sandbox required and bwrap
// unavailable, the refusal came after the claim and after the worktree, so the
// task sat in_progress with no worker and could never be retried.
test("AGENT_OFFICE_SANDBOX=1 without bwrap refuses before the claim and before any worktree", async () => {
  const task = { id: "task-sandbox-refused", title: "Sandboxed work on a host without bwrap", repo: FIXTURE_REPO, status: "todo" }
  const previousPath = process.env.PATH ?? ""
  process.env.PATH = `${noBwrap}:${previousPath}`
  const result = await dispatch([task], { sandbox: "1" }).finally(() => {
    process.env.PATH = previousPath
  })
  assert.deepEqual(claimCalls(result.calls), [], "no claim at all")
  assert.deepEqual(updateCalls(result.calls), [], "and no status write of any kind")
  assert.deepEqual(result.runs, [], "no worker ran")
  assert.equal(existsSync(worktreePath(task.id)), false, "and no worktree was made")
  assert.ok(
    result.notices.some((notice) => /AGENT_OFFICE_SANDBOX requires bwrap/.test(notice.text)),
    "the refusal is reported",
  )
})

test("a task cancelled between the poll and the read is not claimed", async () => {
  const task = { id: "task-cancelled", title: "Work someone already cancelled", repo: "", status: "todo" }
  const { calls, notices: said, runs: ran } = await dispatch([task], { read: { [task.id]: { ...task, status: "cancelled" } } })
  assert.deepEqual(claimCalls(calls), [], "no claim")
  assert.deepEqual(updateCalls(calls), [], "and no status write")
  assert.deepEqual(ran, [], "no worker")
  assert.ok(said.some((notice) => /is cancelled, not todo/.test(notice.text)), "the refusal names the state")
  assert.equal(existsSync(worktreePath(task.id)), false, "and no worktree was made")
})

// ao-pi#4: four copies of one issue are four ids, and used to be four workers.
test("two copies of one issue get one worker", async () => {
  const first = { id: "task-copy-a", title: "Composer pin follow-ups", repo: "", status: "todo" }
  const second = { id: "task-copy-b", title: "  composer   PIN follow-ups ", repo: "", status: "todo" }
  // The first worker stays up long enough that the second copy is considered
  // while its work scope is genuinely held, rather than after it finished.
  process.env.STUB_SLEEP_MS = "400"
  const { calls, notices: said, runs: ran } = await dispatch([first, second]).finally(() => {
    delete process.env.STUB_SLEEP_MS
  })
  assert.equal(ran.length, 1, "one worker for two copies")
  assert.equal(toolNamed("accept_task", calls).length, 1, "one claim")
  assert.equal(updateCalls(calls).length, 1, "one status write")
  assert.equal(withStatus(updateCalls(calls), "in_progress").length, 1)
  assert.ok(
    said.some((notice) => /another copy of .* is already running/.test(notice.text)),
    "the skipped copy says why it was skipped",
  )
})

test("a worktree task whose worker runs is not unclaimed", async () => {
  const task = { id: "task-worktree-ok", title: "Work that needs the repository", repo: FIXTURE_REPO, status: "todo" }
  const { calls, runs: ran } = await dispatch([task])
  const updates = updateCalls(calls)
  assert.equal(withStatus(updates, "in_progress").length, 1, "claimed")
  assert.deepEqual(withStatus(updates, "todo"), [], "no unclaim")
  const path = worktreePath(task.id)
  assert.equal(existsSync(path), true, "the dispatch made a worktree")
  assert.equal(ran.length, 1, "one worker ran")
  assert.ok(samePath(ran[0].cwd, path), "in the worktree, not the session's checkout")
})

// #33: a cancellation has to win over a dispatch already in flight. The task is
// read before the claim and the worker can run for ten minutes after it, so the
// dispatch re-reads the task while the worker runs and stops it when the Office
// has cancelled it. Nothing is unclaimed: the worker did start, and the Office
// owns the status the task is in now.
test("a task cancelled while its worker runs stops that worker, and unclaims nothing", async () => {
  const task = { id: "task-cancelled-midrun", title: "Work cancelled while it runs", repo: FIXTURE_REPO, status: "todo" }
  // Long enough that the watch ticks (100 ms here) while the worker is still up.
  process.env.STUB_SLEEP_MS = "3000"
  const result = await dispatch([task], {
    duringRun: () => {
      office.state.tasks[task.id].status = "cancelled"
    },
  }).finally(() => {
    delete process.env.STUB_SLEEP_MS
  })
  assert.equal(result.runs.length, 1, "the worker started")
  assert.equal(completed(), 0, "and was stopped, not left to finish its work")
  assert.ok(
    result.notices.some((notice) => new RegExp(`task ${task.id} is cancelled; stopping its worker`).test(notice.text)),
    "the dispatch says it stopped the worker",
  )
  const updates = updateCalls(result.calls)
  assert.equal(withStatus(updates, "in_progress").length, 1, "claimed")
  assert.deepEqual(withStatus(updates, "todo"), [], "a worker that was stopped is not unclaimed")
  assert.equal(existsSync(worktreePath(task.id)), true, "and its worktree is kept")
})
