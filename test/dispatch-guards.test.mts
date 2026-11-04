// ao-pi#4: auto-dispatch started a worker for a task that had just been
// cancelled, and one worker per copy of an issue. Both defects were the same
// shape — a guard that existed but was not in the path — so these tests pin the
// logic *and* the wiring.
import assert from "node:assert/strict"
import { mkdirSync, readFileSync, realpathSync, statSync, utimesSync, writeFileSync } from "node:fs"
import { spawnSync } from "node:child_process"
import { mkdtempSync } from "node:fs"
import { homedir, tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { acquireWorkScope, DISPATCH_STALE_MS, releaseWorkScope, sandboxArgs, sandboxAvailable, sandboxDecision, stillDispatchable, unclaimTask, workScopeKey, workerSandboxPaths } from "../lib/office.ts"

const lockDir = () => join(mkdtempSync(join(tmpdir(), "ao-scope-")), "locks")

test("a task that stopped being todo is not dispatched", () => {
  // The predicate has existed since agent-office#540 with a test of its own; what
  // this one adds is the wiring check below.
  assert.deepEqual(stillDispatchable({ status: "todo" }), { ok: true })
  for (const status of ["cancelled", "done", "blocked", "in_progress", "review", "CANCELLED"]) {
    assert.equal(stillDispatchable({ status }).ok, false, `${status} must not dispatch`)
  }
  // An absent status reads as todo, as the poller's filter does.
  assert.equal(stillDispatchable({ id: "t" }).ok, true)
})

test("the dispatch path asks the predicate before claiming, and holds a scope", () => {
  // The defect was a missing call, not a missing function: assert the call sites
  // inside dispatchTask rather than trusting the import.
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8")
  const start = source.indexOf("const dispatchTask = async (id: string) => {")
  const end = source.indexOf("const pumpDispatch = () => {")
  assert.ok(start > 0 && end > start, "dispatchTask is where these tests think it is")
  const body = source.slice(start, end)
  const verdictAt = body.indexOf("stillDispatchable(task)")
  const claimAt = body.indexOf("await claimTask(call, id)")
  assert.ok(verdictAt > 0, "dispatchTask consults stillDispatchable")
  assert.ok(claimAt > verdictAt, "the verdict comes before the claim that flips the task to in_progress")
  assert.ok(body.indexOf("acquireWorkScope(dispatchDir(), scope)") > 0, "dispatchTask takes a work scope")
  assert.ok(body.includes("releaseWorkScope(dispatchDir(), heldScope)"), "and releases it when the dispatch ends")
})

test("copies of one issue share a scope; different work does not", () => {
  const a = workScopeKey({ title: "Composer pin follow-ups", repo: "" })
  const b = workScopeKey({ title: "  composer   PIN follow-ups ", repo: "" })
  assert.equal(a, b, "the same title in different case and spacing is the same work")
  assert.notEqual(a, workScopeKey({ title: "Composer pin follow-ups", repo: "agent-office" }), "a different repo is different work")
  assert.notEqual(a, workScopeKey({ title: "Something else", repo: "" }))
  assert.equal(workScopeKey({ title: "   " }), undefined, "a task with no title cannot be scoped")
  assert.equal(workScopeKey({}), undefined)
})

test("one worker per scope, and a stale scope is reclaimed", () => {
  const dir = lockDir()
  const key = workScopeKey({ title: "Composer pin follow-ups" })!
  assert.equal(acquireWorkScope(dir, key), true, "the first copy dispatches")
  assert.equal(acquireWorkScope(dir, key), false, "the second copy waits for the first")
  releaseWorkScope(dir, key)
  assert.equal(acquireWorkScope(dir, key), true, "a released scope can be taken again")

  // A scope left behind by a killed worker must not block the work for ever.
  releaseWorkScope(dir, key)
  const path = join(dir, "work", `${key}.lock`)
  writeFileSync(path, JSON.stringify({ scope: key, pid: 999_999, created: 0 }))
  const old = (Date.now() - DISPATCH_STALE_MS - 60_000) / 1000
  utimesSync(path, old, old)
  assert.equal(acquireWorkScope(dir, key), true, "a scope older than the dispatch bound is reclaimed")
  assert.ok(statSync(path).mtimeMs > old * 1000)
})

test("scope locks do not consume a dispatch slot", async () => {
  // acquireDispatch counts *.lock in the directory itself; the scope lives in a
  // subdirectory, so holding many scopes cannot shrink the worker cap.
  const { acquireDispatch, MAX_DISPATCHES, releaseDispatch } = await import("../lib/office.ts")
  const dir = lockDir()
  mkdirSync(dir, { recursive: true })
  for (let i = 0; i < MAX_DISPATCHES; i++) {
    assert.equal(acquireWorkScope(dir, `scope-${i}`), true)
  }
  assert.equal(acquireDispatch(dir, "a-real-task"), "ok", "slots are still free with every scope held")
  releaseDispatch(dir, "a-real-task")
})

// ao-pi#4: the write into another task's worktree was defended by instruction
// only. The sandbox is the enforcement that exists; making it opt-in meant the
// default was the unenforced path.
test("a worker is sandboxed where it can be, opted out of deliberately, and required on demand", () => {
  // Default: contained wherever bwrap works, and honest about it where it does not.
  assert.deepEqual(sandboxDecision(undefined, true), { on: true, required: false })
  assert.deepEqual(sandboxDecision("", true), { on: true, required: false })
  assert.deepEqual(sandboxDecision(undefined, false), { on: false, required: false })

  // Opting out is explicit and never required.
  for (const off of ["0", "off", "FALSE", " no "]) {
    assert.deepEqual(sandboxDecision(off, true), { on: false, required: false }, off)
  }

  // Asking for it makes it a requirement: unavailable means refuse, not run loose.
  for (const on of ["1", "on", "yes", " TRUE "]) {
    assert.deepEqual(sandboxDecision(on, true), { on: true, required: true }, on)
    assert.deepEqual(sandboxDecision(on, false), { on: false, required: true }, on)
  }
})

// rodri's must-fix on ao-pi#4: a `=1` refusal has to happen before the claim and
// before any git work, or it leaves the task flipped to in_progress and a worktree
// behind - the litter the refusal exists to avoid.
test("the sandbox refusal comes before the claim and the worktree", () => {
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8")
  const start = source.indexOf("const dispatchTask = async (id: string) => {")
  const end = source.indexOf("const pumpDispatch = () => {")
  const body = source.slice(start, end)
  const decideAt = body.indexOf("sandboxDecision(")
  const claimAt = body.indexOf("await claimTask(call, id)")
  const worktreeAt = body.indexOf("await addWorktree(")
  assert.ok(decideAt > 0 && claimAt > decideAt, "the sandbox decision comes before the claim")
  assert.ok(worktreeAt > decideAt, "and before any worktree is made")
})

// The other half: default-on must not break a worker's token refresh. The
// credentials file is written beside itself (temp file plus rename), so the
// directory has to be a writable bind inside the sandbox.
test("the sandbox makes the worker's credentials directory writable", () => {
  const root = mkdtempSync(join(tmpdir(), "ao-cred-bind-"))
  const agents = join(root, "agents")
  mkdirSync(agents, { recursive: true })
  const worktree = join(root, "worktree")
  mkdirSync(worktree, { recursive: true })
  const credentials = join(agents, "agent.json")
  writeFileSync(credentials, "{}")
  const paths = workerSandboxPaths(worktree, credentials)
  assert.equal(paths.credentials, realpathSync(agents), "the bind is the file's own directory")
  const args = sandboxArgs(paths)
  const writable = args.reduce<string[]>((out, arg, i) => (arg === "--bind" ? [...out, args[i + 1]] : out), [])
  assert.ok(writable.includes(paths.credentials!), `credentials dir is writable, got ${JSON.stringify(writable)}`)
  // No credentials file named, no extra bind: the tight default stays tight.
  assert.equal(workerSandboxPaths(worktree).credentials, undefined)
})

// And the claim above, for real: a sandboxed process writes the credentials file
// where the worker would, and is refused outside its mounts.
test("a real sandboxed write to the credentials directory succeeds, and outside it does not", { skip: !sandboxAvailable() }, () => {
  const root = mkdtempSync(join(tmpdir(), "ao-sandbox-run-"))
  const agents = join(root, "agents")
  const outside = join(root, "outside")
  mkdirSync(agents, { recursive: true })
  mkdirSync(outside, { recursive: true })
  // sandboxArgs always binds the worktree and its git dir: give the fake ones.
  mkdirSync(join(root, ".git"), { recursive: true })
  writeFileSync(join(agents, "agent.json"), JSON.stringify({ access_token: "old" }))
  const paths = workerSandboxPaths(root, join(agents, "agent.json"))
  const args = sandboxArgs(paths)

  const write = (target: string) =>
    spawnSync("bwrap", [...args, "/bin/sh", "-c", `printf new > ${target}`], { encoding: "utf8" })

  const inside = write(join(agents, "agent.json"))
  assert.equal(inside.status, 0, `writing the credentials file inside the sandbox failed: ${inside.stderr}`)
  assert.equal(readFileSync(join(agents, "agent.json"), "utf8"), "new", "the refresh reached the real file")

  const escaped = write(outside)
  assert.notEqual(escaped.status, 0, "a sandboxed worker must not write outside its mounts")
})

// rodri on ao-pi#4: every refusal after the claim used to strand the task. A
// dispatch that claimed and did not start a worker now hands it back.
test("a dispatch that gives up after claiming puts the task back to todo", async () => {
  const calls: Array<{ name: string; args: any }> = []
  const call = async (name: string, args?: any) => { calls.push({ name, args }); return {} }
  await unclaimTask(call, "task-1", "the repo did not resolve")
  assert.equal(calls.length, 1)
  assert.equal(calls[0].name, "update_task")
  assert.equal(calls[0].args.task_id, "task-1")
  assert.equal(calls[0].args.status, "todo", "back to the queue, not left in_progress")
  assert.match(String(calls[0].args.comment), /Dispatch gave up before starting a worker: the repo did not resolve/)

  // A failed update must not turn the dispatch's own error into a second one.
  await unclaimTask(async () => { throw new Error("office down") }, "task-2", "why")
})

test("the claim release runs only when no worker started", () => {
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8")
  const start = source.indexOf("const dispatchTask = async (id: string) => {")
  const body = source.slice(start, source.indexOf("const pumpDispatch = () => {"))
  assert.match(body, /if \(claimed && !workerStarted\) await unclaimTask\(call, id, /, "the release is guarded by both flags")
  assert.ok(body.indexOf("claimed = true") < body.indexOf("await claimTask(call, id)"), "claimed is set as the claim is taken")
})
