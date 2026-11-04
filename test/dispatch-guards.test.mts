// ao-pi#4: auto-dispatch started a worker for a task that had just been
// cancelled, and one worker per copy of an issue. Both defects were the same
// shape — a guard that existed but was not in the path — so these tests pin the
// logic *and* the wiring.
import assert from "node:assert/strict"
import { mkdirSync, readFileSync, statSync, utimesSync, writeFileSync } from "node:fs"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { acquireWorkScope, DISPATCH_STALE_MS, releaseWorkScope, sandboxDecision, stillDispatchable, workScopeKey } from "../lib/office.ts"

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
