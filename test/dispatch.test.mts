import assert from "node:assert/strict"
import { spawn } from "node:child_process"
import { createHash } from "node:crypto"
import { mkdirSync, mkdtempSync, readdirSync, utimesSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { fileURLToPath } from "node:url"
import { acquireDispatch, DISPATCH_STALE_MS, MAX_DISPATCHES, releaseDispatch } from "../lib/office.ts"

const childScript = fileURLToPath(new URL("./acquire-child.mts", import.meta.url))
const lockName = (id: string) => `${createHash("sha256").update(id).digest("hex").slice(0, 16)}.lock`

function race(dir: string, id: string, at: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", childScript, dir, id, String(at)], {
      stdio: ["ignore", "pipe", "pipe"],
    })
    let stdout = ""
    let stderr = ""
    child.stdout.on("data", (chunk) => (stdout += String(chunk)))
    child.stderr.on("data", (chunk) => (stderr += String(chunk)))
    child.on("error", reject)
    child.on("close", (code) => (code === 0 ? resolve(stdout.trim()) : reject(new Error(stderr || `exit ${code}`))))
  })
}

test("the dispatch cap holds across processes", async () => {
  const dir = join(mkdtempSync(join(tmpdir(), "ao-dispatch-")), "locks")
  const at = Date.now() + 1_500
  const results = await Promise.all(Array.from({ length: 12 }, (_, index) => race(dir, `task-${index}`, at)))
  const started = results.filter((result) => result === "ok").length
  assert.ok(started >= 1, "at least one task starts")
  assert.ok(started <= MAX_DISPATCHES, `${started} tasks started, cap is ${MAX_DISPATCHES}`)
  assert.ok(readdirSync(dir).filter((file) => file.endsWith(".lock")).length <= MAX_DISPATCHES)
  assert.equal(readdirSync(dir).includes(".acquire"), false, "the mutex is released")
})

test("a stale lock for the same task is reclaimed", () => {
  const dir = join(mkdtempSync(join(tmpdir(), "ao-lock-")), "locks")
  const id = "task-stale"
  const lock = join(dir, lockName(id))
  mkdirSync(dir, { recursive: true })
  writeFileSync(lock, JSON.stringify({ task: id, pid: 999_999, created: 0 }))
  const old = (Date.now() - DISPATCH_STALE_MS - 60_000) / 1000
  utimesSync(lock, old, old)
  assert.equal(acquireDispatch(dir, id), "ok")
  releaseDispatch(dir, id)
  assert.deepEqual(readdirSync(dir).filter((file) => file.endsWith(".lock")), [])
})

test("a fresh lock for the same task stays busy", () => {
  const dir = join(mkdtempSync(join(tmpdir(), "ao-lock-")), "locks")
  const id = "task-fresh"
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, lockName(id)), JSON.stringify({ task: id, pid: process.pid, created: Date.now() }))
  assert.equal(acquireDispatch(dir, id), "busy")
})

// agent-office#540: the poller's task list is a snapshot. A task that was todo
// when it was read can be cancelled by the time a dispatch slot frees, and
// claiming it starts a worker for work nobody wants.
test("a task that stopped being todo is not dispatched", async () => {
  const { stillDispatchable } = await import("../lib/office.ts")
  assert.deepEqual(stillDispatchable({ id: "t", status: "todo" }), { ok: true })
  // An absent status reads as todo, the way the poller's own filter reads it.
  assert.deepEqual(stillDispatchable({ id: "t" }), { ok: true })
  for (const status of ["cancelled", "done", "blocked", "in_progress", "CANCELLED"]) {
    const verdict = stillDispatchable({ id: "t", status })
    assert.equal(verdict.ok, false, `${status} must not dispatch`)
    assert.equal(verdict.reason, status.toLowerCase())
  }
})

// agent-office#540: a task from another project would be worked in a worktree of
// this repository. The payload names a project, not a repository, so the check
// is a comparison against the repo's own name — and an unnamed project is not a
// mismatch, or every task without one would stop dispatching.
test("a task naming another project is not dispatched here", async () => {
  const { projectMismatch } = await import("../lib/office.ts")
  const repo = "/home/azri/projects/personal/ao-pi"
  assert.equal(projectMismatch({ project: "ao-pi" }, repo), undefined, "the same project dispatches")
  assert.equal(projectMismatch({ project: "AO-PI" }, repo), undefined, "case is not a mismatch")
  assert.equal(projectMismatch({ project_slug: "ao-pi" }, repo), undefined, "the older field name works")
  assert.equal(projectMismatch({}, repo), undefined, "an unnamed project is not a mismatch")
  assert.equal(projectMismatch({ project: "" }, repo), undefined)
  assert.match(String(projectMismatch({ project: "agent-office" }, repo)), /agent-office.*ao-pi/)
})
