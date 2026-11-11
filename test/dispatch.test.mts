import assert from "node:assert/strict"
import { spawn } from "node:child_process"
import { createHash } from "node:crypto"
import { mkdirSync, readdirSync, utimesSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { test } from "node:test"
import { fileURLToPath } from "node:url"
import { acquireDispatch, DISPATCH_STALE_MS, MAX_DISPATCHES, releaseDispatch } from "../lib/office.ts"
import { ownsTempDir } from "./tmp.mts"

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

test("the dispatch cap holds across processes", async (t) => {
  const dir = join(ownsTempDir(t, "dispatch"), "locks")
  const at = Date.now() + 1_500
  const results = await Promise.all(Array.from({ length: 12 }, (_, index) => race(dir, `task-${index}`, at)))
  const started = results.filter((result) => result === "ok").length
  assert.ok(started >= 1, "at least one task starts")
  assert.ok(started <= MAX_DISPATCHES, `${started} tasks started, cap is ${MAX_DISPATCHES}`)
  assert.ok(readdirSync(dir).filter((file) => file.endsWith(".lock")).length <= MAX_DISPATCHES)
  assert.equal(readdirSync(dir).includes(".acquire"), false, "the mutex is released")
})

test("a stale lock for the same task is reclaimed", (t) => {
  const dir = join(ownsTempDir(t, "lock"), "locks")
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

test("a fresh lock for the same task stays busy", (t) => {
  const dir = join(ownsTempDir(t, "lock"), "locks")
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

// agent-office#540: two agents can be handed the same task. Without the owner
// in the name they compute the same worktree path and branch, and the second
// one resumes the first one's tree.
test("two agents given the same task get different worktrees", async (t) => {
  const { addWorktree, worktreeSlug } = await import("../lib/office.ts")
  const repo = join(ownsTempDir(t, "wt"), "ao-pi")
  const { spawnSync } = await import("node:child_process")
  spawnSync("git", ["init", "-q", "-b", "main", repo])
  // A branch to hang a worktree off: git will not add one to an empty repo.
  spawnSync("git", ["-C", repo, "-c", "user.email=t@example.com", "-c", "user.name=t", "commit", "-q", "--allow-empty", "-m", "init"])

  const a = await addWorktree(repo, "94b93cab", "ageo")
  const b = await addWorktree(repo, "94b93cab", "vinn")
  assert.notEqual(a.path, b.path, "different owners, different paths")
  assert.notEqual(a.branch, b.branch, "different owners, different branches")
  assert.match(a.branch, /^office\/ageo-/)
  assert.match(b.branch, /^office\/vinn-/)

  // The same owner retrying resumes its own worktree rather than making another.
  const again = await addWorktree(repo, "94b93cab", "vinn")
  assert.equal(again.path, b.path)

  // An owner-less call keeps the old shape, so an existing tree from an earlier
  // version is still found instead of duplicated.
  const legacy = await addWorktree(repo, "94b93cab")
  assert.doesNotMatch(legacy.branch, /^office\/(ageo|vinn)-/)
  assert.equal(worktreeSlug("  Vinn / Office  "), "vinn-office")
})
