// The spool contract for a child's token usage, without Pi (Office dd185d2b).
//
// A Pi started by `hc run` under a Claude Code session is a child (`AGENT_OFFICE_CHILD=1`)
// with no Office identity of its own, so it does not post its spend: with
// AGENT_OFFICE_USAGE_SPOOL it queues what each settle recorded for the parent's
// ao-claude plugin to post as the agent that started it. These are the rules the
// parent's reader relies on - which entries are new, the key, the file name, the
// body - and they are all pure or filesystem-only, so no Office and no Pi is
// needed to pin them. test/child-spool.test.mts drives the real extension.
import assert from "node:assert/strict"
import { createHash } from "node:crypto"
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs"
import { join } from "node:path"
import { test } from "node:test"
import { spoolBody, spoolDirFor, spoolFileName, spoolKey, spoolRange, spoolUsage, writeSpool } from "../lib/spool.mts"
import { summarizeUsage } from "../lib/usage.mts"
import { tempDir } from "./tmp.mts"

type Entry = { id?: string; type?: string; message?: { role?: string; model?: string; usage?: Record<string, number> }; usage?: Record<string, number> }
const usage = (input: number, output: number, cacheRead: number, cacheWrite: number) => ({ input, output, cacheRead, cacheWrite, totalTokens: input + output + cacheRead + cacheWrite, cost: { total: 0 } })
const assistant = (id: string, model: string, tokens: Record<string, number>): Entry => ({ id, type: "message", message: { role: "assistant", model, usage: tokens } })

const sha = (value: string) => createHash("sha256").update(value).digest("hex").slice(0, 40)
const spoolFiles = (dir: string) => (existsSync(dir) ? readdirSync(dir).filter((name) => name.endsWith(".json")) : [])
const bodyOf = (path: string) => JSON.parse(readFileSync(path, "utf8"))

test("the range is the entries after the cursor; an unknown cursor starts from the beginning", () => {
  const entries = [assistant("e1", "m", usage(1, 0, 0, 0)), assistant("e2", "m", usage(2, 0, 0, 0)), assistant("e3", "m", usage(3, 0, 0, 0))]
  assert.deepEqual(spoolRange(entries, "e1").entries.map((entry) => entry.id), ["e2", "e3"], "after e1")
  assert.equal(spoolRange(entries, "e1").last, "e3", "the key names the last entry included")
  assert.deepEqual(spoolRange(entries).entries.map((entry) => entry.id), ["e1", "e2", "e3"], "no cursor is the whole list")
  assert.deepEqual(spoolRange(entries, "gone").entries.map((entry) => entry.id), ["e1", "e2", "e3"], "a cursor a resumed session does not hold starts from the beginning")
  assert.deepEqual(spoolRange(entries, "e3").entries, [], "nothing after the last entry")
  assert.equal(spoolRange(entries, "e3").last, undefined, "an empty range names no entry")
})

test("nothing recorded since the cursor writes no file", (t) => {
  const dir = join(tempDir("spool-pure"), "spool")
  const entries = [assistant("e1", "m", usage(10, 1, 0, 0))]
  const first = spoolUsage({ entries, sessionID: "s1", dir, now: new Date("2026-11-18T10:00:00.000Z") })
  assert.equal(first.written, true)
  assert.equal(first.cursor, "e1")
  const again = spoolUsage({ entries, cursor: "e1", sessionID: "s1", dir })
  assert.equal(again.written, false, "a settle with no new usage writes nothing")
  assert.equal(again.cursor, "e1", "and leaves the cursor where it was")
  assert.equal(spoolFiles(dir).length, 1)
})

test("the same range spooled twice is one file, keyed by the entry it ends at", (t) => {
  const dir = join(tempDir("spool-pure"), "spool")
  const entries = [assistant("e1", "m", usage(10, 1, 0, 0)), assistant("e2", "m", usage(20, 2, 0, 0))]
  // A resume starts a new entry list, so the cursor is lost and the whole session
  // is offered again: the second spool has to be the same report, not a second one.
  const first = spoolUsage({ entries, sessionID: "s1", dir })
  const second = spoolUsage({ entries, sessionID: "s1", dir })
  assert.equal(first.written, true)
  assert.equal(second.written, false, "the same key is already queued")
  assert.equal(first.path, second.path)
  assert.equal(spoolFiles(dir).length, 1)
  assert.equal(bodyOf(first.path!).idempotency_key, spoolKey("s1", "e2"), "the key names the entry the range ends at")
})

test("a range with no usage, or no model, writes nothing and keeps the entries for later", (t) => {
  const dir = join(tempDir("spool-pure"), "spool")
  const empty = spoolUsage({ entries: [{ id: "e1", type: "message", message: { role: "user" } }], sessionID: "s1", dir })
  assert.equal(empty.written, false)
  assert.equal(empty.cursor, undefined, "the cursor stays so the entries are counted once there is usage")
  const badModel = spoolUsage({ entries: [{ id: "e2", type: "usage", usage: usage(5, 0, 0, 0) }], sessionID: "s1", dir })
  assert.equal(badModel.written, false, "Office cannot price tokens against no model")
  assert.deepEqual(spoolFiles(dir), [])
})

test("the body is usageReport's shape, with a task only when there is one, and never an amount", () => {
  const summary = summarizeUsage([assistant("e1", "claude-sonnet-5-5", usage(1200, 300, 4000, 50))])
  const withTask = spoolBody({ summary, key: spoolKey("s1", "e1"), day: "2026-11-18", taskID: "3f2a1a4e-9c1d-4f2b-8a55-0f3b7c9d1e21" })
  assert.deepEqual(withTask, {
    source: "harness:pi",
    day: "2026-11-18",
    model: "claude-sonnet-5-5",
    tokens: { input: 1200, output: 300, cache_read: 4000, cache_write: 50 },
    task_id: "3f2a1a4e-9c1d-4f2b-8a55-0f3b7c9d1e21",
    idempotency_key: "pi:s1:e1",
  })
  const withoutTask = spoolBody({ summary, key: "pi:s1:e1", day: "2026-11-18" })
  assert.equal("task_id" in withoutTask, false, "a child started for no task attributes its spend to none")
  assert.equal("amount" in withoutTask, false, "Office prices tokens; this side never sends a rate")
  assert.deepEqual(Object.keys(withoutTask).sort(), ["day", "idempotency_key", "model", "source", "tokens"], "no other field")
})

test("the key is the session and the last entry, hashed when it would be too long", () => {
  assert.equal(spoolKey("session-1", "entry-1"), "pi:session-1:entry-1")
  const long = `pi:${"s".repeat(100)}:${"e".repeat(60)}`
  const key = spoolKey("s".repeat(100), "e".repeat(60))
  assert.equal(key, `pi:${sha(long)}`, "a key over the ceiling is its own hash, not a truncation")
  assert.equal(key.length, 43)
  assert.ok(key.length <= 128, "Office's key ceiling")
  assert.equal(spoolKey("session-1", "entry-1"), spoolKey("session-1", "entry-1"), "deterministic, so a replay is the same report")
  assert.notEqual(spoolKey("session-1", "entry-1"), spoolKey("session-1", "entry-2"))
})

test("the file name is the hash of the key, and an existing file is never overwritten", (t) => {
  const dir = join(tempDir("spool-pure"), "spool")
  assert.equal(spoolFileName("pi:s1:e1"), `${sha("pi:s1:e1")}.json`)
  const summary = summarizeUsage([assistant("e1", "m", usage(10, 1, 0, 0))])
  const first = writeSpool(dir, "pi:s1:e1", spoolBody({ summary, key: "pi:s1:e1", day: "2026-11-17" }))
  assert.equal(first.written, true)
  const second = writeSpool(dir, "pi:s1:e1", spoolBody({ summary, key: "pi:s1:e1", day: "2026-11-18" }))
  assert.equal(second.written, false, "the same report is already queued")
  assert.equal(bodyOf(first.path).day, "2026-11-17", "the queued report is left as it was")
  assert.deepEqual(spoolFiles(dir), [spoolFileName("pi:s1:e1")], "no temporary file was left behind")
  assert.equal(statSync(dir).mode & 0o777, 0o700, "the spool directory is the child's alone")
  assert.equal(statSync(first.path).mode & 0o777, 0o600, "a queued report is readable by its owner only")
})

test("spooling is on only for a child that was handed an absolute directory", () => {
  assert.equal(spoolDirFor({}), undefined)
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "1" }), undefined, "a child with no directory behaves as before")
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "1", AGENT_OFFICE_USAGE_SPOOL: "" }), undefined)
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "1", AGENT_OFFICE_USAGE_SPOOL: "   " }), undefined)
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "1", AGENT_OFFICE_USAGE_SPOOL: "spool/here" }), undefined, "a relative path is not where the parent looks")
  assert.equal(spoolDirFor({ AGENT_OFFICE_USAGE_SPOOL: "/tmp/spool" }), undefined, "not a child: today's behaviour")
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "0", AGENT_OFFICE_USAGE_SPOOL: "/tmp/spool" }), undefined)
  assert.equal(spoolDirFor({ AGENT_OFFICE_CHILD: "1", AGENT_OFFICE_USAGE_SPOOL: " /tmp/spool \n" }), "/tmp/spool")
})

test("a body is written as JSON, one file per report, with nothing else in the directory", (t) => {
  const dir = join(tempDir("spool-pure"), "spool")
  const entries = [assistant("e1", "claude-sonnet-5-5", usage(10, 1, 0, 0))]
  const first = spoolUsage({ entries, sessionID: "s1", dir, taskID: "task-1" })
  assert.deepEqual(JSON.parse(readFileSync(first.path!, "utf8")), spoolBody({ summary: summarizeUsage(entries), key: "pi:s1:e1", day: new Date().toISOString().slice(0, 10), taskID: "task-1" }))
  assert.deepEqual(readdirSync(dir), [spoolFileName("pi:s1:e1")], "only the report itself, not the temporary it was written as")
})
