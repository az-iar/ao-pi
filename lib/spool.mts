// A child session's token usage, spooled for its parent to report (Office task
// dd185d2b).
//
// A Pi started as a child - `hc run` under a Claude Code session, with
// AGENT_OFFICE_CHILD=1 - has no Office identity of its own, so it cannot post its
// spend and must not pretend to (ao-pi#476): the request would speak as whoever
// approved the host. The agent that started it can post, so the child leaves the
// counts somewhere the parent's ao-claude plugin reads: one file per settle, in a
// directory the parent names in AGENT_OFFICE_USAGE_SPOOL.
//
// Everything here is pure or filesystem-only, so the range, the key, the file
// name and the body are unit-testable without a running Pi. The report itself is
// the same shape `usageReport` posts, priced by Office the same way: no `amount`.
//
// A long-lived child settles many times, and each file carries only what was
// recorded since the last one - the count a settle reports is the count that
// settle added, so the files sum to the run. The range is chosen by session entry
// id: a session keeps its entries in order, so the entries after the last spooled
// one are exactly the new usage. The key names the last entry it includes, which
// is what makes a re-spool after a resume safe: the same range produces the same
// key, the same file name, and Office stores nothing new.
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, renameSync, unlinkSync, writeFileSync } from "node:fs"
import { isAbsolute, join } from "node:path"
import { dayOf, reportable, summarizeUsage, type UsageEntry, type UsageSummary } from "./usage.mts"

/** A session entry as this file needs it: usage for the summary, an id for the range. */
export type SpoolEntry = UsageEntry & { id?: string }

/** SpoolBody is the document one file holds: usageReport's body, task_id only when there is one. */
export type SpoolBody = {
  source: "harness:pi"
  day: string
  model: string
  tokens: { input: number; output: number; cache_read: number; cache_write: number }
  task_id?: string
  idempotency_key: string
}

// Office refuses an unverifiable task id, so the key the contract fixes has a
// ceiling; a longer one is replaced by its own hash rather than truncated, which
// could collide two different ranges onto one name.
const KEY_MAX = 128
const HASH_CHARS = 40

const digest = (value: string) => createHash("sha256").update(value).digest("hex").slice(0, HASH_CHARS)

/**
 * spoolDirFor answers the directory this child spools into, or undefined when
 * spooling is off: the child marker alone leaves today's behaviour alone, and so
 * does a relative or empty path, because the child's working directory is not
 * where its parent looks.
 */
export function spoolDirFor(env: Record<string, string | undefined>): string | undefined {
  if (env.AGENT_OFFICE_CHILD !== "1") return undefined
  const dir = env.AGENT_OFFICE_USAGE_SPOOL?.trim()
  if (!dir || !isAbsolute(dir)) return undefined
  return dir
}

/**
 * spoolRange answers the entries recorded since `cursor`: the ones after the
 * entry Pi last spooled for this session. A cursor that is not in the list (a
 * resume starts a new entry list, and an old cursor names nothing in it) starts
 * from the beginning, which spoolKey makes harmless. `last` is the id of the
 * last entry included, which the key names; an entry list with no ids at all has
 * none, and is not spooled.
 */
export function spoolRange(entries: readonly SpoolEntry[], cursor?: string): { entries: SpoolEntry[]; last?: string } {
  const list = entries ?? []
  const at = cursor ? list.findIndex((entry) => entry?.id === cursor) : -1
  const range = at >= 0 ? list.slice(at + 1) : list.slice(0)
  let last: string | undefined
  for (const entry of range) if (typeof entry?.id === "string" && entry.id) last = entry.id
  return { entries: range, last }
}

/** spoolKey is the idempotency key of the range that ends at `lastEntryID`. */
export function spoolKey(sessionID: string, lastEntryID: string): string {
  const key = `pi:${sessionID.trim()}:${lastEntryID.trim()}`
  return key.length <= KEY_MAX ? key : `pi:${digest(key)}`
}

/** spoolFileName is the file a key is queued under; the key itself never becomes a path. */
export function spoolFileName(key: string): string {
  return `${digest(key)}.json`
}

/**
 * spoolBody is one file's content: the body usageReport posts, with `task_id`
 * only when the dispatcher named a task - a child that was not started for one
 * has nothing to attribute the spend to. No `amount`: Office prices tokens.
 */
export function spoolBody(options: { summary: UsageSummary; key: string; day: string; taskID?: string }): SpoolBody {
  const body: SpoolBody = {
    source: "harness:pi",
    day: options.day,
    model: options.summary.model.trim(),
    tokens: {
      input: options.summary.tokens.input,
      output: options.summary.tokens.output,
      cache_read: options.summary.tokens.cacheRead,
      cache_write: options.summary.tokens.cacheWrite,
    },
    idempotency_key: options.key,
  }
  const taskID = options.taskID?.trim()
  if (taskID) body.task_id = taskID
  return body
}

/**
 * writeSpool queues one report: a 0600 temporary file inside the spool directory
 * renamed to its final 0600 name, so a parent that reads while the child writes
 * never sees half a report. An existing file of that name means the same report
 * is already queued and is left alone. The directory is created 0700, and
 * nothing is written or followed outside it.
 */
export function writeSpool(dir: string, key: string, body: SpoolBody): { written: boolean; path: string } {
  const name = spoolFileName(key)
  const path = join(dir, name)
  if (existsSync(path)) return { written: false, path }
  mkdirSync(dir, { recursive: true, mode: 0o700 })
  const temporary = join(dir, `.${name}.tmp`)
  try {
    writeFileSync(temporary, `${JSON.stringify(body)}\n`, { mode: 0o600 })
    renameSync(temporary, path)
  } catch (err) {
    try {
      unlinkSync(temporary)
    } catch {
      // The temporary file is inside the spool directory and a later write
      // replaces it; the original failure is the one worth reporting.
    }
    throw err
  }
  return { written: true, path }
}

/**
 * spoolUsage spools what the session recorded since `cursor`, and answers the
 * cursor the next settle starts from. The cursor advances only once the range is
 * queued - written now, or already there under the same key - so a failed write
 * is retried with the same range and the same key instead of being skipped.
 */
export function spoolUsage(options: {
  entries: readonly SpoolEntry[]
  cursor?: string
  sessionID: string
  taskID?: string
  dir: string
  now?: Date
}): { cursor?: string; written: boolean; path?: string } {
  const range = spoolRange(options.entries, options.cursor)
  if (!range.last) return { cursor: options.cursor, written: false }
  const summary = summarizeUsage(range.entries)
  // Nothing recorded, or no model Pi could name: Office would refuse a report
  // with no count to price, and the entries stay in the next range.
  if (!reportable(summary)) return { cursor: options.cursor, written: false }
  const key = spoolKey(options.sessionID, range.last)
  const day = dayOf(undefined, options.now ?? new Date())
  const queued = writeSpool(options.dir, key, spoolBody({ summary, key, day, taskID: options.taskID }))
  return { cursor: range.last, written: queued.written, path: queued.path }
}
