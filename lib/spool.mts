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
// one are exactly the new usage. That cursor lives in the spool, not only in the
// process, because a resumed session (`pi --session <id>`) is a new process with
// no memory of what the last one spooled: read from the beginning it would take
// the newest entry as its key and store the whole session a second time, which
// Office cannot tell from new spend. A per-session `.cursor-<hash>` file keeps the
// last spooled entry id, and the key stays deterministic so a range offered twice
// is the same report rather than a second one.
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs"
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
 * cursorFileName is where a session's cursor lives: one dot-prefixed file per
 * session id, so the parent's reader - which takes `<hash>.json` reports - never
 * mistakes it for one, and two children sharing a spool directory never share a
 * cursor.
 */
export function cursorFileName(sessionID: string): string {
  return `.cursor-${digest(sessionID.trim())}`
}

// writeAtomic writes one small file the way the spool's readers expect it: a 0600
// temporary inside the directory renamed to its 0600 name, so nobody reads a half
// written file. The directory is created 0700 and nothing is written or followed
// outside it.
function writeAtomic(dir: string, name: string, content: string): void {
  mkdirSync(dir, { recursive: true, mode: 0o700 })
  const path = join(dir, name)
  const temporary = join(dir, `.${name}.tmp`)
  try {
    writeFileSync(temporary, content, { mode: 0o600 })
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
}

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
 * readSpoolCursor answers the entry id this session's last report ended at, as the
 * spool records it. An absent file is a first settle, not a failure; anything else
 * - a cursor that cannot be read, or one that is not an id - is reported so the
 * caller can say so once, and the range then starts from the beginning.
 */
export function readSpoolCursor(dir: string, sessionID: string): { cursor?: string; error?: string } {
  let text: string
  try {
    text = readFileSync(join(dir, cursorFileName(sessionID)), "utf8")
  } catch (err) {
    if ((err as NodeJS.ErrnoException)?.code === "ENOENT") return {}
    return { error: String(err) }
  }
  const cursor = text.trim()
  return cursor ? { cursor } : {}
}

/** writeSpoolCursor records the entry id the next report starts after. */
export function writeSpoolCursor(dir: string, sessionID: string, entryID: string): void {
  writeAtomic(dir, cursorFileName(sessionID), `${entryID}\n`)
}

/**
 * spoolRange answers the entries recorded since `cursor`: the ones after the
 * entry Pi last spooled for this session. A cursor the list does not hold - a
 * session rewound past it - starts from the beginning, which is why the cursor is
 * persisted and the key is deterministic. `last` is the id of the last entry
 * included, which the key names; an entry list with no ids at all has none, and
 * is not spooled.
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
  writeAtomic(dir, name, `${JSON.stringify(body)}\n`)
  return { written: true, path }
}

/**
 * spoolUsage spools what the session recorded since `cursor`, and answers the
 * cursor the next settle starts from.
 *
 * `cursor` is this process's memory; with none, the spool's own cursor for the
 * session is the starting point, which is what stops a resumed session from
 * spooling its whole history a second time. The cursor advances only once the
 * range is queued - written now, or already there under the same key - so a
 * failed write is retried with the same range and the same key instead of being
 * skipped. A cursor that cannot be read is answered as `cursorError` and the
 * range then starts from the beginning, as it would without one.
 */
export function spoolUsage(options: {
  entries: readonly SpoolEntry[]
  cursor?: string
  sessionID: string
  taskID?: string
  dir: string
  now?: Date
}): { cursor?: string; written: boolean; path?: string; cursorError?: string } {
  const saved: { cursor?: string; error?: string } = options.cursor
    ? { cursor: options.cursor }
    : readSpoolCursor(options.dir, options.sessionID)
  const range = spoolRange(options.entries, saved.cursor)
  if (!range.last) return { cursor: saved.cursor, written: false, cursorError: saved.error }
  const summary = summarizeUsage(range.entries)
  // Nothing recorded, or no model Pi could name: Office would refuse a report
  // with no count to price, and the entries stay in the next range.
  if (!reportable(summary)) return { cursor: saved.cursor, written: false, cursorError: saved.error }
  const key = spoolKey(options.sessionID, range.last)
  const day = dayOf(undefined, options.now ?? new Date())
  const queued = writeSpool(options.dir, key, spoolBody({ summary, key, day, taskID: options.taskID }))
  // Only now, with the range queued: a process that dies before this file is
  // written retries the same range and the same key, and a new process reads the
  // last one that did land rather than the whole session.
  writeSpoolCursor(options.dir, options.sessionID, range.last)
  return { cursor: range.last, written: queued.written, path: queued.path, cursorError: saved.error }
}
