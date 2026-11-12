// A dispatched worker's token usage (Office task 1b8f755a), read from Pi's own
// record of the run and posted once it finishes.
//
// Nothing here estimates a count. Pi records usage on every assistant message
// (and on the nested model work a tool result can carry, a cache-warm entry, a
// compaction or a branch summary), and this sums exactly the entries
// `getSessionStats` sums, so the number reported is the number `/session` shows.
// A run Pi recorded no usage for reports nothing at all.
//
// The report is one request: POST /api/v1/agent/usage through the MCP
// `report_usage` tool (agent-office#843), which is that route's documented
// client. Its body is
//
//   { source: "harness:pi", day: "YYYY-MM-DD", model,
//     tokens: { input, output, cache_read, cache_write },
//     task_id, idempotency_key }
//
// with no `amount`: Office prices tokens from its model table (priced_by ao) and
// flags a model it has no price for as unpriced, rather than this side guessing
// a rate.
//
// The idempotency key is `<task id>:<run>`, where the run is the worker session
// Pi created for it: a retry of the same run reuses it and the Office stores
// nothing new (200 instead of 201), while a re-dispatch is a new session and so
// new usage. Office scopes the key per agent, so it has to be neither.

export type Tokens = { input: number; output: number; cacheRead: number; cacheWrite: number }
export type UsageSummary = { tokens: Tokens; total: number; model: string }

// The shapes read out of a session entry, structural so this file depends on no
// harness's exported types.
type Counts = { input?: unknown; output?: unknown; cacheRead?: unknown; cacheWrite?: unknown }
type Entry = { type?: string; usage?: Counts; message?: { role?: string; model?: unknown; usage?: Counts } }

const count = (value: unknown) => (typeof value === "number" && Number.isFinite(value) && value > 0 ? Math.round(value) : 0)

// add sums one usage record into `tokens`, answering the record's own total.
function add(tokens: Tokens, usage: Counts | undefined): number {
  if (!usage || typeof usage !== "object") return 0
  const input = count(usage.input)
  const output = count(usage.output)
  const cacheRead = count(usage.cacheRead)
  const cacheWrite = count(usage.cacheWrite)
  tokens.input += input
  tokens.output += output
  tokens.cacheRead += cacheRead
  tokens.cacheWrite += cacheWrite
  return input + output + cacheRead + cacheWrite
}

/**
 * summarizeUsage sums a worker session's recorded usage: every assistant message,
 * the usage a tool result carries, cache-warm `usage` entries, and compaction and
 * branch-summary calls. The model is the one that produced the most tokens, so a
 * run that fell back to another model for one turn is attributed to the model
 * that did the work, deterministically (first seen wins a tie).
 */
export function summarizeUsage(entries: readonly Entry[]): UsageSummary {
  const tokens: Tokens = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }
  const byModel = new Map<string, number>()
  let model = ""
  for (const entry of entries ?? []) {
    if (!entry || typeof entry !== "object") continue
    if (entry.type === "usage") add(tokens, entry.usage)
    else if ((entry.type === "branch_summary" || entry.type === "compaction") && entry.usage) add(tokens, entry.usage)
    if (entry.type !== "message" || !entry.message) continue
    const message = entry.message
    if (message.role !== "assistant" && message.role !== "toolResult") continue
    const produced = add(tokens, message.usage)
    const name = typeof message.model === "string" ? message.model.trim() : ""
    if (message.role === "assistant" && name) {
      const seen = (byModel.get(name) ?? 0) + (produced || 0)
      byModel.set(name, seen)
      if (model === "" || seen > (byModel.get(model) ?? 0)) model = name
    }
  }
  return { tokens, total: tokens.input + tokens.output + tokens.cacheRead + tokens.cacheWrite, model }
}

/**
 * reportable is whether there is a count worth posting: the Office refuses a
 * report that carries neither tokens nor an amount, and a model is what it prices
 * the tokens against, so a run with no recorded usage - or none Pi named a model
 * for - reports nothing rather than a zero.
 */
export function reportable(summary: UsageSummary): boolean {
  return summary.total > 0 && summary.model.trim() !== ""
}

/** usageReport is the request body for one finished run. */
export function usageReport(options: { taskID: string; run: string; day: string; summary: UsageSummary }) {
  return {
    source: "harness:pi",
    day: options.day,
    model: options.summary.model.trim(),
    tokens: {
      input: options.summary.tokens.input,
      output: options.summary.tokens.output,
      cache_read: options.summary.tokens.cacheRead,
      cache_write: options.summary.tokens.cacheWrite,
    },
    task_id: options.taskID,
    idempotency_key: usageKey(options.taskID, options.run),
  }
}

/** usageKey is the per-run idempotency key: the task and the run, nothing else. */
export function usageKey(taskID: string, run: string): string {
  return `pi:${taskID.trim()}:${run.trim()}`
}

/** dayOf is the UTC day a run started, the day Office bills its spend to. */
export function dayOf(timestamp: string | undefined, now: Date = new Date()): string {
  const start = typeof timestamp === "string" ? new Date(timestamp) : new Date(NaN)
  const at = Number.isNaN(start.getTime()) ? now : start
  return at.toISOString().slice(0, 10)
}

const RAN_TIERS = { model: ["fast", "strong", "strongest"], thinking: ["low", "medium", "high"] } as const

/**
 * presetRan reads the tiers a dispatcher handed the worker it started, and
 * answers them only when both are ones Office accepts. Anything else is no
 * report: `update_task`'s `preset_ran` needs both tiers, and a tier this side
 * could not name is not what ran.
 */
export function presetRan(raw: string | undefined): { model: string; thinking: string } | undefined {
  if (!raw?.trim()) return undefined
  let parsed: any
  try {
    parsed = JSON.parse(raw)
  } catch {
    return undefined
  }
  const model = typeof parsed?.model === "string" ? parsed.model.trim() : ""
  const thinking = typeof parsed?.thinking === "string" ? parsed.thinking.trim() : ""
  if (!(RAN_TIERS.model as readonly string[]).includes(model)) return undefined
  if (!(RAN_TIERS.thinking as readonly string[]).includes(thinking)) return undefined
  return { model, thinking }
}
