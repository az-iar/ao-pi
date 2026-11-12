// A task's model and thinking preset, turned into a worker's argv.
//
// Office resolves a task to tiers, not to models: a single-task read carries
// `preset: {model, thinking}` where model is fast, strong or strongest and
// thinking is low, medium or high (agent-office#838, #841). Which concrete model
// and thinking level a tier means is this dispatcher's decision, so the mapping
// is configuration -- the file below, or AGENT_OFFICE_PRESETS naming another one
// -- and never a value written into the dispatch code.
//
// A task with no preset (an Office older than #841, or a clamped read that failed)
// gets no flags at all, so its worker keeps exactly the args it had before.

import { existsSync, readFileSync } from "node:fs"
import { homedir } from "node:os"
import { join } from "node:path"

export type PresetTiers = { model: Record<string, string>; thinking: Record<string, string> }

// The shipped mapping: the cheap model for grunt work, the mid one for an
// ordinary fix or feature, the top one for research and design. Documented in
// README.md ("Model and thinking presets"), and overridable per host: these are
// the ids available with Pi's Claude subscription provider on the reference
// machine, and a host whose credentials name another provider maps its own.
export const DEFAULT_PRESET_TIERS: PresetTiers = {
  model: { fast: "claude-code/claude-haiku-5-5", strong: "claude-code/claude-sonnet-5-5", strongest: "claude-code/claude-opus-5-5" },
  thinking: { low: "low", medium: "medium", high: "high" },
}

// Resolved at call time, never cached at import: a harness sets $HOME before
// importing, and an edit made while the session runs has to reach the next
// dispatch rather than the next restart.
export function presetsConfigPath(): string {
  return process.env.AGENT_OFFICE_PRESETS?.trim() || join(homedir(), ".config", "agent-office", "presets.json")
}

const copy = (tiers: PresetTiers): PresetTiers => ({ model: { ...tiers.model }, thinking: { ...tiers.thinking } })

// tiers overlays one axis of the file onto the shipped defaults, key by key, so a
// file that names only `strongest` changes only that.
function tiers(value: unknown, fallback: Record<string, string>): Record<string, string> {
  const out = { ...fallback }
  if (!value || typeof value !== "object" || Array.isArray(value)) return out
  for (const [tier, concrete] of Object.entries(value as Record<string, unknown>)) {
    if (typeof concrete === "string" && concrete.trim()) out[tier] = concrete.trim()
  }
  return out
}

// loadPresetTiers reads the mapping. A file that is not there is the defaults; one
// that cannot be read or parsed is the defaults too, with a warning, because a
// broken config must not stop a dispatch.
export function loadPresetTiers(path = presetsConfigPath()): PresetTiers {
  let parsed: any
  try {
    if (!existsSync(path)) return copy(DEFAULT_PRESET_TIERS)
    parsed = JSON.parse(readFileSync(path, "utf8"))
  } catch (err) {
    console.warn(`[agent-office] could not read the preset tiers at ${path}, using the defaults: ${err instanceof Error ? err.message : String(err)}`)
    return copy(DEFAULT_PRESET_TIERS)
  }
  return { model: tiers(parsed?.model, DEFAULT_PRESET_TIERS.model), thinking: tiers(parsed?.thinking, DEFAULT_PRESET_TIERS.thinking) }
}

// presetFlags answers a worker's extra argv for the `preset` Office sent. Absent,
// null or not an object is no flags at all -- today's args, exactly. A tier the
// config does not map is left out rather than guessed: the worker then runs with
// Pi's own default for that axis, never with a model the task did not ask for.
export function presetFlags(preset: unknown, tiers: PresetTiers = loadPresetTiers()): string[] {
  const flags: string[] = []
  if (!preset || typeof preset !== "object" || Array.isArray(preset)) return flags
  const { model, thinking } = preset as { model?: unknown; thinking?: unknown }
  const one = (tier: unknown, map: Record<string, string>, flag: string) => {
    if (typeof tier !== "string") return
    const name = tier.trim()
    // Own keys only: `constructor` or `__proto__` in Office's text must not resolve
    // to something off Object.prototype and become a non-string argv entry.
    if (Object.hasOwn(map, name) && map[name]) flags.push(flag, map[name])
  }
  one(model, tiers.model, "--model")
  one(thinking, tiers.thinking, "--thinking")
  return flags
}

// workerPresetFlags is presetFlags for the task read a dispatch holds.
export function workerPresetFlags(task: any, tiers: PresetTiers = loadPresetTiers()): string[] {
  return presetFlags(task?.preset, tiers)
}

export type PresetArgs = { flags: string[]; ran?: { model: string; thinking: string } }

// ranTiers is which tiers of the task's preset actually reached the worker: a
// tier the config does not map is left off the command line, so the worker ran
// Pi's own default for that axis and reporting it as the task's preset would be
// a claim about what ran that is not true. Both tiers or none (agent-office#841).
function ranTiers(preset: unknown, tiers: PresetTiers): { model: string; thinking: string } | undefined {
  if (!preset || typeof preset !== "object" || Array.isArray(preset)) return undefined
  const { model, thinking } = preset as { model?: unknown; thinking?: unknown }
  const applied = (tier: unknown, map: Record<string, string>) => {
    const name = typeof tier === "string" ? tier.trim() : ""
    return name && Object.hasOwn(map, name) && map[name] ? name : ""
  }
  const ranModel = applied(model, tiers.model)
  const ranThinking = applied(thinking, tiers.thinking)
  return ranModel && ranThinking ? { model: ranModel, thinking: ranThinking } : undefined
}

// presetArgs is one dispatch's flags plus the tiers they came from, from one read
// of the mapping: the worker's argv and the `preset_ran` it reports later (Office
// 1b8f755a) must never disagree about which preset ran.
export function presetArgs(preset: unknown, tiers: PresetTiers = loadPresetTiers()): PresetArgs {
  return { flags: presetFlags(preset, tiers), ran: ranTiers(preset, tiers) }
}

// workerPresetArgs is presetArgs for the task read a dispatch holds.
export function workerPresetArgs(task: any, tiers: PresetTiers = loadPresetTiers()): PresetArgs {
  return presetArgs(task?.preset, tiers)
}
