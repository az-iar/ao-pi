// A task's model and thinking preset, and the config that turns its tiers into a
// concrete `--model` and `--thinking` (agent-office#838, #841).
//
// Office sends tiers; which model a tier means is the dispatcher's decision, so it
// lives in a config file. These cases pin the shipped mapping, a file that
// overrides part of it, and the fallback that matters most: a task with no preset
// must get no flags at all, so a worker's args stay exactly what they were.
import assert from "node:assert/strict"
import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { type TestContext, test } from "node:test"
import { DEFAULT_PRESET_TIERS, loadPresetTiers, presetFlags, workerPresetFlags } from "../lib/presets.ts"
import { ownsTempDir } from "./tmp.mts"

// A config file of this test's own, so the host's real presets file cannot decide
// what the tests see.
function config(t: TestContext, body?: unknown): string {
  const dir = ownsTempDir(t, "presets")
  const path = join(dir, "presets.json")
  if (body !== undefined) writeFileSync(path, typeof body === "string" ? body : JSON.stringify(body))
  return path
}

test("with no config file the shipped tiers are the mapping", (t) => {
  const tiers = loadPresetTiers(config(t))
  assert.deepEqual(tiers, DEFAULT_PRESET_TIERS)
  // The shipped values themselves: the cheap model is fast, the top one strongest.
  assert.deepEqual(presetFlags({ model: "fast", thinking: "low" }, tiers), ["--model", "anthropic/claude-haiku-4-5", "--thinking", "low"])
  assert.deepEqual(presetFlags({ model: "strongest", thinking: "high" }, tiers), ["--model", "anthropic/claude-opus-4-8", "--thinking", "high"])
})

test("a config file overrides one tier and leaves the rest", (t) => {
  const tiers = loadPresetTiers(config(t, { model: { strong: "some/other-model" } }))
  assert.equal(tiers.model.strong, "some/other-model")
  assert.equal(tiers.model.fast, DEFAULT_PRESET_TIERS.model.fast, "an unnamed tier keeps the shipped value")
  assert.deepEqual(tiers.thinking, DEFAULT_PRESET_TIERS.thinking, "an axis the file does not mention is untouched")
  // A value that is not a non-empty string is ignored rather than turned into a flag.
  const loose = loadPresetTiers(config(t, { model: { fast: "", strong: 7, strongest: null } }))
  assert.deepEqual(loose.model, DEFAULT_PRESET_TIERS.model)
})

test("a broken config file falls back to the shipped tiers instead of stopping dispatch", (t) => {
  const previous = console.warn
  const warnings: string[] = []
  console.warn = (line: string) => warnings.push(String(line))
  try {
    const tiers = loadPresetTiers(config(t, "{ not json"))
    assert.deepEqual(tiers, DEFAULT_PRESET_TIERS)
  } finally {
    console.warn = previous
  }
  assert.equal(warnings.length, 1, "the broken file is reported once")
})

test("a task with no preset gets no flags, and neither does an empty one", (t) => {
  const tiers = loadPresetTiers(config(t, {}))
  // The shapes an older Office (no field) and #841's null/clamped-failure read send.
  for (const task of [{}, { preset: null }, { preset: undefined }, { preset: "strong" }, { preset: [] }, { id: "t", status: "todo" }]) {
    assert.deepEqual(workerPresetFlags(task, tiers), [], `${JSON.stringify(task)} must start with today's args`)
  }
})

test("a tier the mapping does not know is left out, not guessed", (t) => {
  const tiers = loadPresetTiers(config(t, {}))
  // A tier name this version of Office never sends: no flag for that axis, rather
  // than a model picked by guesswork. The axis that is known still travels.
  assert.deepEqual(presetFlags({ model: "turbo", thinking: "high" }, tiers), ["--thinking", "high"])
  assert.deepEqual(presetFlags({ model: "fast", thinking: "lo" }, tiers), ["--model", DEFAULT_PRESET_TIERS.model.fast])
})
