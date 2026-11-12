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
import { DEFAULT_PRESET_TIERS, loadPresetTiers, presetFlags, workerPresetArgs, workerPresetFlags } from "../lib/presets.ts"
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
  assert.deepEqual(presetFlags({ model: "fast", thinking: "low" }, tiers), ["--model", "claude-code/claude-haiku-5-5", "--thinking", "low"])
  assert.deepEqual(presetFlags({ model: "strongest", thinking: "high" }, tiers), ["--model", "claude-code/claude-opus-5-5", "--thinking", "high"])
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
  // Names that resolve on Object.prototype are not tiers: no flag, no non-string argv.
  for (const name of ["constructor", "__proto__", "toString", "hasOwnProperty"]) {
    assert.deepEqual(presetFlags({ model: name, thinking: name }, tiers), [], `${name} must not resolve`)
  }
})

test("presetArgs reports only the tiers that reached the command line", (t) => {
  // Office 1b8f755a: the worker reports `preset_ran`, but only a tier the mapping
  // named actually ran - the other axis got Pi's own default, so claiming the
  // task's preset for it would be false.
  const tiers = loadPresetTiers(config(t, { model: { strongest: "some/other-model" } }))
  assert.deepEqual(workerPresetArgs({ preset: { model: "strong", thinking: "high" } }, tiers), {
    flags: ["--model", DEFAULT_PRESET_TIERS.model.strong, "--thinking", "high"],
    ran: { model: "strong", thinking: "high" },
  })
  // The override reaches the flags as well as the report, from the same read.
  assert.deepEqual(workerPresetArgs({ preset: { model: "strongest", thinking: "low" } }, tiers), {
    flags: ["--model", "some/other-model", "--thinking", "low"],
    ran: { model: "strongest", thinking: "low" },
  })
  // An unmapped tier on either axis is no report at all: `preset_ran` needs both.
  assert.equal(workerPresetArgs({ preset: { model: "turbo", thinking: "high" } }, tiers).ran, undefined)
  assert.equal(workerPresetArgs({ preset: { model: "strong", thinking: "lo" } }, tiers).ran, undefined)
  // A tier that resolves on Object.prototype is not a tier: nothing ran to report.
  assert.equal(workerPresetArgs({ preset: { model: "constructor", thinking: "high" } }, tiers).ran, undefined)
  // No preset, or a shape an older Office sends: no flags at all, and nothing to
  // report. One axis alone still travels, but is not a preset that ran.
  for (const task of [{}, { preset: null }, { preset: "strong" }, { preset: [] }]) {
    assert.deepEqual(workerPresetArgs(task, tiers), { flags: [], ran: undefined }, JSON.stringify(task))
  }
  assert.deepEqual(workerPresetArgs({ preset: { model: "strong" } }, tiers), {
    flags: ["--model", DEFAULT_PRESET_TIERS.model.strong],
    ran: undefined,
  })
})
