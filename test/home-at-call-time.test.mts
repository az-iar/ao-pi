// The state directory must be resolved when it is used, not when the module loads.
// Several harnesses run in one process and each sets $HOME before importing, so a
// cached directory makes every one of them read the first harness's records - the
// HOME trap that made ao-opencode#35 red only in the full suite (agent-office#560).
import assert from "node:assert/strict"
import { mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"

const first = mkdtempSync(join(tmpdir(), "ao-home-a-"))
process.env.HOME = first
const { agentCredentialsPath, sessionPointerPath } = await import("../lib/office.ts")

test("a second HOME in the same process points at its own records", () => {
  assert.ok(sessionPointerPath("s1").startsWith(join(first, ".config")), sessionPointerPath("s1"))
  const second = mkdtempSync(join(tmpdir(), "ao-home-b-"))
  process.env.HOME = second
  try {
    assert.ok(sessionPointerPath("s1").startsWith(join(second, ".config")), `still the old HOME: ${sessionPointerPath("s1")}`)
    assert.ok(agentCredentialsPath("pi").startsWith(join(second, ".config")), agentCredentialsPath("pi"))
  } finally {
    process.env.HOME = first
  }
})
