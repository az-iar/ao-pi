// The state directory must be resolved when it is used, not when the module loads:
// a harness sets $HOME before importing, so anything that stands up more than one
// harness in a process would otherwise have them all read the first one's records.
// The current suites run one harness per file, so this is preventive - the test
// pins the property rather than a failure anyone has seen.
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
