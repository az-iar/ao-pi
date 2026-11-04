// agent-office#560: session approvals live in one host-wide directory with the
// runtime recorded, so two agents' sessions on one machine resolve their own
// identity. HOME is its own: the directory is host-wide, and a shared HOME would
// let these read what another test wrote.
import assert from "node:assert/strict"
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, statSync, unlinkSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"
import { test } from "node:test"

process.env.HOME = mkdtempSync(join(tmpdir(), "ao-pi-sessions-"))
const { legacySessionPointerPath, saveSessionAgent, sessionAgentOf, sessionPointerPath } = await import("../lib/office.ts")

// A root process ignores the permission bits, so the read-only test would pass
// without exercising EACCES; skip loudly instead.
const asRoot = typeof process.getuid === "function" && process.getuid() === 0

test("the record is host-wide and names the runtime", () => {
  const path = sessionPointerPath("session-ada")
  assert.ok(path.includes(join("agent-office", "sessions")), path)
  assert.ok(!path.includes(join("agent-office", "pi", "sessions")), `not under the runtime: ${path}`)
  saveSessionAgent("session-ada", "pi")
  const record = JSON.parse(readFileSync(path, "utf8"))
  assert.deepEqual({ agent: record.agent, runtime: record.runtime }, { agent: "pi", runtime: "pi" })
})

test("a legacy record is read once, copied forward and retired", () => {
  const id = "session-legacy"
  mkdirSync(dirname(legacySessionPointerPath(id)), { recursive: true })
  writeFileSync(legacySessionPointerPath(id), JSON.stringify({ agent: "pi" }), { mode: 0o600 })
  assert.equal(sessionAgentOf(id), "pi")
  assert.equal(existsSync(legacySessionPointerPath(id)), false, "the legacy file is retired once copied")
  unlinkSync(sessionPointerPath(id))
  assert.equal(sessionAgentOf(id), undefined, "revoking the record stays revoked")
})

test("a read-only directory still returns the agent and writes nothing",
  { skip: asRoot ? "running as root: permission bits are not enforced" : false }, () => {
  const id = "session-readonly"
  mkdirSync(dirname(legacySessionPointerPath(id)), { recursive: true })
  writeFileSync(legacySessionPointerPath(id), JSON.stringify({ agent: "pi" }), { mode: 0o600 })
  const host = dirname(sessionPointerPath(id))
  mkdirSync(host, { recursive: true })
  chmodSync(host, 0o500)
  try {
    assert.equal(sessionAgentOf(id), "pi", "the identity survives a read-only directory")
    assert.equal(existsSync(sessionPointerPath(id)), false, "nothing was copied")
    assert.equal(existsSync(legacySessionPointerPath(id)), true, "the legacy record is left in place, not lost")
  } finally {
    chmodSync(host, 0o700)
  }
})

test("two sessions on one host read their own agent", () => {
  saveSessionAgent("session-one", "pi")
  saveSessionAgent("session-two", "other-agent")
  assert.equal(sessionAgentOf("session-one"), "pi")
  assert.equal(sessionAgentOf("session-two"), "other-agent")
  assert.equal(sessionAgentOf("session-three"), undefined)
})

// mkdirSync's mode applies only to a directory it creates, so an existing
// sessions/ with looser bits stayed loose and the records were readable by anyone
// who could read the parent (bruno on ao-pi#39, ao-pi#40).
test("an existing sessions directory is tightened to 0700", () => {
  const id = "session-loose-dir"
  const dir = dirname(sessionPointerPath(id))
  mkdirSync(dir, { recursive: true })
  chmodSync(dir, 0o755)
  assert.equal(statSync(dir).mode & 0o777, 0o755, "fixture: the directory starts loose")
  saveSessionAgent(id, "pi")
  assert.equal(statSync(dir).mode & 0o777, 0o700, "the record's directory is private after a save")
  assert.equal(statSync(sessionPointerPath(id)).mode & 0o777 & 0o077, 0, "and the record itself is not group- or world-readable")
})

// `runtime` records the writer, not the approver: a re-approval by another plugin
// overwrites it, and the field says so rather than implying it names the approver.
test("runtime names the last writer, so a re-approval overwrites it", () => {
  const id = "session-rewritten"
  saveSessionAgent(id, "pi")
  writeFileSync(sessionPointerPath(id), JSON.stringify({ agent: "pi", runtime: "claude", at: Date.now() }), { mode: 0o600 })
  assert.equal(sessionAgentOf(id), "pi", "the identity is unaffected by who wrote the record")
  assert.equal(JSON.parse(readFileSync(sessionPointerPath(id), "utf8")).runtime, "claude")
})
