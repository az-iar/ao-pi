// agent-office#560: session approvals live in one host-wide directory with the
// runtime recorded, so two agents' sessions on one machine resolve their own
// identity. HOME is its own: the directory is host-wide, and a shared HOME would
// let these read what another test wrote.
import assert from "node:assert/strict"
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, unlinkSync, writeFileSync } from "node:fs"
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
