// agent-office#560: session approvals live in one host-wide directory with the
// runtime recorded, so two agents' sessions on one machine resolve their own
// identity. HOME is its own: the directory is host-wide, and a shared HOME would
// let these read what another test wrote.
import assert from "node:assert/strict"
import { chmodSync, copyFileSync, existsSync, mkdirSync, readFileSync, statSync, unlinkSync, writeFileSync } from "node:fs"
import { dirname, join } from "node:path"
import { test } from "node:test"
import { tempDir } from "./tmp.mts"

process.env.HOME = tempDir("pi-sessions")
const { legacySessionPointerPath, refusedSessionPairCount, saveSessionAgent, sessionAgentOf, sessionPointerPath } = await import("../lib/office.ts")

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

// ao-opencode#40, ported: the record carries the id it approves, and the reader
// compares it. Without the comparison a record copied to another session's
// filename is indistinguishable from a real approval - the file is sha256(id)[:16],
// so the name cannot be read back - and the reader learns the agent of a session
// that never approved anything.
const refused = (run: () => unknown) => {
  const original = console.warn
  const seen: string[] = []
  console.warn = (...args: unknown[]) => void seen.push(args.map(String).join(" "))
  try {
    return { result: run(), warnings: seen }
  } finally {
    console.warn = original
  }
}
const mismatches = (seen: string[]) => seen.filter((warning) => warning.includes("ignoring it"))
// A record on disk that names another session, written the way the reader finds
// it: a JSON record, through the path sessionAgentOf reads.
const recordNaming = (id: string, session: unknown, agent = "pi") => {
  mkdirSync(dirname(sessionPointerPath(id)), { recursive: true })
  writeFileSync(sessionPointerPath(id), JSON.stringify({ agent, session }, null, 2) + "\n", { mode: 0o600 })
}

test("a record naming another session is refused, not honoured", () => {
  const owner = "session-owner-copy"
  const copier = "session-copier-copy"
  saveSessionAgent(owner, "pi")
  copyFileSync(sessionPointerPath(owner), sessionPointerPath(copier))
  const { result, warnings: seen } = refused(() => sessionAgentOf(copier))
  assert.equal(result, undefined, "a record naming another session handed that session its agent")
  const mismatch = mismatches(seen)
  assert.equal(mismatch.length, 1, `the refusal was not reported once: ${seen.join(" | ")}`)
  assert.ok(mismatch[0].includes(copier), `the warning does not name the session it was fetched by: ${mismatch[0]}`)
  assert.ok(mismatch[0].includes(owner), `the warning does not name the session the record belongs to: ${mismatch[0]}`)
  // The poll repeats this read on a timer, so the pair is reported once and the
  // second read is silent but still refused.
  const again = refused(() => [sessionAgentOf(copier), sessionAgentOf(copier)])
  assert.deepEqual(again.result, [undefined, undefined], "a repeat read honoured the mismatched record")
  assert.deepEqual(again.warnings, [], `the same pair warned again on a repeat read: ${again.warnings.join(" | ")}`)
})

test("a record with no session field is still honoured, and no mismatch is warned about", () => {
  const id = "session-no-field"
  // The shape of every record written before the field was read: refusing it
  // would revoke every session on upgrade.
  mkdirSync(dirname(sessionPointerPath(id)), { recursive: true })
  writeFileSync(sessionPointerPath(id), JSON.stringify({ agent: "pi" }, null, 2) + "\n", { mode: 0o600 })
  const { result, warnings: seen } = refused(() => sessionAgentOf(id))
  assert.equal(result, "pi", "a record written before this change was refused on upgrade")
  assert.deepEqual(seen, [], `a record with no session field was warned about: ${seen.join(" | ")}`)
})

test("a mismatched legacy record is not copied forward", () => {
  const owner = "session-owner-legacy"
  const copier = "session-copier-legacy"
  saveSessionAgent(owner, "pi")
  const legacy = legacySessionPointerPath(copier)
  mkdirSync(dirname(legacy), { recursive: true })
  copyFileSync(sessionPointerPath(owner), legacy)
  const before = readFileSync(legacy, "utf8")
  const { result, warnings: seen } = refused(() => sessionAgentOf(copier))
  assert.equal(result, undefined, "a legacy record naming another session granted it that agent")
  // Promoting it would make the stale file the record every later read sees.
  assert.equal(existsSync(sessionPointerPath(copier)), false, "a mismatched legacy record was copied forward into the shared directory")
  assert.equal(readFileSync(legacy, "utf8"), before, "the refused legacy record was rewritten")
  assert.ok(
    seen.some((warning) => warning.includes("ignoring it") && warning.includes(copier) && warning.includes(owner)),
    `the refusal was not reported: ${seen.join(" | ")}`,
  )
})

// ao-pi#48 (bruno): the refusal above returned undefined while the shared record
// was all there was, so a refusal that fell through to the legacy read stayed
// invisible - and falling through is not a smaller refusal, it is the identity:
// the legacy record is copied forward, retires itself and becomes the record every
// later read sees. The legacy record is written for every case here and names this
// session, so a fall through gated on the legacy record naming the requested
// session has something to hand back; a fixture with no session field would keep
// such a fall through green. Its effects are what is asserted - the return value,
// and both files byte for byte - because whether the legacy file was read is not
// observable, and a refusal that hands the legacy record back is a preference, not
// a rule. Asserted on the files as well as the return value, because a promotion
// can hand back the same agent a refusal would have withheld.
const noFallThrough = (id: string, session: unknown) => {
  // Distinct agents: a fall through returns the legacy agent, so a refusal and a
  // promotion cannot be confused in the message.
  recordNaming(id, session, "other-session-agent")
  const sharedPath = sessionPointerPath(id)
  const sharedBefore = readFileSync(sharedPath, "utf8")
  const legacyPath = legacySessionPointerPath(id)
  mkdirSync(dirname(legacyPath), { recursive: true })
  writeFileSync(legacyPath, JSON.stringify({ agent: "pi", session: id }, null, 2) + "\n", { mode: 0o600 })
  const legacyBefore = readFileSync(legacyPath, "utf8")

  const { result, warnings: seen } = refused(() => sessionAgentOf(id))
  assert.equal(result, undefined, `a refused shared record fell through to the legacy record and returned ${String(result)}`)
  const mismatch = mismatches(seen)
  assert.equal(mismatch.length, 1, `the refusal was not reported once: ${seen.join(" | ")}`)
  assert.ok(mismatch[0].includes(id), `the warning does not name the session it was fetched by: ${mismatch[0]}`)
  // The legacy file is neither consumed (unlinked after a successful copy) nor
  // rewritten: a promotion both overwrites the shared record and retires it.
  assert.equal(existsSync(legacyPath), true, "the refused read consumed the legacy record")
  assert.equal(readFileSync(legacyPath, "utf8"), legacyBefore, "the refused read rewrote the legacy record")
  // No new shared record: the refused one is still there, byte for byte, and still
  // names the session it was written for rather than this one.
  assert.equal(readFileSync(sharedPath, "utf8"), sharedBefore, "the refused read wrote the legacy agent into the shared record")
  assert.equal(JSON.parse(readFileSync(sharedPath, "utf8")).session, session, "the shared record was re-written for this session")
}

test("a shared record refused for a mismatched session does not fall through to the legacy record", () => {
  noFallThrough("session-no-fallthrough-mismatch", "session-someone-else")
})

test("a shared record refused for a wrong-typed session does not fall through to the legacy record", () => {
  noFallThrough("session-no-fallthrough-wrong-type", 123)
})

// ao-opencode#41 (bruno): a `session` field that is present but not a string was
// read as "no field" and honoured for every session, so "fail closed" was false.
// The refusal is a mismatch like any other, and the value it logs is quoted: a
// planted string with a newline in it must not be able to forge a log line.
test("a session field of the wrong type is refused, and the warning is JSON-quoted", () => {
  const values: unknown[] = [null, 123, [], {}, true, "other\nforged line"]
  values.forEach((value, index) => {
    const id = `session-wrong-type-${index}`
    recordNaming(id, value)
    const { result, warnings: seen } = refused(() => sessionAgentOf(id))
    assert.equal(result, undefined, `a session field of ${JSON.stringify(value)} honoured the record`)
    const mismatch = mismatches(seen)
    assert.equal(mismatch.length, 1, `the refusal of ${JSON.stringify(value)} was not reported once: ${seen.join(" | ")}`)
    assert.ok(mismatch[0].includes(JSON.stringify(value)), `the warning does not quote the value it refused: ${mismatch[0]}`)
    assert.ok(!/[\r\n]/.test(mismatch[0]), `a planted value forged a line in the warning: ${JSON.stringify(mismatch[0])}`)
    // Once per pair, and the repeat read is still refused.
    const again = refused(() => sessionAgentOf(id))
    assert.equal(again.result, undefined, `a repeat read honoured ${JSON.stringify(value)}`)
    assert.deepEqual(again.warnings, [], `the same pair warned again: ${again.warnings.join(" | ")}`)
  })
})

// ao-opencode#41: the set of pairs already warned about was unbounded, and a
// record whose `session` keeps changing is a new pair on every read, so a
// long-lived poller held one string per pair it ever saw. 150 distinct pairs is
// more than the cap, so a set that grew to 150 fails here; the count is the whole
// of the surface the bound needs, because nothing else about the cap is visible
// from outside. A test that repeated one pair would pass with no cap at all.
test("the warned pairs are bounded", () => {
  const cap = 100
  const sizes: number[] = []
  let warned = 0
  for (let i = 0; i < 150; i++) {
    // A distinct (requested, stored) pair per read, through the path the poll
    // takes: a record on disk that names another session.
    const id = `bounded-session-${i}`
    recordNaming(id, `bounded-owner-${i}`)
    const { result, warnings: seen } = refused(() => sessionAgentOf(id))
    assert.equal(result, undefined, `pair ${i} honoured a record naming another session`)
    warned += mismatches(seen).length
    sizes.push(refusedSessionPairCount())
  }
  // Every distinct pair is reported when it is first read, capped or not.
  assert.equal(warned, 150, `the 150 distinct pairs did not each warn once: ${warned}`)
  const largest = Math.max(...sizes)
  assert.ok(largest <= cap, `the warned pairs set held ${largest} pairs, past the cap of ${cap}`)
})

// The consequence of clearing rather than evicting: a pair lost to the clear is
// refused and warned about again, once, and then stays quiet - a fresh process
// reading the same bad record. 101 distinct pairs overflow any starting state,
// because a clear is what an add at the cap does.
test("a pair that was cleared warns again once", () => {
  const id = "cleared-session"
  recordNaming(id, "cleared-owner")
  const first = refused(() => sessionAgentOf(id))
  assert.equal(first.result, undefined, "fixture: a record naming another session was honoured")
  assert.equal(mismatches(first.warnings).length, 1, `fixture: the pair was not warned about: ${first.warnings.join(" | ")}`)
  const quiet = refused(() => sessionAgentOf(id))
  assert.deepEqual(quiet.warnings, [], `fixture: the same pair warned twice with no clear between: ${quiet.warnings.join(" | ")}`)

  for (let i = 0; i < 101; i++) {
    const other = `cleared-overflow-${i}`
    recordNaming(other, `cleared-overflow-owner-${i}`)
    refused(() => sessionAgentOf(other))
  }

  // The cleared pair is still refused, so the clear costs a repeat of the warning
  // and nothing else.
  const again = refused(() => [sessionAgentOf(id), sessionAgentOf(id)])
  assert.deepEqual(again.result, [undefined, undefined], "a cleared pair's record was honoured after the clear")
  assert.equal(mismatches(again.warnings).length, 1, `a cleared pair did not warn exactly once after the clear: ${again.warnings.join(" | ")}`)
  const silent = refused(() => sessionAgentOf(id))
  assert.deepEqual(silent.warnings, [], `the re-warned pair warned again on a repeat read: ${silent.warnings.join(" | ")}`)
})
