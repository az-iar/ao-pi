// `/office session <name>` resumes a previous conversation by its /name. The
// resolver is the whole decision, so it is tested on its own: exact beats
// prefix beats substring, a shared exact name resumes the newest, and anything
// ambiguous is reported instead of guessed.
import assert from "node:assert/strict"
import test from "node:test"
import { message, pickSession } from "../lib/sessions.ts"

const at = (stamp: string) => new Date(`${stamp}T00:00:00Z`)
const info = (name: string | undefined, stamp: string, extra: Partial<Parameters<typeof pickSession>[0][number]> = {}) => ({
  path: `/sessions/${stamp}-${name ?? "unnamed"}.jsonl`,
  id: stamp,
  cwd: "/repo",
  name,
  created: at(stamp),
  modified: at(stamp),
  messageCount: 3,
  firstMessage: `first ${name ?? "unnamed"}`,
  allMessagesText: "",
  ...extra,
})

test("an exact name resumes that session", () => {
  const choice = pickSession([info("bruno", "2026-11-01"), info("bruno-review", "2026-11-02")], "bruno")
  assert.equal(choice.kind, "resume")
  assert.equal(choice.kind === "resume" && choice.session.name, "bruno")
  assert.equal(choice.kind === "resume" && choice.via, "exact")
})

test("a prefix resumes only when one session has it", () => {
  const one = pickSession([info("bruno-work", "2026-11-01"), info("anna", "2026-11-02")], "bruno")
  assert.equal(one.kind === "resume" && one.via, "prefix")

  const many = pickSession([info("bruno-a", "2026-11-01"), info("bruno-b", "2026-11-02")], "bruno")
  assert.equal(many.kind, "ambiguous")
  assert.equal(many.kind === "ambiguous" && many.matches.length, 2)
})

test("a substring works when it is unique, and names are matched case-insensitively", () => {
  const choice = pickSession([info("Office Bruno Rebase", "2026-11-01"), info("anna", "2026-11-02")], "bruno")
  assert.equal(choice.kind === "resume" && choice.via, "substring")
  assert.equal(pickSession([info("BRUNO", "2026-11-01")], "bruno").kind, "resume")
})

test("a shared exact name resumes the most recent one, which is the order pig -r shows", () => {
  const choice = pickSession([info("bruno", "2026-11-01"), info("bruno", "2026-11-03"), info("bruno", "2026-11-02")], "bruno")
  assert.equal(choice.kind, "resume")
  assert.equal(choice.kind === "resume" && choice.session.modified.toISOString().slice(0, 10), "2026-11-03")
})

test("nothing matching reports the folder's named sessions instead of guessing", () => {
  const choice = pickSession([info("anna", "2026-11-01"), info(undefined, "2026-11-02"), info("carl", "2026-11-03")], "bruno")
  assert.equal(choice.kind, "none")
  assert.deepEqual(choice.kind === "none" && choice.named.map((s) => s.name), ["carl", "anna"])
  assert.match(message(choice), /No session is named "bruno" here/)
  assert.match(message(choice), /carl/)
})

test("an empty query lists what can be resumed, and a folder with none says how to name one", () => {
  const listed = pickSession([info("anna", "2026-11-01")], "")
  assert.equal(listed.kind, "none")
  assert.match(message(listed), /Named sessions in this folder/)
  const empty = pickSession([info(undefined, "2026-11-01")], "")
  assert.match(message(empty), /No named sessions in this folder yet/)
})

test("the ambiguous count only gains a + when the list was cut at five", () => {
  const three = pickSession([info("bruno-a", "2026-11-01"), info("bruno-b", "2026-11-02"), info("bruno-c", "2026-11-03")], "bruno")
  assert.equal(three.kind, "ambiguous")
  assert.equal(three.kind === "ambiguous" && three.total, 3)
  assert.match(message(three), /names 3 sessions/)
  assert.doesNotMatch(message(three), /3\+ sessions/)

  const many = Array.from({ length: 7 }, (_, i) => info(`bruno-${i}`, `2026-11-0${i + 1}`))
  const cut = pickSession(many, "bruno")
  assert.equal(cut.kind, "ambiguous")
  assert.equal(cut.kind === "ambiguous" && cut.matches.length, 5)
  assert.equal(cut.kind === "ambiguous" && cut.total, 7)
  assert.match(message(cut), /names 7\+ sessions/)
})

test("resuming the session already open says so rather than switching", () => {
  const only = info("bruno", "2026-11-01")
  const choice = pickSession([only], "bruno")
  assert.equal(choice.kind, "resume")
  assert.match(message(choice, only.path), /Already in the session named bruno/)
  assert.match(message(choice), /Continuing "bruno"/)
})
