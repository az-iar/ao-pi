// ao-opencode#41 (bruno), ported: adoptAgent skipped the write when the stored
// *agent* already matched, never looking at the session. A record refused for this
// session therefore survived a re-approval as the same agent: /office worked in
// the process, the next start refused the session again, and only deleting the
// file ended the loop. Both directions are covered - the same agent (the one the
// skip swallowed) and a different one.
//
// Its own file because a harness() call is one HOME, one office URL and one
// extension instance (own-identity-harness.mts). The approval runs through
// office_connect, which is how /office approves, and the assertion is the stored
// file rather than the in-process flag: the file is what the next start reads.
import assert from "node:assert/strict"
import { copyFileSync, readFileSync } from "node:fs"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// Never open a real browser from a test: the opener is the one boundary that
// reaches outside this process (az-iar/ao-pi#10).
process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

// Wait for a state the approval is expected to reach, rather than for a sleep:
// the failure mode this guards is a write that never happens, so the wait has to
// end in a failed assertion with the record still wrong (the reference's
// ao-opencode#41 case waits the same way).
const until = async (check: () => boolean, what: string) => {
  for (let i = 0; i < 250; i++) {
    if (check()) return
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
  assert.fail(`timed out waiting for ${what}`)
}

test("a re-approval after a refusal replaces the refused record instead of skipping the write", async () => {
  const h = await harness({})
  const sessions: Awaited<ReturnType<typeof h.start>>[] = []
  try {
    for (const [approveAs, staleAgent] of [["pi", "pi"], ["other", "pi"]] as const) {
      const id = `reapprove-session-${approveAs}`
      const owner = `reapprove-owner-${approveAs}`
      // A record written for another session: this session's read refuses it.
      h.lib.saveSessionAgent(owner, staleAgent)
      copyFileSync(h.lib.sessionPointerPath(owner), h.lib.sessionPointerPath(id))
      assert.equal(h.lib.sessionAgentOf(id), undefined, `fixture (${approveAs}): the mismatched record was honoured`)
      h.office.state.agentSlug = approveAs
      const session = await h.start(id)
      sessions.push(session)
      // The approval a person runs: /office <name>, in this session.
      const approval = await session.tools.get("office_connect").execute("call", { agent_name: approveAs })
      const link = approval.content[0].text.match(/https?:\/\/\S+/)?.[0]
      assert.ok(link, String(approval.content[0].text))
      await fetch(link, { redirect: "follow" })
      // The approval's own adoption writes the record; this is the observable
      // that the write has happened, and what a start reads. The stale record
      // already names this agent in the same-agent case, so the session field is
      // what tells an overwritten record from the refused one.
      const storedRecord = () => {
        try {
          return JSON.parse(readFileSync(h.lib.sessionPointerPath(id), "utf8")) as { agent?: string; session?: unknown }
        } catch {
          return undefined
        }
      }
      await until(
        () => storedRecord()?.agent === approveAs && storedRecord()?.session === id,
        `the ${approveAs} re-approval to replace the refused record`,
      )
      const stored = storedRecord()!
      assert.deepEqual(
        { agent: stored.agent, session: stored.session },
        { agent: approveAs, session: id },
        `a re-approval as ${approveAs} left the refused record in place`,
      )
      assert.equal(h.lib.loadAgentPointer(h.lib.sessionPointerPath(id)), approveAs, `the stored record does not name ${approveAs}`)
      assert.equal(h.lib.sessionAgentOf(id), approveAs, `a load after the re-approval does not return ${approveAs}`)
    }
  } finally {
    for (const session of sessions) await session.shutdown()
    await h.office.close()
  }
})
