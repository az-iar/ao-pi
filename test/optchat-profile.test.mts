import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// The OptChat memory profile follows the folder's Office agent: pi-optchat keeps one
// memory per profile and reads OPTCHAT_PROFILE at its own session_start, so the
// extension sets it while it loads. One harness for the four cases, because the
// modules under test are imported once per process and read HOME and the office URL
// at import time.
test("the OptChat profile follows the folder's Office agent, and nothing else", async () => {
  const h = await harness({})
  const profile = () => process.env.OPTCHAT_PROFILE

  // pi in vinn's checkout opens vinn's memory instead of asking which profile it is.
  h.lib.saveAgentPointers(h.repo, h.repo, "vinn")
  const connected = await h.start("session-connected")
  assert.equal(profile(), "vinn", "a connected folder did not name its agent's profile")
  await connected.shutdown()

  // A profile named by hand is the opt-out: the folder must not overrule it.
  process.env.OPTCHAT_PROFILE = "cody"
  const kept = await h.start("session-kept")
  assert.equal(profile(), "cody", "the extension replaced an explicit profile")
  await kept.shutdown()
  delete process.env.OPTCHAT_PROFILE

  // A dispatched worker runs as a subagent: the profile it would open is locked by the
  // Pi that started it, and its work must not land in that agent's memory.
  process.env.AGENT_OFFICE_CHILD = "1"
  const child = await h.start("session-child")
  assert.equal(profile(), undefined, "a child took the profile its dispatcher holds")
  await child.shutdown()
  delete process.env.AGENT_OFFICE_CHILD

  // optchat accepts 1-64 lowercase letters, digits, hyphens and underscores; an agent
  // name outside that is left unset rather than set to a profile optchat refuses.
  h.lib.saveAgentPointers(h.repo, h.repo, "Not A Prof!le")
  const refused = await h.start("session-refused")
  assert.equal(profile(), undefined, "an invalid profile name was set")
  await refused.shutdown()
  await h.office.close()
})
