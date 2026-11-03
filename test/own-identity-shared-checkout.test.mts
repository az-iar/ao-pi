import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// GH-476
test("a pointer written by another session in a shared checkout does not confer ownership", async () => {
  const h = await harness({})
  const credentials = { office: h.url, client_id: "client", access_token: "token-other", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("other"), credentials)
  h.lib.saveAgentPointer(h.lib.sessionPointerPath("other-session"), "other")
  h.lib.saveAgentPointers(h.repo, h.repo, "other")
  const resumed = await h.start("different-session")
  assert.deepEqual(resumed.registered, [], "a session registered on another session's pointer")
  assert.match(resumed.notices.join(" | "), /no Office identity/i)
  assert.deepEqual(h.office.calls.map((c) => c.path), [], "a session with only a directory pointer called the office")
  await resumed.shutdown()
  await h.office.close()
})
