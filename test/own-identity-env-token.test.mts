import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// GH-476
test("AGENT_OFFICE_TOKEN is an operator grant for a non-child session only", async () => {
  const h = await harness({ AGENT_OFFICE_TOKEN: "operator-token" })
  const host = await h.start("token-session")
  assert.deepEqual(host.registered, ["office"], "a non-child session with AGENT_OFFICE_TOKEN did not register")
  await host.shutdown()
  await h.office.close()
})
