import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// GH-476
test("a child with AGENT_OFFICE_TOKEN is still refused", async () => {
  const h = await harness({ AGENT_OFFICE_TOKEN: "operator-token", AGENT_OFFICE_CHILD: "1" })
  const child = await h.start("token-child")
  assert.deepEqual(child.registered, [], "a child registered on an inherited AGENT_OFFICE_TOKEN")
  assert.deepEqual(h.office.calls.map((c) => c.path), [], "a child with AGENT_OFFICE_TOKEN called the office")
  await child.shutdown()
  await h.office.close()
})
