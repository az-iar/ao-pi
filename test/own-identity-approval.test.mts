import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// Never open a real browser from a test: the opener is the one boundary that
// reaches outside this process (az-iar/ao-pi#10).
process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

// GH-476
test("a fresh session that approves then owns an identity: it checks in and registers the MCP", async () => {
  const h = await harness({})
  const session = await h.start("fresh-session")
  assert.deepEqual(session.registered, [], "a session with no identity registered before approving")
  assert.deepEqual(h.office.calls.map((c) => c.path), [], "a session with no identity called the office before approving")
  const approval = await session.tools.get("office_connect").execute("call", { agent_name: "fresh" })
  const link = approval.content[0].text.match(/https?:\/\/\S+/)?.[0]
  assert.ok(link, "the approval link is returned")
  await fetch(link, { redirect: "follow" })
  const start = Date.now()
  while (!session.registered.includes("office")) {
    if (Date.now() - start > 5000) throw new Error("the approving session never registered the Office MCP")
    await new Promise((r) => setTimeout(r, 25))
  }
  assert.ok(h.office.calls.some((c) => c.path === "/mcp"), "no MCP request after approval")
  assert.ok(h.office.calls.some((c) => c.auth?.startsWith("Bearer ")), "the approved session never presented its credentials")
  await session.shutdown()
  await h.office.close()
})
