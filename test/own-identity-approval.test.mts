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
  try {
    assert.deepEqual(session.registered, [], "a session with no identity registered before approving")
    assert.deepEqual(h.office.calls.map((c) => c.path), [], "a session with no identity called the office before approving")
    const approval = await session.tools.get("office_connect").execute("call", { agent_name: "fresh" })
    const link = approval.content[0].text.match(/https?:\/\/\S+/)?.[0]
    assert.ok(link, "the approval link is returned")
    await fetch(link, { redirect: "follow" })
    // Wait for the Office call the approval exists to make, not for the MCP
    // registration. Registering is not downstream of that call: the approval
    // starts polling on its way to connectedSummary(), and poll() registers the
    // server *before* its own check-in. Registration is a few microtasks while
    // the check-in is a request the office has to serve, so the registration
    // lands first and this test used to assert in the milliseconds before any
    // request arrived - red at three files in parallel, green alone by luck
    // (ao-pi#41).
    const checkin = await h.office.waitForCall((c) => c.path === "/mcp", "no MCP request after approval")
    assert.ok(checkin.auth?.startsWith("Bearer "), "the approved session never presented its credentials")
    assert.ok(session.registered.includes("office"), "the approving session never registered the Office MCP")
  } finally {
    // A failed assertion used to leave the fake office listening and the suite
    // waiting on this file for as long as the runner allowed: clean up whether
    // the body succeeded or not (ao-pi#41).
    await session.shutdown()
    await h.office.close()
  }
})
