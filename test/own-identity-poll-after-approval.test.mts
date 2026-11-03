import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// GH-476 follow-up. A non-child session that boots without an identity (only a
// directory pointer, which confers nothing on its own) used to return before the
// poll timer was started, so a session that then approved connected but never
// polled until it restarted: no inbox, no heartbeats. The boot path now starts
// the timer for it too, and poll() fails closed until the approval lands.
//
// The assertion counts check-ins, not requests: an approval itself makes one
// check-in on the way to connectedSummary(), so the only thing that proves a
// timer is a SECOND check-in arriving with nobody asking for it.
const checkins = (calls: { path: string; body: any }[]) =>
  calls.filter((c) => c.path === "/mcp" && c.body?.params?.name === "checkin").length

test("a non-child session that approves mid-session starts polling", async () => {
  // A short interval so the timer becomes observable inside the test.
  // The extension clamps the interval to 1s, so the timer must be observed over ~1.5s.
  const h = await harness({ AGENT_OFFICE_POLL_MS: "1000" })
  const credentials = { office: h.url, client_id: "client", access_token: "token-host", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("host"), credentials)
  h.lib.saveAgentPointers(h.repo, h.repo, "host") // directory pointer only: no session pointer, so no ownership

  const session = await h.start("boots-unowned")
  assert.deepEqual(session.registered, [], "a session with no identity registered the Office MCP before approving")
  assert.equal(checkins(h.office.calls), 0, "a session with no identity checked in before approving")

  // /office with no arguments is the command path: it starts an approval and
  // hands the link to the session. Do not await it - it resolves when the
  // approval completes, and we drive that with the redirect below.
  const command = session.commands.get("office")
  assert.ok(command, "the /office command is registered")
  void command.handler("", { ui: { notify() {} }, sessionManager: { getSessionId: () => "boots-unowned" }, isIdle: () => true })

  const deadline = Date.now() + 5000
  let link = session.sent.join("\n").match(/https?:\/\/\S+/)?.[0]
  while (!link && Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 25))
    link = session.sent.join("\n").match(/https?:\/\/\S+/)?.[0]
  }
  assert.ok(link, "the approval link was offered by the session that owns nothing yet")
  await fetch(link, { redirect: "follow" })

  // Let the approval settle: one check-in from connectedSummary() is expected.
  const settle = Date.now() + 5000
  while (checkins(h.office.calls) === 0 && Date.now() < settle) await new Promise((r) => setTimeout(r, 25))
  const afterApproval = checkins(h.office.calls)
  assert.ok(afterApproval >= 1, "the approval never checked in")

  // Now stop asking for anything: only a timer can produce more check-ins. Poll
  // for it with a deadline rather than sleeping a fixed window: the interval is
  // clamped to 1s, and a loaded runner may delay one tick, which must not fail a
  // green change. Nothing else in this test asks, so the wait stays honest.
  const deadline2 = Date.now() + 3000
  while (checkins(h.office.calls) < afterApproval + 1 && Date.now() < deadline2) {
    await new Promise((r) => setTimeout(r, 25))
  }
  const after = checkins(h.office.calls)
  assert.ok(after >= afterApproval + 1, `the session stopped polling after the approval: ${after - afterApproval} periodic check-ins in 3s at a 1s interval`)

  await session.shutdown()
  await h.office.close()
})
