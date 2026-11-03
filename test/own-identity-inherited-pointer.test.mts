import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// GH-476: a session with no identity but an inherited directory-pointer credential
// (the incident shape) is offered an approval, not errored at, and never acts.
test("office_connect and /office with no args offer an approval link to a session without an identity", async () => {
  const h = await harness({})
  const credentials = { office: h.url, client_id: "client", access_token: "token-inherited", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("host"), credentials)
  h.lib.saveAgentPointers(h.repo, h.repo, "host")
  const session = await h.start("inherits-pointer")

  const result = await session.tools.get("office_connect").execute("call", {})
  const text = result.content[0].text
  assert.match(text, /https?:\/\/\S+/, `no approval link: ${text}`)
  assert.doesNotMatch(text, /not connected/i)

  await session.commands.get("office").handler("", { sessionManager: { getSessionId: () => "inherits-pointer" }, ui: { notify: (t: string) => session.notices.push(t) }, isIdle: () => true })
  assert.ok(session.sent.some((m) => /https?:\/\/\S+/.test(m)) || session.notices.some((m) => /approve: https?:\/\//.test(m)), "/office offered no approval link")
  assert.ok(!session.notices.some((m) => /not connected/i.test(m)), "/office errored at the session")

  // Starting an approval talks to the OAuth endpoints, but nothing acts as the
  // inherited agent: no MCP request and no request carrying a bearer.
  assert.deepEqual(h.office.calls.filter((c) => c.path === "/mcp" || c.auth), [], "the session acted on the inherited credential")
  assert.deepEqual(session.registered, [])
  await session.shutdown()
  await h.office.close()
})
