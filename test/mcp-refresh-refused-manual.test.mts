// A stopped retry must not stop the session: /office is what renews a refused
// grant, and the two must not interfere. After a refusal has left no timer
// armed, /office reconnect still approves, re-registers the server, clears the
// recorded error, and puts the normal refresh schedule back.
//
// Its own file, because a harness() call is one HOME, one office URL and one
// extension instance (own-identity-harness.mts).
import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// Never open a real browser from a test: the opener is the one boundary that
// reaches outside this process (az-iar/ao-pi#10).
process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

const until = async (check: () => boolean | Promise<boolean>, what: string, timeoutMs = 10_000) => {
  const deadline = Date.now() + timeoutMs
  while (!(await check())) {
    if (Date.now() > deadline) assert.fail(`timed out waiting for ${what}`)
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
}

test("a manual reconnect still works after a refused refresh stopped the retry", async () => {
  const h = await harness({ AGENT_OFFICE_MCP_REFRESH_MIN_MS: "20", AGENT_OFFICE_MCP_RETRY_MS: "40" })
  const credentials = {
    office: h.url,
    client_id: "client-1",
    access_token: "access-0",
    refresh_token: "refresh-0",
    // Just past the early-refresh window, so the boot connect does not renew the
    // token and the timer is the first thing that tries to.
    expires_at: Date.now() + h.lib.REFRESH_EARLY_MS + 3_000,
  }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("pi"), credentials)
  h.lib.saveSessionAgent("refused-reconnect", "pi")
  h.lib.saveAgentPointers(h.repo, h.repo, "pi")

  const session = await h.start("refused-reconnect")
  const status = async () => JSON.parse((await session.tools.get("office_status").execute()).content[0].text)
  const tokenCalls = () => h.office.calls.filter((call) => call.path === "/oauth/token").length
  const ctx = () => ({ ui: { notify: (text: string) => session.notices.push(text) }, sessionManager: { getSessionId: () => "refused-reconnect" }, isIdle: () => true })
  try {
    // The Office refuses the grant once: the timer's attempt. Nothing is queued
    // after it, so the session is now sitting on a dead refresh token.
    h.office.state.tokenRefusals = 1

    await until(async () => Boolean((await status()).lastMcpError), "the timer's refresh to be refused")
    const refused = await status()
    assert.match(refused.lastMcpError, /HTTP 400/, `the refusal was not recorded: ${refused.lastMcpError}`)
    assert.equal(refused.mcpRefreshAt, undefined, "a retry was left armed after the refusal")

    // The manual path, with no timer to lean on: /office reconnect approves
    // again, which re-registers the server and re-arms the normal schedule.
    const tokensBefore = tokenCalls()
    void session.commands.get("office")!.handler("reconnect", ctx())
    const link = await (async () => {
      await until(() => /https?:\/\/\S+/.test(session.sent.join("\n")), "the approval link")
      return session.sent.join("\n").match(/https?:\/\/\S+/)![0]
    })()
    await fetch(link, { redirect: "follow" })

    await until(async () => (await status()).lastMcpError === undefined, "the manual reconnect to clear the error")
    const reconnected = await status()
    assert.ok(tokenCalls() > tokensBefore, "the approval never reached the token endpoint")
    assert.equal(reconnected.mcpRegistered, true, "the manual reconnect did not register the MCP server")
    assert.ok(
      session.notices.concat(session.sent).some((text) => /connected/i.test(text)),
      `the manual reconnect reported nothing: ${JSON.stringify(session.notices.concat(session.sent))}`,
    )
    // The normal schedule is back: a fresh token, renewed an hour from now, not
    // the 40ms retry that was not armed at all.
    const nextRefresh = new Date(reconnected.mcpRefreshAt).getTime() - Date.now()
    assert.ok(nextRefresh > 60_000, `the manual reconnect did not re-arm the normal schedule: next refresh in ${nextRefresh}ms`)
  } finally {
    await session.shutdown()
    await h.office.close()
  }
})
