// A manual /office must keep working while a retry is pending: it is the escape
// hatch the retry exists to stop depending on, not something the retry can block.
// /office reconnect re-registers the server itself, so it also has to take the
// refresh schedule back over - the pending retry is replaced by the normal one,
// and the error the failed refresh recorded is gone.
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

test("a manual reconnect still works while a retry is pending, and takes the schedule back", async () => {
  // The retry wait is long on purpose: the manual connect has to happen while one
  // is genuinely pending, not after it has already fired.
  const h = await harness({ AGENT_OFFICE_MCP_REFRESH_MIN_MS: "20", AGENT_OFFICE_MCP_RETRY_MS: "60000" })
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
  h.lib.saveSessionAgent("manual-reconnect", "pi")
  h.lib.saveAgentPointers(h.repo, h.repo, "pi")

  const session = await h.start("manual-reconnect")
  const status = async () => JSON.parse((await session.tools.get("office_status").execute()).content[0].text)
  const tokenCalls = () => h.office.calls.filter((call) => call.path === "/oauth/token").length
  const ctx = () => ({ ui: { notify: (text: string) => session.notices.push(text) }, sessionManager: { getSessionId: () => "manual-reconnect" }, isIdle: () => true })
  try {
    // From here the Office answers the token endpoint with a gateway page, once.
    h.office.state.tokenFailures = 1

    await until(async () => Boolean((await status()).lastMcpError), "the timer's refresh to fail")
    const failed = await status()
    const retryIn = new Date(failed.mcpRefreshAt).getTime() - Date.now()
    assert.ok(retryIn > 0 && retryIn < 120_000, `no retry pending, but the next refresh is in ${retryIn}ms`)

    // The manual path, while that retry is pending: /office reconnect approves
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
    // The next refresh is the normal one again, not the 60s retry it replaced.
    const nextRefresh = new Date(reconnected.mcpRefreshAt).getTime() - Date.now()
    assert.ok(nextRefresh > 60_000, `the retry was not replaced by the normal schedule: next refresh in ${nextRefresh}ms`)
  } finally {
    await session.shutdown()
    await h.office.close()
  }
})
