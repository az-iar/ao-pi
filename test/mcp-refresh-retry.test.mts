// The MCP bearer is refreshed by a timer, so a session that runs for hours keeps
// its Office tools without anyone running /office. When the refresh itself failed
// - a gateway's 502 on the token endpoint, one bad minute - connectMcp returned
// and nothing re-armed the timer: the session was left holding an expired bearer
// for the rest of its life. It now retries with a backoff, and a success puts the
// normal schedule back.
//
// Its own file, because a harness() call is one HOME, one office URL and one
// extension instance (own-identity-harness.mts).
import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// Never open a real browser from a test: the opener is the one boundary that
// reaches outside this process (az-iar/ao-pi#10).
process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

// Wait for a state rather than sleeping a fixed window: the wait ends in a failed
// assertion with the state still wrong, which is what makes it a test (the other
// timer tests in this suite wait the same way).
const until = async (check: () => boolean | Promise<boolean>, what: string, timeoutMs = 10_000) => {
  const deadline = Date.now() + timeoutMs
  while (!(await check())) {
    if (Date.now() > deadline) assert.fail(`timed out waiting for ${what}`)
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
}

test("a failed scheduled refresh is retried, and its success re-arms the normal timer", async () => {
  // Milliseconds, not production timings: this test drives the timer, so the
  // normal delay has to be short and the retry shorter than it.
  const h = await harness({ AGENT_OFFICE_MCP_REFRESH_MIN_MS: "20", AGENT_OFFICE_MCP_RETRY_MS: "40" })
  const credentials = {
    office: h.url,
    client_id: "client-1",
    access_token: "access-0",
    refresh_token: "refresh-0",
    // Just past the early-refresh window: the session boots on this token without
    // renewing it, so the first renewal is the timer's, and the timer's normal
    // wait is this margin rather than 55 minutes.
    expires_at: Date.now() + h.lib.REFRESH_EARLY_MS + 3_000,
  }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("pi"), credentials)
  h.lib.saveSessionAgent("refresh-retry", "pi")
  h.lib.saveAgentPointers(h.repo, h.repo, "pi")

  const session = await h.start("refresh-retry")
  const status = async () => JSON.parse((await session.tools.get("office_status").execute()).content[0].text)
  const tokenCalls = () => h.office.calls.filter((call) => call.path === "/oauth/token").length
  try {
    assert.deepEqual(session.registered, ["office"], "the session did not register the Office MCP server at boot")
    assert.equal(tokenCalls(), 0, "the session renewed a token that had not expired yet")
    assert.ok((await status()).mcpRefreshAt, "no refresh was scheduled for the boot token")

    // From here the Office answers the token endpoint with a gateway page, twice:
    // one failure only shows a retry, two show the backoff keeps retrying.
    h.office.state.tokenFailures = 2

    // The failed attempt: the failure names the status, not a SyntaxError, and a
    // retry is scheduled instead of the schedule simply ending.
    await until(async () => Boolean((await status()).lastMcpError), "the timer's refresh to fail")
    const failed = await status()
    assert.match(failed.lastMcpError, /HTTP 502/, `the refresh failure does not name the status: ${failed.lastMcpError}`)
    assert.ok(tokenCalls() >= 1, "the timer's refresh did not reach the token endpoint")
    assert.ok(failed.mcpRefreshAt, "the failed refresh left nothing scheduled, so nothing would try again")

    // The second retry succeeds, and what it arms is the normal schedule again.
    await until(async () => (await status()).lastMcpError === undefined, "the retry to clear the error")
    const recovered = await status()
    assert.equal(tokenCalls(), 3, `the refresh was attempted ${tokenCalls()} times instead of 3 (two 502s, then success)`)
    assert.equal(recovered.mcpRegistered, true, "the retry did not register the MCP server again")
    const nextRefresh = new Date(recovered.mcpRefreshAt).getTime() - Date.now()
    assert.ok(nextRefresh > 60_000, `the success did not re-arm the normal schedule: next refresh in ${nextRefresh}ms`)
  } finally {
    await session.shutdown()
    await h.office.close()
  }
})
