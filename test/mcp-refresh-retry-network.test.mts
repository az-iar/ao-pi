// The retry has to survive the failures that have no HTTP status at all: an
// Office that cannot be reached, a connection dropped mid-request. Those are
// transient - the next attempt can get through - so the backoff keeps going and
// the session recovers on its own, unlike a refused grant (mcp-refresh-refused).
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

test("a refresh that fails on the network is retried until it succeeds", async () => {
  // Two failures before the success: one retry would not tell "retried once" from
  // "kept retrying", and 40ms doubling to 80ms fits well inside the wait below.
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
  h.lib.saveSessionAgent("refresh-network", "pi")
  h.lib.saveAgentPointers(h.repo, h.repo, "pi")

  const session = await h.start("refresh-network")
  const status = async () => JSON.parse((await session.tools.get("office_status").execute()).content[0].text)
  const tokenCalls = () => h.office.calls.filter((call) => call.path === "/oauth/token").length
  try {
    assert.deepEqual(session.registered, ["office"], "the session did not register the Office MCP server at boot")

    // The Office drops the next two token requests without answering: fetch
    // rejects, there is no status to report.
    h.office.state.tokenDrops = 2

    await until(async () => Boolean((await status()).lastMcpError), "the timer's refresh to fail on the network")
    const failed = await status()
    assert.match(failed.lastMcpError, /fetch failed|socket|ECONNRESET/i, `the network failure is not reported as one: ${failed.lastMcpError}`)
    assert.doesNotMatch(failed.lastMcpError, /HTTP \d{3}/, `a dropped connection was reported as an HTTP status: ${failed.lastMcpError}`)
    assert.ok(failed.mcpRefreshAt, "a network failure left nothing scheduled, so nothing would try again")

    // The third attempt gets through, and the normal schedule is back.
    await until(async () => (await status()).lastMcpError === undefined, "the retry to get through")
    const recovered = await status()
    assert.equal(tokenCalls(), 3, `the refresh was attempted ${tokenCalls()} times instead of 3 (two network failures, then success)`)
    assert.equal(recovered.mcpRegistered, true, "the retry did not register the MCP server again")
    const nextRefresh = new Date(recovered.mcpRefreshAt).getTime() - Date.now()
    assert.ok(nextRefresh > 60_000, `the success did not re-arm the normal schedule: next refresh in ${nextRefresh}ms`)
  } finally {
    await session.shutdown()
    await h.office.close()
  }
})
