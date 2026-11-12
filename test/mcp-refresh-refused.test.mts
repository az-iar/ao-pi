// The retry exists for failures that pass: a gateway's 502, an unreachable
// Office, a timeout. A refresh the Office *refuses* (400/401 on the token
// endpoint) is not one of those - the same refresh token would be refused
// again, so retrying only keeps the session on an expired bearer and repeats a
// warning nobody can act on by waiting. It stops the timer, keeps the error and
// asks once for /office, which is what actually renews the grant.
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

test("a refresh the Office refuses is not retried, and asks for /office once", async () => {
  // A retry wait short enough that a retry would certainly have happened inside
  // the window below: "not retried" is then a measured fact, not a guess.
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
  h.lib.saveSessionAgent("refresh-refused", "pi")
  h.lib.saveAgentPointers(h.repo, h.repo, "pi")

  // warn() is console.warn, which is where a session's log sees it; counting it
  // is how "once" is checked rather than assumed.
  const warnings: string[] = []
  const originalWarn = console.warn
  console.warn = (...args: unknown[]) => void warnings.push(args.map(String).join(" "))
  const session = await h.start("refresh-refused")
  const status = async () => JSON.parse((await session.tools.get("office_status").execute()).content[0].text)
  const tokenCalls = () => h.office.calls.filter((call) => call.path === "/oauth/token").length
  try {
    assert.deepEqual(session.registered, ["office"], "the session did not register the Office MCP server at boot")

    // From here the Office answers the token endpoint with invalid_grant.
    h.office.state.tokenRefusals = 1

    await until(async () => Boolean((await status()).lastMcpError), "the timer's refresh to be refused")
    const failed = await status()
    assert.match(failed.lastMcpError, /HTTP 400/, `the refusal does not name the status: ${failed.lastMcpError}`)
    assert.match(failed.lastMcpError, /expired or revoked/, `the refusal does not keep the Office's reason: ${failed.lastMcpError}`)
    assert.equal(tokenCalls(), 1, "the timer's refresh did not reach the token endpoint")
    assert.equal(failed.mcpRefreshAt, undefined, `a refused refresh still armed a retry: ${failed.mcpRefreshAt}`)

    // Ten retry windows at 40ms: a retry would have come and gone, and it would
    // have succeeded (only the first refusal is queued) - so a second request
    // here is the bug this guards against.
    await new Promise((resolve) => setTimeout(resolve, 400))
    assert.equal(tokenCalls(), 1, `the refused refresh was retried ${tokenCalls() - 1} time(s)`)
    const stillFailed = await status()
    assert.equal(stillFailed.lastMcpError, failed.lastMcpError, "the refusal was dropped from the status")
    assert.equal(stillFailed.mcpRefreshAt, undefined, "something re-armed a refresh after the refusal")

    const asked = warnings.filter((text) => /refused the token refresh/.test(text))
    assert.equal(asked.length, 1, `the refusal was warned ${asked.length} times: ${JSON.stringify(asked)}`)
    assert.match(asked[0], /\/office/, `the warning does not say what to do: ${asked[0]}`)
    assert.match(asked[0], /HTTP 400/, `the warning does not say why: ${asked[0]}`)
  } finally {
    console.warn = originalWarn
    await session.shutdown()
    await h.office.close()
  }
})
