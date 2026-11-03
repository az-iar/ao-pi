// GH-476: the identity headers ride on the calls the extension already makes.
// Its own file, because the extension is imported once per process and the test
// above it leaves a connected session behind - the same order-dependence that
// bit the child-identity test, fixed the same way.
import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

// The point is what the office receives, so this asserts on the wire rather
// than on a function's return value.
test("an approved session sends its own agent on the session-agent header", async () => {
  const h = await harness({})
  const credentials = { office: h.url, client_id: "client", access_token: "token-host", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("ada"), credentials)
  h.lib.saveAgentPointer(h.lib.sessionPointerPath("session-ada"), "ada")
  h.lib.saveAgentPointers(h.repo, h.repo, "ada")

  const session = await h.start("session-ada")
  // A booting session may make no Office call of its own before the first poll,
  // so ask for one the way a person would: /office with no arguments reads the
  // session's own state, which is a real call through the same client.
  const command = session.commands.get("office")
  assert.ok(command, "the /office command is registered")
  await command.handler("", { ui: { notify() {} }, sessionManager: { getSessionId: () => "session-ada" }, isIdle: () => true })

  const carried = h.office.calls
    .filter((c) => c.path === "/mcp")
    .map((c) => c.headers?.["x-office-session-agent"])
  assert.ok(carried.length > 0, "no Office call reached the fake office")
  assert.deepEqual([...new Set(carried)], ["ada"], `the header carried ${JSON.stringify([...new Set(carried)])}`)
  // And on the registration, which is the snapshot a tool call inherits: a header
  // that only reached the client's own calls would leave every Office tool call
  // unlabelled, which is the gap bruno found in ao-opencode#24.
  const config = session.servers.get("office")
  assert.ok(config, "the office server was not registered")
  assert.equal(config.headers["X-Office-Session-Agent"], "ada")
  await session.shutdown()
  await h.office.close()
})

// sunny's P1 on ao-pi#12: the pointer chooses the *header* while the token
// chooses the *bearer*, and under AGENT_OFFICE_TOKEN those are two different
// agents. A resumed session whose pointer names ada, running with another
// agent's token, must declare no agent at all - the header would name an agent
// the call is not acting as, and since #508 went live prod answers that with a
// 409 on every call rather than a warning.
//
// Asserted on the registered config, which is where this session's Office calls
// get their headers: a token session makes no call of its own before the first
// poll, and the registration is the snapshot a tool call inherits.
test("a session running on an environment token names no agent, even with a pointer", async () => {
  const h = await harness({ AGENT_OFFICE_TOKEN: "token-someone-else" })
  const credentials = { office: h.url, client_id: "client", access_token: "token-someone-else", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  h.lib.saveCredentials(h.lib.agentCredentialsPath("ada"), credentials)
  // The stale pointer: it names ada, and nothing in this process acts as ada.
  h.lib.saveAgentPointer(h.lib.sessionPointerPath("session-ada"), "ada")
  h.lib.saveAgentPointers(h.repo, h.repo, "ada")

  const session = await h.start("session-ada")
  const config = session.servers.get("office")
  assert.ok(config, "the office server was not registered")
  assert.equal(config.headers["X-Office-Session-Agent"], undefined, `a token-backed session declared ${config.headers["X-Office-Session-Agent"]}`)
  assert.match(String(config.headers.Authorization), /^!/, "the bearer still comes from the token command")
  await session.shutdown()
  await h.office.close()
})
