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
  await session.shutdown()
  await h.office.close()
})
