// Office 6420db85 Bridge M3: an agent whose checkin says delivery_mode "bridge"
// is woken by the host's daemon, so the extension must not also wake the session.
// The mode is read from every checkin, and only the exact value stands the
// plugin down; every other value (and a server with no field) keeps today's path.
import assert from "node:assert/strict"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"

const waitFor = async (check: () => boolean | Promise<boolean>, what: string, timeoutMs = 10_000) => {
  const start = Date.now()
  while (!(await check())) {
    if (Date.now() - start > timeoutMs) throw new Error(`timed out waiting for ${what}`)
    await new Promise((resolve) => setTimeout(resolve, 25))
  }
}

const checkins = (calls: { path: string; body: any }[]) =>
  calls.filter((call) => call.path === "/mcp" && call.body?.params?.name === "checkin").length

const toolNames = (calls: { path: string; body: any }[]) =>
  calls.filter((call) => call.path === "/mcp" && call.body?.method === "tools/call").map((call) => String(call.body?.params?.name ?? ""))

test("a bridge agent stands down on the first poll, records mail as seen, and a flip back to poll wakes only new mail", async () => {
  const h = await harness({ AGENT_OFFICE_TOKEN: "token-bridge", AGENT_OFFICE_POLL_MS: "1000" })
  h.office.state.openTasks = []
  h.office.state.unclaimedTasks = []
  h.office.state.deliveryMode = "bridge"
  h.office.state.unreadMessages = [{ id: "message-bridge-1", from: "azri", subject: "Bridge mail" }]
  const warnings: string[] = []
  const originalWarn = console.warn
  console.warn = (...args: unknown[]) => warnings.push(args.map(String).join(" "))
  let session: Awaited<ReturnType<typeof h.start>> | undefined
  try {
    session = await h.start("bridge-stand-down")
    const status = async () => JSON.parse((await session!.tools.get("office_status")!.execute()).content[0].text)
    // The first poll after connecting already carries a message: it must be
    // recorded as handled rather than handed to the session.
    await waitFor(async () => (await status()).seen.messages >= 1, "the bridge poll to record the message as seen")
    assert.equal(session.sent.length, 0, "a bridge agent was woken by mail on its first poll")
    assert.ok(!toolNames(h.office.calls).includes("mark_read"), "the extension marked bridge mail read on the server")
    await waitFor(() => checkins(h.office.calls) >= 2, "a second bridge poll")
    assert.equal(session.sent.length, 0, "a later bridge poll woke the session")
    assert.equal(
      warnings.filter((line) => line.includes("delivery mode is bridge")).length,
      1,
      "the bridge was logged more than once while the mode did not change",
    )

    // Flipping back resumes delivery on the next poll. The bridge message is
    // still unread on the server, but local seen state keeps it from coming back;
    // only the message that arrived after the flip is handed over.
    h.office.state.deliveryMode = "poll"
    h.office.state.unreadMessages = [...h.office.state.unreadMessages, { id: "message-poll-2", from: "azri", subject: "After flip" }]
    await waitFor(() => session.sent.some((text) => text.includes("message-poll-2")), "the post-flip message to be handed over")
    const prompt = session.sent.join("\n")
    assert.match(prompt, /message:message-poll-2/)
    assert.doesNotMatch(prompt, /message-bridge-1/, "mail recorded while on bridge was delivered after the flip")

    // Every other presence value keeps today's behaviour byte for byte, including
    // an older server that sends no delivery_mode at all.
    for (const mode of ["webhook", "agent37", "poll", undefined]) {
      const id = `message-${mode ?? "missing"}`
      h.office.state.deliveryMode = mode
      h.office.state.unreadMessages = [...h.office.state.unreadMessages, { id, from: "azri", subject: `For ${mode ?? "missing"}` }]
      await waitFor(() => session.sent.some((text) => text.includes(`message:${id}`)), `${mode ?? "a missing delivery_mode"} to be handed over`)
    }
  } finally {
    console.warn = originalWarn
    await session?.shutdown()
    await h.office.close()
  }
})
