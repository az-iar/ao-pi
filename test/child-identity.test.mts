// GH-476: the Office extension used to hold *host* identity, so a pi child
// started from an approved session spoke as that session's agent - rodri's child
// posted as theao. Identity is per session now: a child has its own id, no
// approval of its own, and therefore no Office tools at all.
import assert from "node:assert/strict"
import { test } from "node:test"
import { mkdirSync } from "node:fs"
import { join } from "node:path"
import { fakeOffice } from "./fake-office.mts"
import { ownsTempDir } from "./tmp.mts"

test("a child session with no approval of its own cannot speak as its host", async (t) => {
  const office = fakeOffice()
  const url = await office.listen()
  const home = ownsTempDir(t, "child-identity")
  const repo = join(home, "repo")
  mkdirSync(repo)
  process.chdir(repo)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_CHILD = "1"
  delete process.env.AGENT_OFFICE_TOKEN
  delete process.env.AGENT_OFFICE_PI_CREDENTIALS

  const lib = await import("../lib/office.ts")
  const credentials = { office: url, client_id: "client", access_token: "token-host", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  // The host session is approved as `host`, and the directory pointer names it
  // too. That pointer plus the host's token is exactly what a child inherits and
  // what it must not be able to speak with.
  lib.saveCredentials(lib.agentCredentialsPath("host"), credentials)
  lib.saveAgentPointer(lib.sessionPointerPath("host-session"), "host")
  lib.saveAgentPointers(repo, repo, "host")

  const { default: extension } = await import("../index.ts")
  const start = async (sessionId: string) => {
    const registered: string[] = []
    const notices: string[] = []
    const handlers = new Map<string, Function>()
    extension({
      on: (name: string, fn: Function) => handlers.set(name, fn),
      registerCommand() {},
      registerTool() {},
      registerMcpServer(name: string) { registered.push(name) },
      sendUserMessage() {},
    } as any)
    const ctx = { sessionManager: { getSessionId: () => sessionId }, ui: { notify: (text: string) => notices.push(text) }, isIdle: () => true }
    await handlers.get("session_start")!({}, ctx)
    return { registered, notices, shutdown: () => handlers.get("session_shutdown")!({}, ctx) }
  }

  // The child: same directory, same environment, same pointer files, no approval.
  const child = await start("child-session")
  assert.deepEqual(child.registered, [], "the Office MCP server was registered for a child with no approval")
  const notice = child.notices.join(" | ")
  assert.ok(/child session/i.test(notice) && /\/office/.test(notice), `the child was not told why it cannot post: ${notice}`)
  // The refusal has to be traceable: it names the session and the agent it would
  // have spoken as, because the sender line is unverifiable after the fact.
  assert.match(notice, /child-session/, `the notice does not name the session: ${notice}`)
  assert.match(notice, /\bhost\b/, `the notice does not name the agent it would have spoken as: ${notice}`)
  // Nothing reached the office at all: the fake records every request it serves.
  assert.deepEqual(office.calls.map((c) => c.path), [], "a child with no approval still called the office")
  await child.shutdown()

  // The host session, approved as itself, still registers: the fix is a gate on
  // unapproved children, not a blanket refusal.
  const host = await start("host-session")
  assert.deepEqual(host.registered, ["office"], "the approved host session did not register the Office MCP server")
  await host.shutdown()

  await office.close()
})
