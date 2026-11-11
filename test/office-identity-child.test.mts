// GH-476 / ao-pi#12 P2: the child's via header, which nothing covered. Its own
// file because index.ts reads AGENT_OFFICE_PI_CREDENTIALS at import time and the
// child only reaches an Office identity through the dispatcher's override, so
// the variable has to be in the environment before the extension loads.
import assert from "node:assert/strict"
import { mkdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { test } from "node:test"

import { fakeOffice } from "./fake-office.mts"
import { ownsTempDir, tempDir } from "./tmp.mts"

const office = fakeOffice()
office.state.unreadMessages = []
office.state.openTasks = []
office.state.unclaimedTasks = []
const url = await office.listen()
const home = tempDir("child-via")
const repo = join(home, "repo")
mkdirSync(repo)
process.chdir(repo)
process.env.HOME = home
process.env.AGENT_OFFICE_URL = url
process.env.AGENT_OFFICE_POLL_MS = "60000"
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_AUTODISPATCH

// The dispatcher's grant: a credentials file handed in through the environment.
const override = join(home, "override.json")
writeFileSync(override, JSON.stringify({ office: url, client_id: "client", access_token: "token-child", refresh_token: "refresh", expires_at: Date.now() + 3600000 }))
process.env.AGENT_OFFICE_PI_CREDENTIALS = override
process.env.AGENT_OFFICE_CHILD = "1"
// The launcher's name as a person would type it: spaces, capitals, punctuation.
process.env.AGENT_OFFICE_CHILD_NAME = "  Office Inbox!  "

const lib = await import("../lib/office.ts")
const { default: extension } = await import("../index.ts")

// The parent a child names is the directory pointer it inherited.
lib.saveAgentPointers(repo, repo, "ada")

async function boot(sessionId: string) {
  const servers = new Map<string, any>()
  const handlers = new Map<string, Function>()
  extension({
    on: (name: string, fn: Function) => handlers.set(name, fn),
    registerCommand() {},
    registerTool() {},
    registerMcpServer(name: string, config: any) {
      servers.set(name, config)
    },
    unregisterMcpServer() {},
    sendUserMessage() {},
  } as any)
  const ctx = { sessionManager: { getSessionId: () => sessionId }, ui: { notify() {} }, isIdle: () => true }
  await handlers.get("session_start")!({}, ctx)
  return { config: servers.get("office"), shutdown: () => handlers.get("session_shutdown")!({}, ctx) }
}

test("a child names itself and its parent, normalised the way the office accepts", async () => {
  const session = await boot("child-session-1")
  assert.ok(session.config, "the child registered no office server")
  assert.equal(session.config.headers["X-Office-Via"], "child=office-inbox; parent=ada; runtime=pi")
  // The override means the token is the identity, so no agent is declared: the
  // pointer is not what chose the bearer (sunny's P1 on ao-pi#12).
  assert.equal(session.config.headers["X-Office-Session-Agent"], undefined)
  await session.shutdown()
  await office.close()
})

test("a child with no parent pointer sends no via at all", async (t) => {
  // A second terminal-shaped session with the directory pointer gone: a child
  // with no parent to name must not hand-build a partial header, because the
  // office refuses one that is missing a key.
  const other = ownsTempDir(t, "child-via-none")
  mkdirSync(join(other, "repo"))
  process.chdir(join(other, "repo"))
  const session = await boot("child-session-2")
  assert.ok(session.config, "the child registered no office server")
  assert.equal(session.config.headers["X-Office-Via"], undefined)
  await session.shutdown()
})
