import assert from "node:assert/strict"
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { fakeOffice } from "./fake-office.mts"

test("resumed sessions restore their own agent before MCP registration", async () => {
  const office = fakeOffice()
  const url = await office.listen()
  const home = mkdtempSync(join(tmpdir(), "ao-sessions-"))
  const repo = join(home, "repo")
  mkdirSync(repo)
  process.chdir(repo)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_CHILD = "1"
  delete process.env.AGENT_OFFICE_TOKEN
  delete process.env.AGENT_OFFICE_PI_CREDENTIALS
  const lib = await import("../lib/office.ts")
  const credentials = { office: url, client_id: "client", access_token: "token-a", refresh_token: "refresh", expires_at: Date.now() + 3600000 }
  lib.saveCredentials(lib.agentCredentialsPath("a"), credentials)
  lib.saveCredentials(lib.agentCredentialsPath("b"), { ...credentials, access_token: "token-b" })
  lib.saveAgentPointer(lib.sessionPointerPath("session-a"), "a")
  lib.saveAgentPointer(lib.sessionPointerPath("session-b"), "b")
  lib.saveAgentPointers(repo, repo, "b")
  const { default: extension } = await import("../index.ts")
  for (const id of ["session-a", "session-b"]) {
    const handlers = new Map<string, Function>()
    let header = ""
    extension({ on: (name: string, fn: Function) => handlers.set(name, fn), registerCommand() {}, registerTool() {}, registerMcpServer(_name: string, config: any) { header = config.headers.Authorization }, sendUserMessage() {} } as any)
    const ctx = { sessionManager: { getSessionId: () => id }, ui: { notify() {} }, isIdle: () => true }
    await handlers.get("session_start")!({}, ctx)
    assert.ok(header.includes(lib.agentCredentialsPath(id === "session-a" ? "a" : "b")), header)
    await handlers.get("session_shutdown")!({}, ctx)
  }
  await office.close()
})
