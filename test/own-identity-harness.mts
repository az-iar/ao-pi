// Shared by the own-identity tests. Each test file imports index.ts and lib/office.ts
// once per process, and those read HOME and AGENT_OFFICE_URL at import time, so every
// scenario lives in its own file and calls harness() exactly once.
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { after } from "node:test"
import { fakeOffice } from "./fake-office.mts"

// Shared harness for the cases below: a fresh HOME, a fake office, an extension
// instance whose tools and MCP registrations are visible.
export async function harness(env: Record<string, string | undefined>) {
  const office = fakeOffice()
  office.state.unreadMessages = []
  office.state.openTasks = []
  office.state.unclaimedTasks = []
  const url = await office.listen()
  // The listener is this harness's own handle, so the harness closes it rather than
  // leaving it to the caller's cleanup: `after()` registered while a test runs
  // attaches to that test and runs whether it passes or fails, so a throw after
  // harness() returns and before start() hands back a session still closes it. A
  // caller's `finally` cannot cover that window - it has no session to shut down,
  // and a throw in its own body skips the block entirely, which is how a failing
  // file used to keep the fake office's listener open (ao-pi#45). Callers that
  // close it themselves are unaffected: close() is safe to call twice.
  after(() => office.close())
  const home = mkdtempSync(join(tmpdir(), "ao-own-identity-"))
  const repo = join(home, "repo")
  mkdirSync(repo)
  process.chdir(repo)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_POLL_MS = "60000"
  // OPTCHAT_PROFILE is the host's own environment when the suite runs inside Pi;
  // the extension sets it, so a leftover would decide the outcome of a case.
  for (const k of ["AGENT_OFFICE_TOKEN", "AGENT_OFFICE_CHILD", "AGENT_OFFICE_PI_CREDENTIALS", "AGENT_OFFICE_AUTODISPATCH", "OPTCHAT_PROFILE"]) delete process.env[k]
  for (const [k, v] of Object.entries(env)) if (v !== undefined) process.env[k] = v
  const lib = await import("../lib/office.ts")
  const { default: extension } = await import("../index.ts")
  const start = async (sessionId: string) => {
    const registered: string[] = []
    // The registered config, headers included: the identity headers ride here as
    // well as on the client's calls, and a header missing at registration is the
    // snapshot bug bruno found in ao-opencode#24.
    const servers = new Map<string, any>()
    const sent: string[] = []
    const notices: string[] = []
    const handlers = new Map<string, Function>()
    const tools = new Map<string, any>()
    const commands = new Map<string, any>()
    extension({
      on: (name: string, fn: Function) => handlers.set(name, fn),
      registerCommand: (name: string, def: any) => commands.set(name, def),
      registerTool: (tool: any) => tools.set(tool.name, tool),
      registerMcpServer(name: string, config: any) {
        registered.push(name)
        servers.set(name, config)
      },
      unregisterMcpServer() {},
      sendUserMessage: (text: string) => sent.push(text),
    } as any)
    const ctx = { sessionManager: { getSessionId: () => sessionId }, ui: { notify: (text: string) => notices.push(text) }, isIdle: () => true }
    await handlers.get("session_start")!({}, ctx)
    return { registered, servers, notices, tools, commands, sent, shutdown: () => handlers.get("session_shutdown")!({}, ctx) }
  }
  return { office, url, home, repo, lib, start }
}
