// Shared by the own-identity tests. Each test file imports index.ts and lib/office.ts
// once per process, and those read HOME and AGENT_OFFICE_URL at import time, so every
// scenario lives in its own file and calls harness() exactly once.
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { fakeOffice } from "./fake-office.mts"

// Shared harness for the cases below: a fresh HOME, a fake office, an extension
// instance whose tools and MCP registrations are visible.
export async function harness(env: Record<string, string | undefined>) {
  const office = fakeOffice()
  office.state.unreadMessages = []
  office.state.openTasks = []
  office.state.unclaimedTasks = []
  const url = await office.listen()
  const home = mkdtempSync(join(tmpdir(), "ao-own-identity-"))
  const repo = join(home, "repo")
  mkdirSync(repo)
  process.chdir(repo)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_POLL_MS = "60000"
  for (const k of ["AGENT_OFFICE_TOKEN", "AGENT_OFFICE_CHILD", "AGENT_OFFICE_PI_CREDENTIALS", "AGENT_OFFICE_AUTODISPATCH"]) delete process.env[k]
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
