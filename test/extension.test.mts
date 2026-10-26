import assert from "node:assert/strict"
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { fakeOffice } from "./fake-office.mts"

async function waitFor(check: () => boolean, timeoutMs = 2_000) {
  const start = Date.now()
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error("timed out waiting for the extension")
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
}

test("the extension registers the office MCP server and schedules its refresh", async () => {
  const office = fakeOffice()
  office.state.unreadMessages = []
  office.state.openTasks = []
  office.state.unclaimedTasks = []
  const url = await office.listen()
  const home = mkdtempSync(join(tmpdir(), "ao-extension-"))
  const workspace = join(home, "repo")
  mkdirSync(workspace)
  process.chdir(workspace)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_POLL_MS = "60000"
  delete process.env.AGENT_OFFICE_TOKEN
  delete process.env.AGENT_OFFICE_CHILD
  delete process.env.AGENT_OFFICE_AUTODISPATCH
  const credentials = join(home, "credentials.json")
  writeFileSync(
    credentials,
    JSON.stringify({ office: url, client_id: "client-1", access_token: "access-1", refresh_token: "refresh-1", expires_at: Date.now() + 3_600_000 }),
  )
  process.env.AGENT_OFFICE_PI_CREDENTIALS = credentials

  const { default: agentOffice } = await import("../index.ts")
  const handlers = new Map<string, ((event: unknown, ctx: unknown) => unknown)[]>()
  const tools = new Map<string, { execute: () => Promise<{ content: { text: string }[] }> }>()
  const mcp = new Map<string, { config: any; registrations: number }>()
  const api = {
    on(name: string, handler: (event: unknown, ctx: unknown) => unknown) {
      handlers.set(name, [...(handlers.get(name) ?? []), handler])
    },
    registerCommand() {},
    registerTool(tool: { name: string; execute: () => Promise<{ content: { text: string }[] }> }) {
      tools.set(tool.name, tool)
    },
    registerMcpServer(name: string, config: any) {
      const current = mcp.get(name)
      mcp.set(name, { config, registrations: (current?.registrations ?? 0) + 1 })
    },
    unregisterMcpServer(name: string) {
      mcp.delete(name)
    },
    sendUserMessage() {},
  }
  agentOffice(api as any)
  const ctx = { isIdle: () => true, ui: { notify() {} } }
  for (const handler of handlers.get("session_start") ?? []) await handler({}, ctx)
  await waitFor(() => office.calls.some((call) => call.path === "/mcp"))

  const status = async () => JSON.parse((await tools.get("office_status")!.execute()).content[0].text)
  const registered = mcp.get("office")
  assert.ok(registered, "the office MCP server is registered")
  assert.match(registered.config.headers.Authorization, /^!.*token\.mts/)
  assert.equal((await status()).mcpRegistered, true)
  assert.ok((await status()).mcpRefreshAt, "a refresh is scheduled before the token expires")

  for (const handler of handlers.get("agent_start") ?? []) await handler({}, ctx)
  await waitFor(() => office.calls.some((call) => call.body?.params?.arguments?.status === "working"))
  for (const handler of handlers.get("agent_settled") ?? []) await handler({}, ctx)
  await waitFor(() => office.calls.some((call) => call.body?.params?.arguments?.status === "idle"))

  for (const handler of handlers.get("session_shutdown") ?? []) await handler({}, ctx)
  assert.equal((await status()).mcpRefreshAt, undefined, "shutdown clears the refresh timer")
  await office.close()
})
