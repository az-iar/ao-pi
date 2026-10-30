import assert from "node:assert/strict"
import { createHash } from "node:crypto"
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { fakeOffice } from "./fake-office.mts"

async function waitFor(check: () => boolean, timeoutMs = 5_000) {
  const start = Date.now()
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error("timed out waiting for the approval")
    await new Promise((resolve) => setTimeout(resolve, 25))
  }
}

test("approving as a new agent does not overwrite the previous agent's token", async () => {
  const office = fakeOffice()
  office.state.unreadMessages = []
  office.state.openTasks = []
  office.state.unclaimedTasks = []
  const url = await office.listen()
  const home = mkdtempSync(join(tmpdir(), "ao-approval-"))
  const workspace = join(home, "repo")
  const config = join(home, ".config", "agent-office", "pi")
  const agents = join(config, "agents")
  const pointers = join(config, "workspaces")
  mkdirSync(workspace)
  mkdirSync(agents, { recursive: true })
  mkdirSync(pointers, { recursive: true })
  const oldCredentials = { office: url, client_id: "old", access_token: "old-a", refresh_token: "old-a-refresh", expires_at: Date.now() + 3_600_000 }
  writeFileSync(join(agents, "a.json"), JSON.stringify(oldCredentials))
  const key = createHash("sha256").update(workspace).digest("hex").slice(0, 16)
  const pointer = join(pointers, `${key}.json`)
  writeFileSync(pointer, JSON.stringify({ agent: "a" }))

  process.chdir(workspace)
  process.env.HOME = home
  process.env.AGENT_OFFICE_URL = url
  process.env.AGENT_OFFICE_POLL_MS = "60000"
  delete process.env.AGENT_OFFICE_TOKEN
  delete process.env.AGENT_OFFICE_CHILD
  delete process.env.AGENT_OFFICE_AUTODISPATCH
  delete process.env.AGENT_OFFICE_PI_CREDENTIALS

  const { default: agentOffice } = await import("../index.ts")
  const handlers = new Map<string, ((event: unknown, ctx: unknown) => unknown)[]>()
  const tools = new Map<string, { execute: (id: string, params: unknown) => Promise<{ content: { text: string }[] }> }>()
  const api = {
    on(name: string, handler: (event: unknown, ctx: unknown) => unknown) {
      handlers.set(name, [...(handlers.get(name) ?? []), handler])
    },
    registerCommand() {},
    registerTool(tool: { name: string; execute: (id: string, params: unknown) => Promise<{ content: { text: string }[] }> }) {
      tools.set(tool.name, tool)
    },
    registerMcpServer() {},
    unregisterMcpServer() {},
    sendUserMessage() {},
  }
  agentOffice(api as any)
  const ctx = { sessionManager: { getSessionId: () => "approval-session" }, isIdle: () => true, ui: { notify() {} } }
  for (const handler of handlers.get("session_start") ?? []) await handler({}, ctx)
  await waitFor(() => existsSync(join(agents, "pi.json")))
  const approval = await tools.get("office_connect")!.execute("call", { agent_name: "new" })
  const link = approval.content[0].text.match(/https?:\/\/\S+/)?.[0]
  assert.ok(link, "the approval link is returned")
  await fetch(link, { redirect: "follow" })
  await waitFor(() => JSON.parse(readFileSync(join(agents, "pi.json"), "utf8")).access_token !== "old-a")

  assert.equal(JSON.parse(readFileSync(join(agents, "a.json"), "utf8")).access_token, "old-a", "the old agent is untouched")
  assert.notEqual(JSON.parse(readFileSync(join(agents, "pi.json"), "utf8")).access_token, "old-a")
  assert.equal(JSON.parse(readFileSync(pointer, "utf8")).agent, "pi")
  const sessionKey = createHash("sha256").update("approval-session").digest("hex").slice(0, 16)
  assert.equal(JSON.parse(readFileSync(join(config, "sessions", `${sessionKey}.json`), "utf8")).agent, "pi")
  assert.deepEqual(readdirSync(config).filter((file) => file.startsWith("pending-")), [])
  for (const handler of handlers.get("session_shutdown") ?? []) await handler({}, { isIdle: () => true, ui: { notify() {} } })
  await office.close()
})
