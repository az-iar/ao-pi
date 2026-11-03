import assert from "node:assert/strict"
import { execFileSync, spawnSync } from "node:child_process"
import { mkdirSync, mkdtempSync } from "node:fs"
import { createRequire } from "node:module"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { after, test } from "node:test"
import { fileURLToPath } from "node:url"
import { fakeOffice } from "./fake-office.mts"

// index.ts resolves the repo root from `require`, which Pi's loader provides and
// plain node ESM does not. Shim it, so the extension here resolves a worktree
// against the main checkout exactly as it does in Pi.
;(globalThis as { require?: NodeRequire }).require ??= createRequire(import.meta.url)

// An unconnected worktree -- a linked worktree (`git worktree add`) with no
// agent pointer of its own -- must not take on the identity of the main
// checkout it branches from: no credentials, no office MCP tools, no polling,
// no status writes. Connect it with /office (which writes the worktree's own
// pointer) or hand it an identity on purpose with AGENT_OFFICE_TOKEN or
// AGENT_OFFICE_PI_CREDENTIALS.

const office = fakeOffice()
const url = await office.listen()
// The fake Office holds the process open otherwise, and a suite that never
// exits is a hang, not a pass.
after(() => office.close())

// lib/office.ts resolves ~/.config/agent-office from HOME when it loads.
const home = mkdtempSync(join(tmpdir(), "ao-worktree-"))
process.env.HOME = home
process.env.AGENT_OFFICE_URL = url
process.env.AGENT_OFFICE_POLL_MS = "60000"
delete process.env.AGENT_OFFICE_TOKEN
delete process.env.AGENT_OFFICE_PI_CREDENTIALS
delete process.env.AGENT_OFFICE_CHILD
delete process.env.AGENT_OFFICE_AUTODISPATCH
delete process.env.AGENT_OFFICE_TOKEN_SCOPE

const lib = await import("../lib/office.ts")
const { default: agentOffice } = await import("../index.ts")
const tokenScript = fileURLToPath(new URL("../lib/token.mts", import.meta.url))

const credentials = (token: string) => ({
  office: url,
  client_id: "client-1",
  access_token: token,
  refresh_token: `refresh-${token}`,
  expires_at: Date.now() + 3_600_000,
})

const git = (directory: string, args: string[]) =>
  execFileSync("git", ["-C", directory, "-c", "user.name=t", "-c", "user.email=t@t", ...args], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] })

// A real repository with one commit, plus a linked worktree of it.
function repoWithWorktree() {
  const root = mkdtempSync(join(home, "git-"))
  const main = join(root, "repo")
  const worktree = join(root, "repo-worktree")
  mkdirSync(main)
  git(main, ["init", "-q"])
  git(main, ["commit", "-q", "--allow-empty", "-m", "init"])
  git(main, ["worktree", "add", "-q", "-b", "feature", worktree])
  return { root, main, worktree }
}

// The extension loads the workspace it is given through process.cwd(), so one
// process can drive several workspaces.
let sessions = 0
function session(directory: string) {
  const handlers = new Map<string, ((event: unknown, ctx: unknown) => unknown)[]>()
  const mcp = new Map<string, any>()
  const tools = new Map<string, any>()
  const api = {
    on(name: string, handler: (event: unknown, ctx: unknown) => unknown) {
      handlers.set(name, [...(handlers.get(name) ?? []), handler])
    },
    registerCommand() {},
    registerTool(tool: { name: string }) {
      tools.set(tool.name, tool)
    },
    registerMcpServer(name: string, config: any) {
      mcp.set(name, config)
    },
    unregisterMcpServer(name: string) {
      mcp.delete(name)
    },
    sendUserMessage() {},
  }
  const previous = process.cwd()
  process.chdir(directory)
  try {
    agentOffice(api as any)
  } finally {
    process.chdir(previous)
  }
  const ctx = { sessionManager: { getSessionId: () => `session-worktree-${++sessions}` }, isIdle: () => true, ui: { notify() {} } }
  const run = async (name: string) => {
    for (const handler of handlers.get(name) ?? []) await handler({}, ctx)
  }
  return {
    mcp,
    start: () => run("session_start"),
    working: () => run("agent_start"),
    shutdown: () => run("session_shutdown"),
    status: async () => JSON.parse((await tools.get("office_status")!.execute()).content[0].text) as Record<string, any>,
  }
}

const tokenHelper = (directory: string, env: Record<string, string> = {}) =>
  spawnSync(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", tokenScript, "--dir", directory], {
    env: { ...process.env, HOME: home, ...env },
    encoding: "utf8",
  })

const mcpCalls = () => office.calls.filter((call) => call.path === "/mcp")

async function waitFor(check: () => boolean, timeoutMs = 2_000) {
  const start = Date.now()
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error("timed out waiting for the extension")
    await new Promise((resolve) => setTimeout(resolve, 20))
  }
}

test("an unconnected worktree resolves no credentials, registers no office tools and does not poll", async () => {
  office.calls.length = 0
  office.state.agentSlug = "agent-a"
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")

  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
  assert.equal(lib.isUnconnectedWorktree(worktree), true)
  const resolved = lib.resolveCredentials(worktree, main)
  assert.equal(resolved.unconnected, true)
  assert.equal(resolved.credentials, undefined, "the main checkout's agent is not this worktree's")
  assert.notEqual(resolved.credentialsFile, lib.credentialsPath(main, main), "nor is its token file")

  // The MCP bearer helper produces no header there, and says what to do.
  const helper = tokenHelper(worktree)
  assert.equal(helper.status, 1)
  assert.equal(helper.stdout, "")
  assert.match(helper.stderr, /no Office token: run \/office/)
  assert.equal(tokenHelper(main).stdout, "Bearer access-a")

  const worktreeSession = session(worktree)
  await worktreeSession.start()
  await worktreeSession.working()
  await new Promise((resolve) => setTimeout(resolve, 150))
  assert.equal(worktreeSession.mcp.size, 0, "no office MCP server in an unconnected worktree")
  const status = await worktreeSession.status()
  assert.equal(status.unconnected, true)
  assert.equal(status.pollingAllowed, false)
  assert.deepEqual(mcpCalls(), [], "no checkin, poll or status write")
  await worktreeSession.shutdown()

  assert.equal(lib.loadAgentPointer(lib.agentPointerPath(worktree)), undefined, "nothing connected the worktree")
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
})

test("the token file the worktree shares with the main checkout is not a fallback", () => {
  const { main, worktree } = repoWithWorktree()
  // The legacy per-repo file is keyed by the canonical root, so both resolve
  // the same path.
  assert.equal(lib.credentialsPath(worktree, main), lib.credentialsPath(main, main))
  lib.saveCredentials(lib.credentialsPath(main, main), credentials("access-legacy"))
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-legacy")
  assert.equal(lib.resolveCredentials(worktree, main).credentials, undefined)
  assert.equal(tokenHelper(worktree).status, 1)
})

test("/office in a worktree connects that worktree alone", async () => {
  office.calls.length = 0
  office.state.agentSlug = "agent-b"
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveCredentials(lib.agentCredentialsPath("agent-b"), credentials("access-b"))
  lib.saveAgentPointers(main, main, "agent-a")
  assert.equal(lib.resolveCredentials(worktree, main).credentials, undefined)

  // What the approval in the worktree does: it saves the token under the
  // approved agent and writes the worktree's own pointer.
  lib.saveAgentPointers(worktree, main, "agent-b")

  assert.equal(lib.isUnconnectedWorktree(worktree), false)
  assert.equal(lib.resolveCredentials(worktree, main).credentials?.access_token, "access-b")
  assert.equal(tokenHelper(worktree).stdout, "Bearer access-b")
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a", "the main checkout keeps its agent")

  const worktreeSession = session(worktree)
  await worktreeSession.start()
  await waitFor(() => mcpCalls().length > 0)
  const registered = worktreeSession.mcp.get("office")
  assert.ok(registered, "the office MCP server is registered in the connected worktree")
  assert.ok(registered.headers.Authorization.includes(lib.agentCredentialsPath("agent-b")), registered.headers.Authorization)
  const status = await worktreeSession.status()
  assert.equal(status.unconnected, false)
  assert.equal(status.pollingAllowed, true)
  assert.equal(mcpCalls()[0].auth, "Bearer access-b")
  await worktreeSession.shutdown()

  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a", "the worktree did not move the main checkout")
})

test("a task worktree whose own pointer is written resolves that agent", async () => {
  const { main } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  const { path, branch } = await lib.addWorktree(main, "task-identity")
  try {
    assert.equal(lib.linkedWorktreeRoot(path), path)
    assert.equal(lib.isUnconnectedWorktree(path), true, "a fresh task worktree has no identity yet")
    assert.equal(lib.resolveCredentials(path, main).credentials, undefined)

    // The dispatcher adopts the worktree by writing its own pointer.
    lib.saveAgentPointers(path, main, "agent-a")
    assert.equal(lib.isUnconnectedWorktree(path), false)
    assert.equal(lib.resolveCredentials(path, main).credentials?.access_token, "access-a")
    assert.equal(tokenHelper(path).stdout, "Bearer access-a")

    // And a subdirectory of that worktree resolves the same agent.
    const nested = join(path, "nested")
    mkdirSync(nested)
    assert.equal(lib.resolveCredentials(nested, main).credentials?.access_token, "access-a")
  } finally {
    git(main, ["worktree", "remove", "--force", path])
    git(main, ["branch", "-D", branch])
  }
})

test("a dispatched worker in a task worktree gets its identity from the handed-over credentials file", async () => {
  const { main } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  const { path, branch } = await lib.addWorktree(main, "task-worker")
  try {
    assert.equal(lib.isUnconnectedWorktree(path), true)
    // dispatchTask and the poller's runPi hand a worker AGENT_OFFICE_PI_CREDENTIALS
    // and AGENT_OFFICE_CHILD=1 rather than writing a pointer; that is an
    // identity handed over on purpose and still works in a fresh worktree.
    const child = tokenHelper(path, { AGENT_OFFICE_PI_CREDENTIALS: lib.agentCredentialsPath("agent-a"), AGENT_OFFICE_CHILD: "1" })
    assert.equal(child.status, 0, child.stderr)
    assert.equal(child.stdout, "Bearer access-a")
  } finally {
    git(main, ["worktree", "remove", "--force", path])
    git(main, ["branch", "-D", branch])
  }
})

test("the cron poller does nothing for an unconnected worktree", () => {
  office.calls.length = 0
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  // The token file the worktree shares with the main checkout, for the case the
  // poller used to pick it up.
  lib.saveCredentials(lib.credentialsPath(main, main), credentials("access-legacy"))
  const pollScript = fileURLToPath(new URL("../poll.ts", import.meta.url))
  const env: Record<string, string> = { ...(process.env as Record<string, string>), HOME: home, AGENT_OFFICE_URL: url, AGENT_OFFICE_POLL_MS: "60000" }
  delete env.AGENT_OFFICE_TOKEN
  delete env.AGENT_OFFICE_PI_CREDENTIALS
  const result = spawnSync(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", pollScript, "--dir", worktree], {
    env,
    encoding: "utf8",
  })
  assert.equal(result.status, 2, result.stderr)
  assert.match(result.stderr, /is a git worktree with no Office agent of its own; run \/office there/)
  assert.deepEqual(mcpCalls(), [], "no checkin and no dispatch from the worktree")
})

test("an explicit token still works in an unconnected worktree", async () => {
  office.calls.length = 0
  office.state.agentSlug = "env-agent"
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  process.env.AGENT_OFFICE_TOKEN = "token-env"
  try {
    // The environment token is the identity; the main checkout's files still are not.
    assert.equal(lib.isUnconnectedWorktree(worktree), true)
    const resolved = lib.resolveCredentials(worktree, main)
    assert.equal(resolved.credentials, undefined)
    assert.notEqual(resolved.credentialsFile, lib.credentialsPath(main, main), "not the token file the main checkout uses")
    assert.equal(tokenHelper(worktree).stdout, "Bearer token-env")

    const worktreeSession = session(worktree)
    await worktreeSession.start()
    await waitFor(() => mcpCalls().length > 0)
    assert.ok(worktreeSession.mcp.get("office"), "the office MCP server is registered for the handed-over identity")
    assert.equal(mcpCalls()[0].auth, "Bearer token-env")
    const status = await worktreeSession.status()
    assert.equal(status.pollingAllowed, true)
    assert.notEqual(status.credentialsFile, lib.credentialsPath(main, main), "status does not name the main checkout's token file")
    await worktreeSession.shutdown()
  } finally {
    delete process.env.AGENT_OFFICE_TOKEN
  }
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
})

test("the main checkout, its subdirectories and plain directories are unchanged", async () => {
  office.calls.length = 0
  office.state.agentSlug = "agent-a"
  const { main } = repoWithWorktree()
  const sub = join(main, "src")
  mkdirSync(sub)
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")

  assert.equal(lib.linkedWorktreeRoot(main), undefined)
  assert.equal(lib.isUnconnectedWorktree(main), false)
  assert.equal(lib.isUnconnectedWorktree(sub), false, "a subdirectory of the main checkout is not a worktree")
  assert.equal(lib.resolveCredentials(sub, main).credentials?.access_token, "access-a")
  assert.equal(tokenHelper(sub).stdout, "Bearer access-a")

  const plain = join(home, "plain")
  mkdirSync(plain)
  assert.equal(lib.isUnconnectedWorktree(plain), false)
  lib.saveCredentials(lib.credentialsPath(plain, plain), credentials("access-plain"))
  assert.equal(lib.resolveCredentials(plain, plain).credentials?.access_token, "access-plain")

  const mainSession = session(main)
  await mainSession.start()
  await waitFor(() => mcpCalls().length > 0)
  assert.ok(mainSession.mcp.get("office"))
  assert.equal(mcpCalls()[0].auth, "Bearer access-a")
  assert.equal((await mainSession.status()).pollingAllowed, true)
  await mainSession.shutdown()
})
