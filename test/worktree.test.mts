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
// pointer, marked as written on purpose) or hand it an identity on purpose with
// AGENT_OFFICE_TOKEN or AGENT_OFFICE_PI_CREDENTIALS. A pointer an older version
// wrote there implicitly, unmarked, does not connect it.

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
// process can drive several workspaces. A session id can be seeded with an
// approval first: since #476 only a session's own approval (or an explicit
// token) makes it own an identity, and these tests are about the worktree.
let sessions = 0
function session(directory: string, id = `session-worktree-${++sessions}`) {
  const handlers = new Map<string, ((event: unknown, ctx: unknown) => unknown)[]>()
  const mcp = new Map<string, any>()
  const tools = new Map<string, any>()
  const notices: string[] = []
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
  const ctx = { sessionManager: { getSessionId: () => id }, isIdle: () => true, ui: { notify: (text: string) => notices.push(text) } }
  const run = async (name: string) => {
    for (const handler of handlers.get(name) ?? []) await handler({}, ctx)
  }
  return {
    mcp,
    tools,
    notices,
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

const poller = (directory: string) => {
  const pollScript = fileURLToPath(new URL("../poll.ts", import.meta.url))
  const env: Record<string, string> = { ...(process.env as Record<string, string>), HOME: home, AGENT_OFFICE_URL: url, AGENT_OFFICE_POLL_MS: "60000" }
  delete env.AGENT_OFFICE_TOKEN
  delete env.AGENT_OFFICE_PI_CREDENTIALS
  return spawnSync(process.execPath, ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", pollScript, "--dir", directory], {
    env,
    encoding: "utf8",
  })
}

const mcpCalls = () => office.calls.filter((call) => call.path === "/mcp")

async function waitFor(check: () => boolean, timeoutMs = 5_000) {
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

  // A session approved as agent-a somewhere else still does not hand that
  // identity to this worktree: the worktree's own pointer is what connects it.
  const sessionId = "session-unconnected"
  lib.saveAgentPointer(lib.sessionPointerPath(sessionId), "agent-a")
  const worktreeSession = session(worktree, sessionId)
  await worktreeSession.start()
  await worktreeSession.working()
  await new Promise((resolve) => setTimeout(resolve, 150))
  assert.equal(worktreeSession.mcp.size, 0, "no office MCP server in an unconnected worktree")
  assert.ok(worktreeSession.notices.some((text) => text.includes("this git worktree has no Office agent of its own")), "it says so in the worktree's words")
  const status = await worktreeSession.status()
  assert.equal(status.unconnected, true)
  assert.equal(status.pollingAllowed, false)
  assert.notEqual(status.credentialsFile, lib.agentCredentialsPath("agent-a"), "status does not name the main checkout's agent token file")
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
  lib.saveAgentPointers(main, main, "agent-a")
  assert.equal(lib.resolveCredentials(worktree, main).credentials, undefined)

  const worktreeSession = session(worktree)
  await worktreeSession.start()
  const approval = await worktreeSession.tools.get("office_connect").execute("call", { agent_name: "agent-b" })
  const link = approval.content[0].text.match(/https?:\/\/\S+/)?.[0]
  assert.ok(link, "the approval link is returned")
  await fetch(link, { redirect: "follow" })
  await waitFor(() => worktreeSession.mcp.has("office"))

  // The approval writes the worktree's own pointer, marked as written here.
  assert.equal(lib.isUnconnectedWorktree(worktree), false)
  assert.equal(lib.loadMarkedAgentPointer(lib.agentPointerPath(worktree)), "agent-b")
  assert.equal(lib.loadAgentPointer(lib.agentPointerPath(main)), "agent-a", "the worktree did not move the main checkout")
  assert.equal(lib.resolveCredentials(worktree, main).credentialsFile, lib.agentCredentialsPath("agent-b"))
  assert.match(tokenHelper(worktree).stdout, /^Bearer access-/)
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a", "the main checkout keeps its agent")

  const registered = worktreeSession.mcp.get("office")
  assert.ok(registered.headers.Authorization.includes(lib.agentCredentialsPath("agent-b")), registered.headers.Authorization)
  const status = await worktreeSession.status()
  assert.equal(status.unconnected, false)
  assert.equal(status.pollingAllowed, true)
  await waitFor(() => mcpCalls().some((call) => call.auth?.startsWith("Bearer ")))
  await worktreeSession.shutdown()

  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a", "the worktree did not move the main checkout")
})

test("a task worktree a dispatcher adopts resolves that agent, at its root and below it", async () => {
  const { main } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveCredentials(lib.agentCredentialsPath("agent-b"), credentials("access-b"))
  lib.saveAgentPointers(main, main, "agent-a")
  const { path, branch } = await lib.addWorktree(main, "task-identity")
  try {
    assert.equal(lib.linkedWorktreeRoot(path), path)
    assert.equal(lib.isUnconnectedWorktree(path), true, "a fresh task worktree has no identity yet")
    assert.equal(lib.resolveCredentials(path, main).credentials, undefined)

    // A plain write in a worktree -- a poll, an inherited token -- leaves no
    // pointer behind: only an adoption on purpose does.
    lib.saveAgentPointers(path, main, "agent-b")
    assert.equal(lib.loadAgentPointer(lib.agentPointerPath(path)), undefined, "an unmarked write is dropped in a worktree")
    assert.equal(lib.isUnconnectedWorktree(path), true)

    // The dispatcher adopts the worktree by writing its own marked pointer.
    lib.saveAgentPointers(path, main, "agent-b", "dispatch")
    assert.equal(lib.loadMarkedAgentPointer(lib.agentPointerPath(path)), "agent-b")
    assert.equal(lib.isUnconnectedWorktree(path), false)
    assert.equal(lib.resolveCredentials(path, main).credentials?.access_token, "access-b")
    assert.equal(tokenHelper(path).stdout, "Bearer access-b")

    // A subdirectory of that worktree resolves the same agent, never the main
    // checkout's: the repo pointer beside the worktree is agent-a's.
    const nested = join(path, "nested")
    mkdirSync(nested)
    assert.equal(lib.resolveCredentials(nested, main).credentials?.access_token, "access-b")
    assert.equal(tokenHelper(nested).stdout, "Bearer access-b")
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
  const result = poller(worktree)
  assert.equal(result.status, 2, result.stderr)
  assert.match(result.stderr, /is a git worktree with no Office agent of its own; run \/office there/)
  assert.deepEqual(mcpCalls(), [], "no checkin and no dispatch from the worktree")
})

test("a worktree the old plugin gave an implicit pointer to stays unconnected", async () => {
  office.calls.length = 0
  office.state.agentSlug = "agent-a"
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  // What the plugins before this rule wrote on every poll in a worktree: the
  // repo-pointer fallback's agent, under the worktree's own path, unmarked.
  lib.saveAgentPointer(lib.agentPointerPath(worktree), "agent-a")
  assert.equal(lib.loadAgentPointer(lib.agentPointerPath(worktree)), "agent-a", "the old pointer is there")

  // It does not connect the worktree, so the fix applies to worktrees the old
  // plugin already used.
  assert.equal(lib.isUnconnectedWorktree(worktree), true)
  assert.deepEqual(lib.resolveCredentials(worktree, main).credentials, undefined)
  assert.equal(tokenHelper(worktree).status, 1)
  assert.equal(poller(worktree).status, 2)

  // Even a session that owns an identity of its own stays without Office tools
  // there, because the worktree has no agent of its own.
  const sessionId = "session-stale-pointer"
  lib.saveAgentPointer(lib.sessionPointerPath(sessionId), "agent-a")
  const worktreeSession = session(worktree, sessionId)
  await worktreeSession.start()
  await worktreeSession.working()
  await new Promise((resolve) => setTimeout(resolve, 150))
  assert.equal(worktreeSession.mcp.size, 0, "no office MCP server behind a stale pointer")
  assert.deepEqual(mcpCalls(), [], "no checkin, poll or status write")
  await worktreeSession.shutdown()

  // The main checkout is untouched by any of this.
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
  assert.equal(lib.loadAgentPointer(lib.agentPointerPath(main)), "agent-a")
})

test("a marked pointer connects the worktree, and running /office there once is enough", () => {
  const { main, worktree } = repoWithWorktree()
  lib.saveCredentials(lib.agentCredentialsPath("agent-a"), credentials("access-a"))
  lib.saveAgentPointers(main, main, "agent-a")
  lib.saveAgentPointer(lib.agentPointerPath(worktree), "agent-a", "office")
  assert.equal(lib.isUnconnectedWorktree(worktree), false)
  assert.equal(lib.resolveCredentials(worktree, main).credentials?.access_token, "access-a")
  assert.equal(tokenHelper(worktree).stdout, "Bearer access-a")
  // An unmarked write does not take the pointer over: a poll or a token handed
  // in must not quietly re-point a worktree someone connected.
  lib.saveAgentPointers(worktree, main, "agent-c")
  assert.equal(lib.loadMarkedAgentPointer(lib.agentPointerPath(worktree)), "agent-a")
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
})

test("an explicit token still works in an unconnected worktree and leaves no pointer behind", async () => {
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

    // office_connect with a token bearer reports who it is and starts an
    // approval; neither writes a pointer, because neither is an approval that
    // finished here.
    await worktreeSession.tools.get("office_connect").execute("call", {})
    await worktreeSession.tools.get("office_connect").execute("call", { agent_name: "agent-c" })
    assert.equal(lib.loadAgentPointer(lib.agentPointerPath(worktree)), undefined, "office_connect on a token bearer wrote a pointer")
    assert.equal(lib.loadCredentials(lib.agentCredentialsPath("env-agent")), undefined, "office_connect on a token bearer copied a token file")
    await worktreeSession.shutdown()

    // The token was handed in, not approved here: nothing connects the worktree
    // once it is gone.
    assert.equal(lib.loadAgentPointer(lib.agentPointerPath(worktree)), undefined, "a token-only worktree leaves no pointer behind")
    assert.equal(lib.isUnconnectedWorktree(worktree), true)
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
  // An unmarked pointer, the shape every version before this one wrote: it
  // still connects the main checkout, only a worktree demands the mark.
  lib.saveAgentPointers(main, main, "agent-a")
  assert.equal(lib.loadMarkedAgentPointer(lib.agentPointerPath(main)), undefined)

  assert.equal(lib.linkedWorktreeRoot(main), undefined)
  assert.equal(lib.isUnconnectedWorktree(main), false)
  assert.equal(lib.isUnconnectedWorktree(sub), false, "a subdirectory of the main checkout is not a worktree")
  assert.equal(lib.resolveCredentials(main, main).credentials?.access_token, "access-a")
  assert.equal(lib.resolveCredentials(sub, main).credentials?.access_token, "access-a")
  assert.equal(tokenHelper(sub).stdout, "Bearer access-a")

  const plain = join(home, "plain")
  mkdirSync(plain)
  assert.equal(lib.isUnconnectedWorktree(plain), false)
  lib.saveCredentials(lib.credentialsPath(plain, plain), credentials("access-plain"))
  assert.equal(lib.resolveCredentials(plain, plain).credentials?.access_token, "access-plain")

  lib.saveAgentPointer(lib.sessionPointerPath("session-main"), "agent-a")
  const mainSession = session(main, "session-main")
  await mainSession.start()
  await waitFor(() => mcpCalls().length > 0)
  assert.ok(mainSession.mcp.get("office"))
  assert.equal(mcpCalls()[0].auth, "Bearer access-a")
  assert.equal((await mainSession.status()).pollingAllowed, true)
  await mainSession.shutdown()
})
