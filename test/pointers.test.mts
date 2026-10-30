import assert from "node:assert/strict"
import { mkdirSync, mkdtempSync } from "node:fs"
import { execFileSync } from "node:child_process"
import { fileURLToPath } from "node:url"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"

// lib/office.ts resolves ~/.config/agent-office from HOME when it loads.
const home = mkdtempSync(join(tmpdir(), "ao-home-"))
process.env.HOME = home
const { agentCredentialsPath, resolveCredentials, saveAgentPointers, saveAgentPointer, sessionPointerPath, loadAgentPointer, saveCredentials, cronLine } = await import("../lib/office.ts")

const credentials = {
  office: process.env.AGENT_OFFICE_URL ?? "https://office.byma.tech",
  client_id: "client-1",
  access_token: "access-1",
  refresh_token: "refresh-1",
  expires_at: Date.now() + 3_600_000,
}

test("approving in a subdirectory connects the repo root and siblings", () => {
  const repo = join(home, "repo")
  const sub = join(repo, "sub")
  const sibling = join(repo, "sibling")
  mkdirSync(sub, { recursive: true })
  mkdirSync(sibling, { recursive: true })
  saveCredentials(agentCredentialsPath("my-agent"), credentials)
  saveAgentPointers(sub, repo, "my-agent")

  assert.equal(resolveCredentials(sub, repo).credentials?.access_token, "access-1")
  assert.equal(resolveCredentials(repo, repo).credentials?.access_token, "access-1")
  assert.equal(resolveCredentials(sibling, repo).credentials?.access_token, "access-1")
})

test("cron selects the session or an explicit credentials override", () => {
  assert.match(cronLine("/repo", "/poll.ts", "/bin", "session-a"), /--session 'session-a'/)
  assert.match(cronLine("/repo", "/poll.ts", "/bin", "session-a", "/secret/file"), /--credentials '\/secret\/file'/)
  assert.doesNotMatch(cronLine("/repo", "/poll.ts", "/bin"), /--session|--credentials/)
})

test("session pointers and pinned bearers remain independent in one checkout", () => {
  const repo = join(home, "shared")
  mkdirSync(repo)
  const a = agentCredentialsPath("agent-a")
  const b = agentCredentialsPath("agent-b")
  saveCredentials(a, credentials)
  saveCredentials(b, { ...credentials, access_token: "access-2" })
  saveAgentPointer(sessionPointerPath("session-a"), "agent-a")
  saveAgentPointer(sessionPointerPath("session-b"), "agent-b")
  saveAgentPointers(repo, repo, "agent-b")
  assert.equal(loadAgentPointer(sessionPointerPath("session-a")), "agent-a")
  assert.equal(loadAgentPointer(sessionPointerPath("session-b")), "agent-b")
  const tokenScript = fileURLToPath(new URL("../lib/token.mts", import.meta.url))
  const env = { ...process.env, HOME: home }
  delete env.AGENT_OFFICE_TOKEN
  const token = execFileSync(process.execPath, [tokenScript, "--dir", repo, "--credentials", a], { env, encoding: "utf8" })
  assert.equal(token, "Bearer access-1")
})
