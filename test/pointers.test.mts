import assert from "node:assert/strict"
import { mkdirSync, mkdtempSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"

// lib/office.ts resolves ~/.config/agent-office from HOME when it loads.
const home = mkdtempSync(join(tmpdir(), "ao-home-"))
process.env.HOME = home
const { agentCredentialsPath, resolveCredentials, saveAgentPointers, saveCredentials } = await import("../lib/office.ts")

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
