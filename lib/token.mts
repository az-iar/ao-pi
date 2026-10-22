#!/usr/bin/env node
// MCP bearer helper. Pi runs this as a `!command` Authorization header for the
// `office` MCP server, so every connection gets a fresh token: it reads the
// workspace's Office credentials, refreshes them under the shared lock when
// they are close to expiry, and prints `Bearer <access token>` on stdout.
//
//   node lib/token.mts [--dir <workspace>]
//
// It shares lib/office.ts with the extension and the poller, so the refresh
// and storage match. Tokens are printed only as the header value, never logged.

import { parseArgs } from "node:util"
import { resolve } from "node:path"
import { spawnSync } from "node:child_process"
import { createBearer, resolveCredentials } from "./office.ts"

function canonicalRoot(directory: string) {
  const git = spawnSync("git", ["-C", directory, "rev-parse", "--path-format=absolute", "--git-common-dir"], { encoding: "utf8" })
  const common = git.status === 0 ? git.stdout.trim() : ""
  return common.endsWith("/.git") ? common.slice(0, -"/.git".length) : directory
}

const { values } = parseArgs({ options: { dir: { type: "string" } } })
const directory = resolve(values.dir ?? process.cwd())
const resolved = resolveCredentials(directory, canonicalRoot(directory))
const { bearer } = createBearer(() => resolved.credentialsFile, resolved.credentials)

const token = await bearer().catch(() => undefined)
if (!token) {
  process.stderr.write("[agent-office] no Office token: run /office in Pi first, or set AGENT_OFFICE_TOKEN\n")
  process.exit(1)
}
process.stdout.write(`Bearer ${token}`)
