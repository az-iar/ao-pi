// A failed label is retried, not remembered: a timed-out hc or a restarting
// herdr holder must not leave a live session unlabelled for the rest of its life.
// Only a final outcome (labelled, missing hc, no terminal) stops the retry, and a
// run of failures warns once until the error text changes.
import assert from "node:assert/strict"
import { chmodSync, readFileSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { test } from "node:test"
import { harness } from "./own-identity-harness.mts"
import { ownsTempDir } from "./tmp.mts"

const waitFor = async (check: () => boolean, what: string, timeoutMs = 10_000) => {
  const start = Date.now()
  while (!check()) {
    if (Date.now() - start > timeoutMs) throw new Error(`timed out waiting for ${what}`)
    await new Promise((resolve) => setTimeout(resolve, 25))
  }
}

type Attempt = { argv: string[]; ok: boolean }

test("a transient label failure is retried until it succeeds, with one warning", async (t) => {
  const bin = ownsTempDir(t, "label-retry")
  const log = join(bin, "attempts.jsonl")
  const counter = join(bin, "count")
  const hc = join(bin, "hc")
  writeFileSync(
    hc,
    `#!${process.execPath}\n` +
      `const fs = require("fs")\n` +
      `let n = 0\n` +
      `try { n = Number(fs.readFileSync(${JSON.stringify(counter)}, "utf8")) } catch {}\n` +
      `n++\n` +
      `fs.writeFileSync(${JSON.stringify(counter)}, String(n))\n` +
      `const ok = n > Number(process.env.HC_FAIL_TIMES ?? 0)\n` +
      `fs.appendFileSync(${JSON.stringify(log)}, JSON.stringify({ argv: process.argv.slice(2), ok }) + "\\n")\n` +
      `if (!ok) { process.stderr.write("transient label failure\\n"); process.exit(1) }\n` +
      `process.exit(0)\n`,
    { mode: 0o755 },
  )
  chmodSync(hc, 0o755)
  const attempts = () =>
    (() => {
      try {
        return readFileSync(log, "utf8")
          .split("\n")
          .filter(Boolean)
          .map((line) => JSON.parse(line) as Attempt)
      } catch {
        return []
      }
    })()

  const h = await harness({
    AGENT_OFFICE_TOKEN: "token-label-retry",
    AGENT_OFFICE_POLL_MS: "1000",
    PATH: `${bin}:${process.env.PATH}`,
    HC_TERMINAL_ID: "term-retry",
    HC_FAIL_TIMES: "2",
  })
  h.office.state.deliveryMode = "poll"
  h.office.state.openTasks = []
  h.office.state.unreadMessages = []
  const warnings: string[] = []
  const originalWarn = console.warn
  console.warn = (...args: unknown[]) => warnings.push(args.map(String).join(" "))
  let session: Awaited<ReturnType<typeof h.start>> | undefined
  try {
    session = await h.start("label-retry")
    await waitFor(() => attempts().some((attempt) => attempt.ok), "the label to succeed on a later poll")
    const success = attempts().find((attempt) => attempt.ok)
    assert.deepEqual(success?.argv, ["label", "term-retry", "ao.agent=pi", "ao.role=session"])
    assert.ok(attempts().length >= 3, "the failed calls were not retried")
    assert.equal(
      warnings.filter((line) => line.includes("could not label this session's terminal")).length,
      1,
      "the run of identical failures warned more than once",
    )
  } finally {
    console.warn = originalWarn
    await session?.shutdown()
    await h.office.close()
  }
})
