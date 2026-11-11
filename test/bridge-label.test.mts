// Bridge M3's pane identity: a session labels its herdr terminal so a later wake
// can find it, in every delivery mode (Office refuses a flip to bridge without a
// live labelled pane). Never a child, and a refused label must not break the
// connect or the poll that made it.
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

test("the session labels its terminal on connect in every mode, never as a child, and a failed label keeps it polling", async (t) => {
  const bin = ownsTempDir(t, "label-bin")
  const log = join(bin, "labels.jsonl")
  const hc = join(bin, "hc")
  writeFileSync(
    hc,
    `#!${process.execPath}\n` +
      `require("fs").appendFileSync(process.env.HC_LABEL_LOG, JSON.stringify(process.argv.slice(2)) + "\\n")\n` +
      `process.exit(Number(process.env.HC_EXIT ?? 0))\n`,
    { mode: 0o755 },
  )
  chmodSync(hc, 0o755)
  delete process.env.HC_EXIT
  const labels = () =>
    (() => {
      try {
        return readFileSync(log, "utf8")
          .split("\n")
          .filter(Boolean)
          .map((line) => JSON.parse(line) as string[])
      } catch {
        return []
      }
    })()
  const h = await harness({
    AGENT_OFFICE_TOKEN: "token-label",
    AGENT_OFFICE_POLL_MS: "1000",
    PATH: `${bin}:${process.env.PATH}`,
    HC_LABEL_LOG: log,
  })
  const checkins = () => h.office.calls.filter((call) => call.path === "/mcp" && call.body?.params?.name === "checkin").length
  try {
    // Every presence value, bridge included, and no field at all: the label is
    // what makes a flip to bridge possible, so it cannot depend on the mode.
    for (const mode of ["poll", "webhook", "agent37", "bridge", undefined]) {
      const terminal = `term-${mode ?? "missing"}`
      process.env.HC_TERMINAL_ID = terminal
      h.office.state.deliveryMode = mode
      const session = await h.start(`label-${mode ?? "missing"}`)
      await waitFor(() => labels().some((argv) => argv[1] === terminal), `a label for ${mode ?? "a missing delivery_mode"}`)
      assert.deepEqual(labels().find((argv) => argv[1] === terminal), ["label", terminal, "ao.agent=pi", "ao.role=session"])
      await session.shutdown()
    }

    // A child would receive the agent's wakes if it carried the session label.
    const childTerminal = "term-child"
    const child = h.lib.labelSessionPane("pi", { ...process.env, AGENT_OFFICE_CHILD: "1", HC_TERMINAL_ID: childTerminal })
    assert.deepEqual({ labeled: child.labeled, reason: child.reason }, { labeled: false, reason: "child" })
    assert.ok(!labels().some((argv) => argv[1] === childTerminal), "a child labelled itself as the session")

    // A failing hc is logged and never thrown: the session still connects and
    // keeps polling afterwards.
    process.env.HC_EXIT = "1"
    process.env.HC_TERMINAL_ID = "term-fails"
    h.office.state.deliveryMode = "poll"
    let failed
    assert.doesNotThrow(() => {
      failed = h.lib.labelSessionPane("pi")
    })
    assert.deepEqual({ labeled: failed!.labeled, reason: failed!.reason }, { labeled: false, reason: "failed" })

    const warnings: string[] = []
    const originalWarn = console.warn
    console.warn = (...args: unknown[]) => warnings.push(args.map(String).join(" "))
    let session: Awaited<ReturnType<typeof h.start>> | undefined
    const checkinsBefore = checkins()
    try {
      session = await h.start("label-fails")
      await waitFor(() => warnings.some((line) => line.includes("could not label this session's terminal")), "the failed label to be logged")
      const status = JSON.parse((await session.tools.get("office_status")!.execute()).content[0].text)
      assert.equal(status.pollingAllowed, true, "a failed label stopped the session polling")
      await waitFor(() => checkins() > checkinsBefore, "the session to check in after its label failed")
    } finally {
      console.warn = originalWarn
      await session?.shutdown()
    }
  } finally {
    await h.office.close()
  }
})
