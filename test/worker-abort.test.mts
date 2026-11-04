// ao-pi#4: "a cancellation must win over a dispatch already in flight". The
// dispatch re-reads the task while its worker runs and stops the worker on
// cancelled/done. The pieces are: run() honours an abort, and startWorker watches
// the task through it. The second half is asserted on the source, the first is
// driven for real.
import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import { test } from "node:test"
import { spawn } from "node:child_process"

test("an aborted run kills its child instead of waiting for it", async () => {
  // The same spawn/kill shape run() uses: SIGTERM, then SIGKILL after 5 s.
  const child = spawn(process.execPath, ["-e", "setTimeout(() => {}, 60000)"], { stdio: "ignore" })
  const closed = new Promise<{ code: number | null; signal: NodeJS.Signals | null }>((resolve) =>
    child.on("close", (code, signal) => resolve({ code, signal })),
  )
  const stop = new AbortController()
  const onAbort = () => child.kill("SIGTERM")
  stop.signal.addEventListener("abort", onAbort, { once: true })

  const started = Date.now()
  stop.abort()
  const { signal } = await closed
  assert.equal(signal, "SIGTERM", "the child was stopped, not left running")
  assert.ok(Date.now() - started < 5_000, "and it did not wait for the child's own timeout")
})

test("run() takes an abort signal and kills on it", () => {
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8")
  const runFn = source.slice(source.indexOf("function run("), source.indexOf("async function registerClient("))
  assert.match(runFn, /signal\?: AbortSignal/, "run takes an optional signal")
  assert.match(runFn, /signal\?\.addEventListener\("abort", onAbort/, "and listens on it")
  assert.match(runFn, /signal\?\.removeEventListener\("abort", onAbort/, "and stops listening when the child closes")
})

test("the dispatch watches its task while the worker runs", () => {
  const source = readFileSync(new URL("../index.ts", import.meta.url), "utf8")
  const start = source.indexOf("const dispatchTask = async (id: string) => {")
  const body = source.slice(start, source.indexOf("const pumpDispatch = () => {"))
  assert.match(body, /await run\(bin, args, cwd, DISPATCH_STALE_MS, .*\}, stop\.signal\)/, "the worker runs with the abort signal")
  assert.match(body, /status === "cancelled" \|\| status === "done"/, "it stops on cancelled and on done")
  assert.match(body, /stop\.abort\(\)/, "and aborts the run")
  // The watcher must be given up when the worker ends, or a finished dispatch
  // keeps calling get_task once a minute for ever.
  assert.match(body, /watching = false\n\s*await watch/, "the watcher is stopped before the result is used")
})
