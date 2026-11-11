// A suite run must leave the system temp directory as it found it. Every leaking
// fixture used to look the same as a fresh one, so nobody noticed the count
// climbing; this test is the count.
//
// It runs the suite once, in a child process, with a temp directory of its own as
// TMPDIR. The isolation matters: this file runs beside the rest of the suite, and
// a before/after count of the shared temp directory would race every other file's
// fixtures. With its own TMPDIR the measured run's whole namespace is one
// directory, `os.tmpdir()` resolves there for it and every process it spawns, and
// "no new top-level entries" is the plain count of what is left inside.
//
// The child is the suite minus this file, so a run cannot recurse into itself.
import assert from "node:assert/strict"
import { spawn } from "node:child_process"
import { readdirSync } from "node:fs"
import { join } from "node:path"
import { test } from "node:test"
import { fileURLToPath } from "node:url"

import { tempDir } from "./tmp.mts"

const testDir = fileURLToPath(new URL(".", import.meta.url))
const repoDir = fileURLToPath(new URL("..", import.meta.url))

type Result = { code: number | null; signal: NodeJS.Signals | null; stdout: string; stderr: string }

function runSuite(files: string[], tmp: string): Promise<Result> {
  return new Promise((resolve, reject) => {
    // The runner sets NODE_TEST_CONTEXT in every test file process, and a nested
    // `node --test` that inherits it treats its own arguments as a worker's and
    // exits without running anything. The compile cache would also put its own
    // directory in this run's TMPDIR, which is not the suite's temp entry to
    // account for.
    const env: NodeJS.ProcessEnv = { ...process.env, TMPDIR: tmp, TMP: tmp, TEMP: tmp, AGENT_OFFICE_CALLBACK_PORT: "0", AGENT_OFFICE_OPEN_BROWSER: "off", NODE_DISABLE_COMPILE_CACHE: "1" }
    delete env.NODE_TEST_CONTEXT
    delete env.NODE_COMPILE_CACHE
    const child = spawn(
      process.execPath,
      ["--experimental-strip-types", "--disable-warning=ExperimentalWarning", "--test", "--test-force-exit", ...files],
      {
        cwd: repoDir,
        // The whole measured run, including the files it spawns, resolves its temp
        // entries inside `tmp`. The opener is switched off for it too, since this
        // child does not go through package.json's test script.
        env,
        stdio: ["ignore", "pipe", "pipe"],
      },
    )
    let stdout = ""
    let stderr = ""
    child.stdout.setEncoding("utf8").on("data", (chunk: string) => (stdout += chunk))
    child.stderr.setEncoding("utf8").on("data", (chunk: string) => (stderr += chunk))
    // Generous, because the whole suite runs inside: a hang must fail here rather
    // than hold the job until CI's own timeout.
    const timer = setTimeout(() => child.kill("SIGKILL"), 10 * 60_000)
    child.on("error", reject)
    child.on("close", (code, signal) => {
      clearTimeout(timer)
      resolve({ code, signal, stdout, stderr })
    })
  })
}

test("a full run leaves no new entry in the temp directory", async () => {
  const files = readdirSync(testDir)
    .filter((name) => name.endsWith(".test.mts") && name !== "tmp-guard.test.mts")
    .sort()
    .map((name) => join(testDir, name))
  assert.ok(files.length >= 10, `only ${files.length} test files found beside ${import.meta.url}`)

  const tmp = tempDir("guard")
  assert.deepEqual(readdirSync(tmp), [], "the measured run starts with an empty temp directory")

  const result = await runSuite(files, tmp)
  const output = `${result.stdout}${result.stderr}`
  assert.equal(result.code, 0, `the measured run failed (${result.signal ?? result.code}):\n${output.slice(-6000)}`)

  const left = readdirSync(tmp).sort()
  assert.deepEqual(left, [], `the run left ${left.length} temp entr${left.length === 1 ? "y" : "ies"} behind: ${left.join(", ")}`)
})
