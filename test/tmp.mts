// Every temp directory a test creates goes through here. A suite run used to
// leave one `ao-*` directory per fixture in the system temp directory - hundreds
// of them, and a reviewer could not tell a leaked fixture from a fresh one - so
// cleanup is now the helper's job rather than each file's.
//
// Every directory is named with one prefix the suite owns. test/tmp-guard.test.mts
// counts those entries before and after a whole run, and a directory some other
// tool left behind cannot be mistaken for one of ours.
//
// Removal runs in an `after` hook, so it happens when a test fails too, not only
// on the green path. A process-exit sweep is the backstop for a file that aborts
// before its hooks run (`--test-force-exit` and a throw both reach it).
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { after, type TestContext } from "node:test"

export const TEMP_PREFIX = "ao-test-"

const owned = new Set<string>()

function remove(dir: string) {
  // `force` covers a directory a test already removed itself; a recursive remove
  // is what a git worktree, a socket or a read-only record needs.
  rmSync(dir, { recursive: true, force: true })
  owned.delete(dir)
}

export function removeAllTempDirs() {
  for (const dir of [...owned]) remove(dir)
}

// A directory owned by the file that asked for it. The hook below runs at the
// end of that file, after its tests and after their own hooks, so a fixture a
// test still uses during another hook is not pulled out from under it.
export function tempDir(label: string): string {
  const dir = mkdtempSync(join(tmpdir(), `${TEMP_PREFIX}${label}-`))
  owned.add(dir)
  return dir
}

// A directory owned by one test, for a file with many cases: it goes away when
// that test ends, pass or fail, so a long file does not hold every fixture at
// once. Hooks run in registration order, so this cleanup runs before any hook the
// test registers after this call: a directory a later hook needs is a `tempDir()`
// (removed at the end of the file), not one of these.
export function ownsTempDir(t: TestContext, label: string): string {
  const dir = tempDir(label)
  t.after(() => remove(dir))
  return dir
}

after(removeAllTempDirs)
process.on("exit", removeAllTempDirs)
