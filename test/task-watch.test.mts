// ao-pi#34: the runner re-reads the task while the worker runs. The point of the
// separate module is that stopping the watch is immediate - vinn on #33: a
// `while (watching) await sleep(poll)` loop cannot see the stop until the sleep
// it is parked in wakes, so every dispatch held its slot for up to a poll after
// the worker had exited.
import assert from "node:assert/strict"
import { test } from "node:test"
import { watchTask } from "../lib/task-watch.ts"

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

test("a stopped watch settles at once, not at the end of a poll", async () => {
  const watch = watchTask({ readStatus: async () => "todo", stops: ["cancelled"], onStop: () => {}, intervalMs: 60_000 })
  const started = Date.now()
  watch.stop()
  await watch.done
  assert.ok(Date.now() - started < 250, "the watch settled without waiting for the poll interval")
})

test("a watch that stopped reads nothing more", async () => {
  let reads = 0
  const watch = watchTask({ readStatus: async () => (reads++, "todo"), stops: ["cancelled"], onStop: () => {}, intervalMs: 10 })
  watch.stop()
  await watch.done
  await sleep(60)
  assert.equal(reads, 0, "a stopped watch never calls back into the office")
})

test("it stops the worker when the task becomes cancelled", async () => {
  const seen: string[] = []
  const watch = watchTask({ readStatus: async () => "cancelled", stops: ["cancelled"], onStop: (s) => seen.push(s), intervalMs: 10 })
  await watch.done
  assert.deepEqual(seen, ["cancelled"], "onStop ran once with the status")
})

test("a task that is done, todo or missing does not stop the worker", async () => {
  for (const status of ["todo", "in_progress", "done", "", undefined]) {
    const seen: string[] = []
    const watch = watchTask({ readStatus: async () => status as any, stops: ["cancelled"], onStop: (s) => seen.push(s), intervalMs: 5 })
    await sleep(40)
    watch.stop()
    await watch.done
    assert.deepEqual(seen, [], `status ${JSON.stringify(status)} must not stop the worker`)
  }
})

test("a failing read is treated as unknown, not as a stop", async () => {
  let n = 0
  const seen: string[] = []
  const watch = watchTask({
    readStatus: async () => {
      n++
      if (n < 3) throw new Error("office down")
      return "cancelled"
    },
    stops: ["cancelled"],
    onStop: (s) => seen.push(s),
    intervalMs: 5,
  })
  await watch.done
  assert.deepEqual(seen, ["cancelled"], "the watch survived the failures and stopped on the real verdict")
  assert.equal(n, 3)
})

test("stopping twice is harmless", async () => {
  const watch = watchTask({ readStatus: async () => "todo", stops: ["cancelled"], onStop: () => {}, intervalMs: 60_000 })
  watch.stop(); watch.stop()
  await watch.done
})
