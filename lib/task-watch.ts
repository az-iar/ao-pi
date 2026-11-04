// Watching a running worker's task (ao-pi#4 / #34): the dispatch reads the task
// before it claims it, and a worker can run for ten minutes after that, so the
// only way a cancellation wins over a dispatch already in flight is to re-read
// the task while it runs.
//
// The watcher keeps its timer so stopping it is immediate: a `while (watching)`
// loop around `await sleep(poll)` cannot see the stop until the sleep it is
// parked in wakes, which holds the dispatch slot for up to a whole poll after the
// worker has already exited (vinn on ao-pi#33).

export type WatchStop = { stop: () => void; done: Promise<void> }

/**
 * Re-read `readStatus` every `intervalMs` until it returns a status in `stops`,
 * then call `onStop` once. `stop()` ends the watch at once - timer cleared, no
 * further reads - and resolves `done`, so a finished worker never waits a poll.
 */
export function watchTask(options: {
  readStatus: () => Promise<string | undefined>
  stops: string[]
  onStop: (status: string) => void
  intervalMs: number
}): WatchStop {
  const { readStatus, stops, onStop, intervalMs } = options
  let timer: ReturnType<typeof setTimeout> | undefined
  let settle: () => void = () => {}
  const done = new Promise<void>((resolve) => (settle = resolve))
  let stopped = false

  const finish = () => {
    if (stopped) return
    stopped = true
    if (timer) clearTimeout(timer)
    timer = undefined
    settle()
  }

  const tick = async () => {
    timer = undefined
    if (stopped) return
    const status = await readStatus().catch(() => undefined)
    if (stopped) return
    if (status !== undefined && stops.includes(status)) {
      onStop(status)
      finish()
      return
    }
    schedule()
  }

  const schedule = () => {
    if (stopped) return
    timer = setTimeout(tick, intervalMs)
    // Never hold a session open for a watcher: the worker's own exit stops it.
    if (typeof timer === "object" && timer && "unref" in timer) (timer as any).unref()
  }

  schedule()
  return { stop: finish, done }
}
