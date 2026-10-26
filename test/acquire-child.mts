// Spawned by dispatch.test.mts: wait for a shared start time, then race one
// acquireDispatch call. Result is printed on stdout.
import { acquireDispatch } from "../lib/office.ts"

const [dir, id, at] = process.argv.slice(2)
while (Date.now() < Number(at)) {}
process.stdout.write(acquireDispatch(dir!, id!))
