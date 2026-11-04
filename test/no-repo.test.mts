// Azri's decision on card 53828e66: a task with repo "" gets NO worktree.
// Defaulting it to the session's repo is what put a worker into a checkout nobody
// chose (agent-office#551), so the policy is pinned here without spawning.
import assert from "node:assert/strict";
import { test } from "node:test";

import { dispatchTarget, scratchTaskDir } from "../lib/office.ts";

test("a repo-less task gets a scratch directory, never the session's repo", () => {
  const target = dispatchTarget("/home/azri/projects/personal/agent-office", "");
  assert.deepEqual(target, { kind: "scratch" });
});

test("whitespace is not a repo either", () => {
  assert.deepEqual(dispatchTarget("/repo", "   "), { kind: "scratch" });
});

test("a named repo keeps the worktree path", () => {
  assert.deepEqual(dispatchTarget("/repo", "az-iar/ao-pi"), { kind: "worktree", repo: "/repo" });
});

test("the scratch directory is under the agent's state, not under any repo", () => {
  const dir = scratchTaskDir("1efe615b-b1ad-4689-9596-85225573406b");
  assert.ok(dir.includes("/agent-office/pi/scratch/"), dir);
  assert.ok(!dir.includes("/projects/"), dir);
});
