// agent-office#551: a worker may write only inside its own worktree. The path came
// from the prompt, not from a wrong cwd, so the guard is on the tool call.
import assert from "node:assert/strict";
import { mkdirSync, mkdtempSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

import { outsideWorktree } from "../lib/office.ts";

const root = join(mkdtempSync(join(tmpdir(), "ao-guard-")), "worktree");
mkdirSync(join(root, "internal"), { recursive: true });

test("a path inside the worktree is allowed", () => {
  assert.equal(outsideWorktree(root, join(root, "internal/service.go")), undefined);
  assert.equal(outsideWorktree(root, "internal/service.go"), undefined, "relative paths resolve against the worktree");
});

test("an absolute path into a sibling worktree is refused, and the error names both", () => {
  const sibling = join(root, "..", "agent-office/535-drop-to-azri/internal/service.go");
  const refusal = outsideWorktree(root, sibling);
  assert.ok(refusal && refusal.includes(sibling), String(refusal));
  assert.ok(refusal && refusal.includes(root), "the refusal names the worktree too");
});

test("a .. escape is refused", () => {
  assert.ok(outsideWorktree(root, join(root, "..", "elsewhere.txt")));
});

test("a symlink pointing outside is refused", () => {
  const outside = mkdtempSync(join(tmpdir(), "ao-outside-"));
  writeFileSync(join(outside, "secret.txt"), "x\n");
  symlinkSync(outside, join(root, "escape"));
  assert.ok(outsideWorktree(root, join(root, "escape/secret.txt")), "a symlinked directory is followed by realpath and caught");
});
