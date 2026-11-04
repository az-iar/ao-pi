// agent-office#551: a dispatched worker's shell gets its worktree and the git data
// a commit writes, and nothing else.
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

import { sandboxArgs } from "../lib/office.ts";

const paths = { worktree: "/w/tree", gitDir: "/w/tree/.git", piState: "/home/x/.config/pig/agent", caches: ["/home/x/.npm"] };

test("the mount order puts the read-only exceptions on top of the writable git dir", () => {
  // A later bind overrides an earlier one, so `.git` writable then `config` and
  // `hooks` read-only is the difference between a commit working and a hook
  // running outside the sandbox.
  const args = sandboxArgs(paths);
  const flat = args.join(" ");
  assert.equal(args[0], "--ro-bind", "everything starts read-only");
  const git = flat.indexOf("--bind /w/tree/.git /w/tree/.git");
  const config = flat.indexOf("--ro-bind /w/tree/.git/config /w/tree/.git/config");
  const hooks = flat.indexOf("--ro-bind /w/tree/.git/hooks /w/tree/.git/hooks");
  assert.ok(git > 0 && config > git && hooks > git, `order was ${flat}`);
  assert.ok(flat.includes("--tmpfs /tmp"), "a worker needs a writable temp dir");
  assert.ok(flat.includes("--bind /home/x/.config/pig/agent"), "Pi writes its transcript every turn");
});

// The behavioural tests need bubblewrap and unprivileged user namespaces. They
// skip with a reason rather than passing silently when either is missing.
function bwrapAvailable(): string | undefined {
  if (spawnSync("bwrap", ["--version"]).status !== 0) return "bubblewrap is not installed here";
  if (spawnSync("unshare", ["--user", "--map-root-user", "true"]).status !== 0) return "unprivileged user namespaces are not permitted here";
  return undefined;
}
const unavailable = bwrapAvailable();

test("a write outside the worktree fails, and a commit inside it still works", { skip: unavailable }, () => {
  const root = mkdtempSync(join(tmpdir(), "ao-sbx-"));
  const repo = join(root, "repo");
  const bare = join(root, "bare.git");
  mkdirSync(repo);
  spawnSync("git", ["init", "-q", bare, "--bare"]);
  spawnSync("git", ["-C", repo, "init", "-q"]);
  spawnSync("git", ["-C", repo, "remote", "add", "origin", bare]);
  writeFileSync(join(repo, "a.txt"), "a\n");
  spawnSync("git", ["-C", repo, "add", "-A"]);
  spawnSync("git", ["-C", repo, "-c", "user.email=t@t", "-c", "user.name=t", "commit", "-qm", "init"]);
  const tree = join(root, "tree");
  // A real linked worktree, not a clone: a clone would pass while every dispatch broke.
  assert.equal(spawnSync("git", ["-C", repo, "worktree", "add", "-q", "-b", "office/x", tree]).status, 0);

  // A push writes into the remote, so a bare remote on this machine has to be a
  // writable bind - the test found that, and it is true in production too.
  // Outside the worktree AND outside every writable mount: piState must not be the
  // sandbox root, or the escape target would sit inside a writable bind (this fixture
  // bug is what the first run reported).
  const piState = join(root, "pi-state")
  mkdirSync(piState)
  const elsewhere = mkdtempSync(join(tmpdir(), "ao-outside-"))
  const outside = join(elsewhere, "escape.txt");
  const script = [
    `echo nope > ${outside} 2>&1 || echo OUTSIDE_REFUSED`,
    `echo ok > ${join(tree, "inside.txt")} && echo INSIDE_OK`,
    `cd ${tree} && git add -A && git -c user.email=t@t -c user.name=t commit -qm inside && echo COMMIT_OK`,
    `cd ${tree} && git push -q origin HEAD:refs/heads/inside 2>&1 && echo PUSH_OK`,
  ].join("; ");
  const run = spawnSync("bwrap", [...sandboxArgs({ worktree: tree, gitDir: join(repo, ".git"), piState, caches: [bare] }), "sh", "-c", script], { encoding: "utf8" });
  const out = `${run.stdout}${run.stderr}`;
  assert.match(out, /OUTSIDE_REFUSED/, out);
  assert.match(out, /INSIDE_OK/, out);
  assert.match(out, /COMMIT_OK/, out);
  assert.match(out, /PUSH_OK/, out);
});
