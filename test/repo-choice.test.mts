// agent-office#540: a dispatch must run in the task's repository. Both dispatch
// sites used to pass the session's repo to addWorktree, so a task in another
// repo ran in a worktree of the wrong one.
import assert from "node:assert/strict";
import { test } from "node:test";

import { chooseRepo, repoFromRemote, resolveDispatchRepo } from "../lib/office.ts";

test("a checkout is matched by its remote, not its directory name", () => {
  assert.equal(repoFromRemote("https://github.com/az-iar/agent-office.git"), "az-iar/agent-office");
  assert.equal(repoFromRemote("git@github.com:az-iar/ao-pi.git"), "az-iar/ao-pi");
  assert.equal(repoFromRemote("/srv/git/agent-office"), undefined, "a bare path names no owner");
});

test("the same checkout written three ways is one repo", () => {
  // An SSH remote, an HTTPS remote with and without .git, a trailing slash and a
  // case difference are the same checkout; treating them as different would fire
  // the refuse path on every real repository (neo, agent-office#540).
  for (const remote of [
    "git@github.com:az-iar/ao-pi.git",
    "https://github.com/az-iar/ao-pi",
    "https://github.com/az-iar/ao-pi.git",
    "https://github.com/az-iar/ao-pi/",
    "HTTPS://GitHub.com/Az-Iar/AO-Pi.git",
  ]) {
    assert.equal(repoFromRemote(remote), "az-iar/ao-pi", remote);
  }
});

test("the task's repo wins over the session's", () => {
  const chosen = chooseRepo("az-iar/agent-office", [
    { path: "/work/ao-pi", remote: "git@github.com:az-iar/ao-pi.git" },
    { path: "/work/agent-office", remote: "git@github.com:az-iar/agent-office.git" },
  ]);
  assert.deepEqual(chosen, { path: "/work/agent-office" });
});

test("no checkout of that repo is a refusal that names it", () => {
  const chosen = chooseRepo("az-iar/elsewhere", [{ path: "/work/ao-pi", remote: "git@github.com:az-iar/ao-pi.git" }]);
  assert.ok("error" in chosen && chosen.error.includes("az-iar/elsewhere"), JSON.stringify(chosen));
});

test("two checkouts of one repo is a refusal, not a guess", () => {
  const chosen = chooseRepo("az-iar/ao-pi", [
    { path: "/a/ao-pi", remote: "git@github.com:az-iar/ao-pi.git" },
    { path: "/b/ao-pi-copy", remote: "git@github.com:az-iar/ao-pi.git" },
  ]);
  assert.ok("error" in chosen && chosen.error.includes("refusing"), JSON.stringify(chosen));
});

test("linked worktrees of one repository are one checkout", () => {
  // A task worktree shares its origin and git common dir with the main checkout,
  // so without dedupe "two checkouts" fires as soon as any worktree exists and
  // every dispatch is refused (neo, ao-pi#17 review).
  const candidates = [
    { path: "/work/agent-office", remote: "git@github.com:az-iar/agent-office.git", commonDir: "/work/agent-office/.git" },
    { path: "/work/agent-office-office-abc12345", remote: "git@github.com:az-iar/agent-office.git", commonDir: "/work/agent-office/.git" },
  ];
  const chosen = resolveDispatchRepo("/work/ao-pi", "az-iar/agent-office", () => candidates);
  assert.deepEqual(chosen, { path: "/work/agent-office" }, "the main checkout, not its worktree");
});

test("a task repo in another case or as a URL still matches", () => {
  const candidates = [{ path: "/work/agent-office", remote: "git@github.com:az-iar/agent-office.git", commonDir: "/x/.git" }];
  assert.deepEqual(resolveDispatchRepo("/work/ao-pi", "Az-Iar/Agent-Office", () => candidates), { path: "/work/agent-office" });
  assert.deepEqual(resolveDispatchRepo("/work/ao-pi", "https://github.com/Az-Iar/Agent-Office.git", () => candidates), { path: "/work/agent-office" });
});

test("a task naming no repo keeps the session's", () => {
  assert.deepEqual(resolveDispatchRepo("/work/ao-pi", "", () => []), { path: "/work/ao-pi" });
});
