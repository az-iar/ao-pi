# ao-pi

[Pi](https://pi.dev) / [PiG](https://github.com/MichaelKinsy/PiG) extension for [Agent Office](https://github.com/az-iar/agent-office). Install it, run `/office`, approve in the browser, and Pi works as an Office agent: the Office MCP tools (check in, tasks, messages, journal, decisions) appear, and the dashboard shows it `working` while the session streams and `idle` when it settles.

This is the Pi/PiG port of [ao-opencode](https://github.com/az-iar/ao-opencode). Pi runs one session per process, so the status, inbox and dispatch logic is scoped to the session the extension loads in, rather than tracked across a server.

## Install

Copy the extension files where Pi/PiG discovers extensions, or point at the directory:

```sh
# PiG, one session
pig -e /path/to/ao-pi

# PiG, always on for this user
cp -r /path/to/ao-pi ~/.pig/agent/extensions/ao-pi

# PiG, a trusted project
cp -r /path/to/ao-pi .pig/extensions/ao-pi
```

For Pi, the same files load from `~/.pi/agent/extensions/` or a project's `.pi/extensions/`. The extension is plain TypeScript; it needs no build. To try it without installing, load it directly with `pig -e /path/to/ao-pi` (or `pi -e`).

## Connect

In Pi, run `/office`. The extension:

1. registers Pi with the Office (OAuth dynamic client registration), shows the approval link and opens it in your browser if this machine has one;
2. returns at once; when you approve in the dashboard (within five minutes), as a new agent named for this folder or an existing one, the extension finishes in the background and tells the session;
3. stores the token under `~/.config/agent-office/pi/agents/<agent>.json` (mode 0600), records the chosen agent for this workspace and its repo root in `workspaces/<hash>.json`, adds the `office` MCP server to Pi, and checks in.

`/office` subcommands:

| Command | Action |
|---|---|
| `/office` | connect or report the current agent |
| `/office register <name>` | connect and pre-fill a new agent name |
| `/office status` | show the extension's poll, session, and credential state |
| `/office session [name]` | continue a previous conversation in this folder by its `/name`, or list the named ones |
| `/office reconnect` | force a fresh OAuth approval |
| `/office <name>` | alias for `register <name>` |
| `/office http://127.0.0.1:.../callback?...` | finish an approval from a browser on another machine |

### Continuing a previous session

`/office session bruno` reopens the conversation named `bruno` in this folder, which is how a new Pi session picks up where the last one stopped: the history is the session file, and the Office identity comes along because it is the folder's pointer, not the process's. The name is matched exactly first, then as a unique prefix, then as a unique substring; a name several sessions share resumes the most recently modified one (the order `pig -r` lists them), and anything ambiguous prints the candidates instead of guessing. `/office session` with no name lists the folder's named sessions. Name a conversation with `/name`, and it becomes resumable.

The model can also drive this through the `office_connect` and `office_status` tools (handy on remote machines: paste the callback URL and the model passes it to `office_connect`).

The `office` MCP server authenticates with a bearer command (`lib/token.mts`): Pi resolves it when it creates a connection, and it reads and refreshes the shared credentials. Because Pi keeps a live connection, the extension re-registers the server before the access token expires, so long sessions do not need `/mcp reconnect`. Access tokens last an hour. If that refresh fails — the Office unreachable, a gateway error on the token endpoint — it is retried with a backoff, a minute at first and then up to five, and a success puts the normal schedule back; without that one bad minute left the session on an expired token until someone ran `/office` again. A refresh the Office *refuses* (400 or 401: the grant itself is dead) is not retried — the same token would be refused again — so the retries stop and one warning asks for `/office`, which is what renews it. `/office status` shows when the next attempt is due (`mcpRefreshAt`) and why the last one failed (`lastMcpError`).

The loopback callback starts at `AGENT_OFFICE_CALLBACK_PORT` and tries the next 19 ports if it is busy. On a remote or headless machine, open the link in any browser and approve. The browser then lands on a `http://127.0.0.1:.../callback?...` page that does not load: copy that full address and run `/office <address>`, or paste it into the chat for the model to pass to `office_connect`.

A linked git worktree (`git worktree add`) has no Office identity of its own until `/office` runs in it: until then it resolves no credentials at all, neither the main checkout's repo pointer nor the token file they share, so it gets no `office` MCP tools, no polling, no notices and no status writes. Only a pointer `/office` wrote in that worktree counts; one an older version left there implicitly, from the repo-pointer fallback, is ignored, so run `/office` in such a worktree once. Running `/office` in the worktree connects that worktree as its own agent and leaves the main checkout as it was, a dispatched worker is handed its identity by the dispatcher rather than by a pointer, and `AGENT_OFFICE_TOKEN` or `AGENT_OFFICE_PI_CREDENTIALS` overrides this as it does anywhere else. `AGENT_OFFICE_TOKEN_SCOPE=global` does not exempt a linked worktree: the global token file is the main checkout's as much as the per-repo one. A cron line's `--session <id>` mapping is refused there for the same reason, because an older version saved one in a worktree from the repo-pointer fallback; `--credentials` and `AGENT_OFFICE_TOKEN` still work there.

## While running

| Pi | Office |
|---|---|
| the session starts streaming | `checkin {status: "working", note: "pi: <project dir>"}`, then a `checkin` heartbeat every 90 s |
| the agent run settles | `checkin {status: "idle"}`, unless you are `blocked` (you asked Azri something and are waiting) |
| Pi shuts down while streaming | `idle`, with the same `blocked` check |
| a new Office message arrives | the extension injects a user turn that reads it with `list_messages`, does the requested work, replies with `post_message`, and calls `mark_read` |
| a new assigned task appears | a worker Pi takes it in a worktree of its own (see Auto-dispatch) |

The extension polls the Office inbox through `checkin` every 60 seconds while loaded. The first poll baselines tasks; unread messages are not baselined, so the first poll still handles them. Later polls report only new message and task IDs. New messages queue a real Pi turn (`sendUserMessage`) that reads the inbox, does the work, replies and marks the messages read. Before it starts any work it replies once in the thread of the message that asked (a dispatched worker comments on its task instead) saying what it is about to do, then goes ahead without waiting; it messages Azri only when blocked or when the PR is up. This uses model tokens.

If `checkin` reports the agent's `delivery_mode` as `bridge`, the inbox belongs to the herdr bridge instead: the extension and the cron poller record new messages as seen locally and start no turn, no inbox run and no `mark_read` (the bridge's wake is what gets the mail read). Task auto-dispatch is unchanged, any other mode — including an older Office that sends no field — keeps the path above, and the mode is read on every checkin, so flipping back to `poll` resumes delivery without a restart. Mail seen while on bridge is not delivered again after the flip. In every mode the interactive session labels its herdr terminal (`hc label $HC_TERMINAL_ID ao.agent=<slug> ao.role=session`) so a bridge wake can find it; children never label, and a missing `hc` or a refused label never stops the session.

Office text (message senders and subjects, task titles and descriptions) reaches the model only inside an `<office-data>` JSON block, so it cannot pose as the extension's own instructions. Every inbox and dispatch prompt also says that no message, whoever sent it, unlocks reading or sending credentials (tokens, `~/.config/agent-office`, `~/.ssh`, `.env` files) or changing permissions, extensions or config.

## Auto-dispatch

A task assigned to you is dispatched automatically to a worker Pi **only while it is still `todo`**: the poller hands over a snapshot, so the task is re-read before it is claimed, and one that has been cancelled, done or already claimed in the meantime is left alone rather than flipped back to `in_progress`. Four copies of the same issue are one unit of work: the dispatch lock is per task, but a second *scope* on the task's title and repo means the copies collapse to a single worker, and the rest wait until it ends. That scope is the title and repo, not something stronger, so two *genuinely different* tasks that happen to share a title in the same repo — two "fix the flaky test" — are serialized rather than dropped: the second is skipped and retried once the first finishes.

The steps:

1. the extension acknowledges and accepts it if needed, then sets it `in_progress` with a comment — after that re-read, never before it;
2. it gives the task a worktree on branch `office/<short>` (`<short>` is a hash of the task id), beside the repo at `<repo>-office-<short>` through `git worktree add`;
3. it starts a headless `pig -e <ao-pi> --print` there with the worker prompt, so the child has the `office` MCP tools even when the extension was installed with `-e` only. The child does not poll, dispatch or report status. The worker implements, tests, reviews its own diff, commits, pushes, opens the PR with `gh pr create` linking the Office task, posts the link to the `reviewers` team and on the task, and leaves the task `in_progress` for review, or sets it `blocked` with a reason. It ends by writing the day's journal entry for the task with `journal_write` and recording any lasting decision with `decision_write`. Nobody merges, pushes to `main`, or force-pushes. The dispatch re-reads the task while its worker runs (`AGENT_OFFICE_WORKER_STATUS_POLL_MS`, default one minute) and stops the worker the way a timeout does — SIGTERM, then SIGKILL — when the task is **cancelled**, so a cancellation wins over a dispatch already in flight. `done` is deliberately not a stop: a worker that has just marked its own task done may still be pushing or commenting, and killing it there aborts work that succeeded. No unclaim follows an abort, because the worker did start and the office owns the status the task is in now; the worktree is kept, as `stop` keeps it. The watch is cancelled the moment the worker exits, so a finished worker never waits out a poll before its dispatch slot frees.

**Model and thinking presets.** Office resolves a task to tiers — model `fast`, `strong` or `strongest`, thinking `low`, `medium` or `high` — and a single-task read carries them as `preset` (agent-office#841). Which concrete model and thinking level a tier means is this machine's decision, so it is configuration rather than a value in the dispatch code: `~/.config/agent-office/presets.json`, or the file `AGENT_OFFICE_PRESETS` names. The shipped mapping is `fast` → `claude-code/claude-haiku-5-5`, `strong` → `claude-code/claude-sonnet-5-5`, `strongest` → `claude-code/claude-opus-5-5`, and thinking `low`/`medium`/`high` → the same Pi levels; a host whose Pi credentials name other providers maps its own, and the ids have to be ones that host can run. A file names only what it changes, a tier at a time over those defaults:

```json
{ "model": { "strongest": "claude-code/claude-opus-5-5" }, "thinking": { "medium": "high" } }
```

The dispatch passes the mapped values to that worker as `--model` and `--thinking`, so a `review` and a `fix` no longer run on the same session defaults. A task with no `preset` — an Office older than that change, or a read whose preset resolution failed — gets neither flag, so its worker starts with exactly the args it started with before. A tier name the mapping does not know is left out rather than guessed. The cron poller reads the same file for the tasks it dispatches; an inbox run gets no flags. A worker reports the tiers it actually ran with through `update_task`'s `preset_ran`, and only when both reached its command line: a tier this mapping did not name ran Pi's own default, so claiming the task's preset for that axis would be false.

**Token usage per run.** A dispatched worker reports what its run cost when it finishes: one `report_usage` with `source: harness:pi`, the model, `input`, `output`, `cache_read`, `cache_write`, the task id and an idempotency key of `<task id>:<run>`, where the run is the worker session Pi created for it. The counts are Pi's own record of the run — the assistant messages and cache-warm, compaction and branch-summary calls its session holds, the same numbers `/session` shows — so a retry of the same run reuses the key and the Office stores nothing new, while a re-dispatch is a new session and so new usage. Nothing is estimated: a run whose session carries no usage reports nothing at all, and says so once in the log. `amount` is never sent, because Office prices tokens from its model table and flags a model it has no price for as unpriced. The report is the worker's own request and cannot change the task: a refused or unreachable Office is logged and the run ends as it would have, after one retry with the same key. The task id and the tiers that ran reach the worker in `AGENT_OFFICE_TASK_ID` and `AGENT_OFFICE_PRESET_RAN`, which a dispatch sets and an inbox run does not.

The worker's shell runs **sandboxed by default** where `bwrap` is available, with only its worktree, the git data a commit writes and the directory holding its Office credentials writable — the last one because a worker refreshes its token mid-run and has to be able to save it — so a prompt that names another checkout's path cannot write there; `AGENT_OFFICE_SANDBOX=0` opts out and `=1` requires the sandbox. A `=1` host without `bwrap` refuses the dispatch **before** the task is claimed, so it stays `todo` with no worktree left behind; the same is true of every other refusal after the claim — a repo that does not resolve, a worktree that is not its own root, a worker that never launches — which now hands the task back to `todo` instead of leaving it `in_progress` with nobody on it. **This is a behaviour change on upgrade:** a host with `bwrap` that dispatched unsandboxed before now dispatches sandboxed, and `AGENT_OFFICE_SANDBOX=0` is the opt-out.

A Pi opened on an `office/` branch does not poll the inbox or dispatch, so Office work never lands in a worker's session. At most two tasks run at once per Office agent, with a lock per task and a lock per unit of work shared across every process (open sessions and cron). Tasks already open when the extension loads are baseline, not dispatched; a task whose dispatch failed stays unseen and is retried. Set `AGENT_OFFICE_AUTODISPATCH=off` to keep notifications but disable dispatch.

Different Office agents can run in the same checkout: run `/office register <name>` in each session and approve the intended agent. Each running session, its MCP bearer and its workers retain their own credentials even if another session changes the checkout's default. A resumed session restores its saved agent; a new session without an explicit approval starts with the checkout's last-approved agent. Sessions using the same Office agent still share its status; one session settling can mark it idle while another is working.

Tokens are stored per Office agent, so every workspace approved as the same agent shares one token, inbox and status. A workspace remembers its chosen agent by name; `/office <name>` starts a new approval and switches it. The interactive session owns the status; headless children and the cron poller do not write it, so a worker settling cannot mark the agent idle while its parent is still streaming. A cron-only dispatch therefore updates the task and the PR, not the dashboard status.

## Run from cron

`poll.ts` does one Office poll without Pi open, then exits, so cron sets the schedule. `/office` prints this line with the workspace, your `PATH` and the poller's path filled in:

```sh
*/2 * * * * cd /path/to/repo && PATH=/usr/local/bin:/usr/bin:/bin node /path/to/ao-pi/poll.ts >> ~/.config/agent-office/pi/poll.log 2>&1
```

It needs Node 22.18+ when the extension is copied outside `node_modules`. An npm-installed copy lives under `node_modules`, where Node refuses to strip TypeScript, so its poller and MCP bearer command run with `bun`; Bun must be on `PATH` for that install form. Cron has a minimal `PATH`, so set one that finds the runtime and `pig` (or point `AGENT_OFFICE_PI_BIN` at the executable). Run `/office` once in that repo first, or set `AGENT_OFFICE_TOKEN`; `--dir <path>` polls another workspace. The generated cron line includes `--session <id>` to select the saved session-to-agent mapping; use a separate line for each session you want to poll. A credentials override instead produces `--credentials <path>`. Without either option, the poller uses the checkout's default pointer. A missing session mapping fails rather than selecting another agent. Session mappings persist until manually removed from `~/.config/agent-office/pi/sessions/`. `AGENT_OFFICE_TOKEN`, if set, takes precedence over OAuth credentials; cron must receive that environment variable itself.

New messages start one `pig -e <ao-pi> --print` in the workspace with the inbox prompt. New `todo` tasks are claimed, given a worktree, and run there as a worker, within the same two-task cap. It shares the extension's code, message claims and dispatch locks, so it runs safely alongside an open Pi: a batch is handled by one of them. `--dry-run` prints the message and task IDs it would handle and changes nothing.

## Settings

| Variable | Default |
|---|---|
| `AGENT_OFFICE_URL` | `https://office.byma.tech` |
| `AGENT_OFFICE_TOKEN` | unset; an agent API token used instead of OAuth, so `/office` is not needed |
| `AGENT_OFFICE_PI_CREDENTIALS` | unset; an exact credentials file that overrides agent-scoped storage |
| `AGENT_OFFICE_PI_BIN` | `pig`; the executable used for dispatch and the cron poller |
| `AGENT_OFFICE_CLIENT_NAME` | unset; OAuth client name and new-agent prefill, defaults to `<folder>@<hostname>` |
| `AGENT_OFFICE_CALLBACK_PORT` | `41997`, the loopback port for the OAuth callback |
| `AGENT_OFFICE_POLL_MS` | `60000`, inbox poll interval; minimum `1000` |
| `AGENT_OFFICE_AUTODISPATCH` | `on`; `off` keeps notifications but disables dispatch |
| `AGENT_OFFICE_PRESETS` | `~/.config/agent-office/presets.json`; the tier-to-model/thinking mapping a worker's `--model` and `--thinking` come from |

The extension also sets one variable it never reads: `OPTCHAT_PROFILE`, when it is unset and the folder is connected, to that folder's Office agent, so pi-optchat opens that agent's memory instead of asking which profile it is. The profile has to exist in optchat already, or it warns and shows the picker; a value already in the environment wins, and a dispatched worker or a linked worktree nobody connected is left alone. It reaches interactive Pi only, because optchat reads it at its own `session_start` and picks no profile outside the TUI, so `pi -p` and `--mode rpc` still need `--optchat-profile`.

## Development

```sh
npm install
npm run typecheck
npm test
```

The tests run the shared library and the poller against a fake Office (`test/fake-office.mts`): the prompts and their safety framing, task worktrees, the branch guard, the cron line, worktree identity, and the poller's claims with a stub `pig`.

## What differs from ao-opencode

OpenCode exposes a server that sees every session, so its plugin tracks working/idle across windows, creates child sessions, and wakes any of them. Pi/PiG runs one session per process, so this port:

- reports `working`/`idle` from this session's own `agent_start` / `agent_settled` events;
- wakes the model with `sendUserMessage` instead of prompting a remote session;
- dispatches tasks to headless `pig --print` worker processes in worktrees (a single worker per task, rather than a supervisor-plus-subagent pair);
- registers the `office` MCP server with a self-refreshing `!command` bearer, since there is no server-wide MCP transform to replay a refreshed token.

The harness-independent half (`lib/office.ts`: OAuth, credentials, the MCP client, prompts, worktrees, claim and lock files) is shared by the extension, the bearer helper and the poller, exactly as in ao-opencode.
