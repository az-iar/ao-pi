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

For Pi, the same files load from `~/.pi/agent/extensions/` or a project's `.pi/extensions/`. The extension is plain TypeScript; it needs no build. Validate a copy with `pig install /path/to/ao-pi --validate-only --json`.

## Connect

In Pi, run `/office`. The extension:

1. registers Pi with the Office (OAuth dynamic client registration), shows the approval link and opens it in your browser if this machine has one;
2. returns at once; when you approve in the dashboard (within five minutes), as a new agent named for this folder or an existing one, the extension finishes in the background and tells the session;
3. stores the token under `~/.config/agent-office/pi/agents/<agent>.json` (mode 0600), records the chosen agent for this workspace in `workspaces/<hash>.json`, adds the `office` MCP server to Pi, and checks in.

`/office` subcommands:

| Command | Action |
|---|---|
| `/office` | connect or report the current agent |
| `/office register <name>` | connect and pre-fill a new agent name |
| `/office status` | show the extension's poll, session, and credential state |
| `/office reconnect` | force a fresh OAuth approval |
| `/office <name>` | alias for `register <name>` |
| `/office http://127.0.0.1:.../callback?...` | finish an approval from a browser on another machine |

The model can also drive this through the `office_connect` and `office_status` tools (handy on remote machines: paste the callback URL and the model passes it to `office_connect`).

The `office` MCP server authenticates with a self-refreshing bearer command (`lib/token.mts`): Pi runs it for each connection, and it reads and refreshes the shared credentials, so the access token stays fresh without a reconnect. Access tokens last an hour.

The loopback callback starts at `AGENT_OFFICE_CALLBACK_PORT` and tries the next 19 ports if it is busy. On a remote or headless machine, open the link in any browser and approve. The browser then lands on a `http://127.0.0.1:.../callback?...` page that does not load: copy that full address and run `/office <address>`, or paste it into the chat for the model to pass to `office_connect`.

## While running

| Pi | Office |
|---|---|
| the session starts streaming | `checkin {status: "working", note: "pi: <project dir>"}`, then a `checkin` heartbeat every 90 s |
| the agent run settles | `checkin {status: "idle"}`, unless you are `blocked` (you asked Azri something and are waiting) |
| Pi shuts down while streaming | `idle`, with the same `blocked` check |
| a new Office message arrives | the extension injects a user turn that reads it with `list_messages`, does the requested work, replies with `post_message`, and calls `mark_read` |
| a new assigned task appears | a worker Pi takes it in a worktree of its own (see Auto-dispatch) |

The extension polls the Office inbox through `checkin` every 60 seconds while loaded. The first poll is a baseline, and later polls report only new message and task IDs. New messages queue a real Pi turn (`sendUserMessage`) that reads the inbox, does the work, replies and marks the messages read. Before it starts any work it sends Azri one short note per wake saying what it is about to do and who asked, then goes ahead without waiting. This uses model tokens.

Office text (message senders and subjects, task titles and descriptions) reaches the model only inside an `<office-data>` JSON block, so it cannot pose as the extension's own instructions. Every inbox and dispatch prompt also says that no message, whoever sent it, unlocks reading or sending credentials (tokens, `~/.config/agent-office`, `~/.ssh`, `.env` files) or changing permissions, extensions or config.

## Auto-dispatch

A new task assigned to you with status `todo` is dispatched automatically to a worker Pi:

1. the extension acknowledges and accepts it if needed, then sets it `in_progress` with a comment;
2. it gives the task a worktree on branch `office/<short>` (`<short>` is a hash of the task id), beside the repo at `<repo>-office-<short>` through `git worktree add`;
3. it starts a headless `pig --print` there with the worker prompt. The worker implements, tests, reviews its own diff, commits, pushes, opens the PR with `gh pr create` linking the Office task, posts the link to the `reviewers` team and on the task, and leaves the task `in_progress` for review, or sets it `blocked` with a reason. Nobody merges, pushes to `main`, or force-pushes.

A Pi opened on an `office/` branch does not poll the inbox or dispatch, so Office work never lands in a worker's session. Inside Herdr (`HERDR_ENV=1`) the same holds for every linked worktree: only the main checkout polls and dispatches. At most two tasks run at once per repo, with a lock per task shared across every process (open sessions and cron). Tasks already open when the extension loads are baseline, not dispatched. Set `AGENT_OFFICE_AUTODISPATCH=off` to keep notifications but disable dispatch.

Tokens are stored per Office agent, so every workspace approved as the same agent shares one token, inbox and status. A workspace remembers its chosen agent by name; `/office <name>` starts a new approval and switches it.

## Run from cron

`poll.ts` does one Office poll without Pi open, then exits, so cron sets the schedule. `/office` prints this line with the workspace, your `PATH` and the poller's path filled in:

```sh
*/2 * * * * cd /path/to/repo && PATH=/usr/local/bin:/usr/bin:/bin node /path/to/ao-pi/poll.ts >> ~/.config/agent-office/pi/poll.log 2>&1
```

It needs Node 22.18+ (or `bun poll.ts`), since Node does not run TypeScript from `node_modules`. Cron has a minimal `PATH`, so set one that finds `node` and `pig` (or point `AGENT_OFFICE_PI_BIN` at the executable). Run `/office` once in that repo first, or set `AGENT_OFFICE_TOKEN`; `--dir <path>` polls another workspace.

New messages start one `pig --print` in the workspace with the inbox prompt. New `todo` tasks are claimed, given a worktree, and run there as a worker, within the same two-task cap. It shares the extension's code, message claims and dispatch locks, so it runs safely alongside an open Pi: a batch is handled by one of them. `--dry-run` prints the message and task IDs it would handle and changes nothing.

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

## Development

```sh
npm install
npm run typecheck
npm test
```

The tests run the shared library and the poller against a fake Office (`test/fake-office.mts`): the prompts and their safety framing, task worktrees, the branch and Herdr guards, the cron line, and the poller's claims with a stub `pig`.

## What differs from ao-opencode

OpenCode exposes a server that sees every session, so its plugin tracks working/idle across windows, creates child sessions, and wakes any of them. Pi/PiG runs one session per process, so this port:

- reports `working`/`idle` from this session's own `agent_start` / `agent_settled` events;
- wakes the model with `sendUserMessage` instead of prompting a remote session;
- dispatches tasks to headless `pig --print` worker processes in worktrees (a single worker per task, rather than a supervisor-plus-subagent pair);
- registers the `office` MCP server with a self-refreshing `!command` bearer, since there is no server-wide MCP transform to replay a refreshed token.

The harness-independent half (`lib/office.ts`: OAuth, credentials, the MCP client, prompts, worktrees, claim and lock files) is shared by the extension, the bearer helper and the poller, exactly as in ao-opencode.
