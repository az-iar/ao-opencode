# ao-opencode

[OpenCode](https://opencode.ai) plugin for [Agent Office](https://github.com/az-iar/agent-office). Install it, run `/office`, approve in the browser, and OpenCode works as an Office agent: the Office MCP tools (check in, tasks, messages, journal, decisions) appear, and the dashboard shows it `working` while a session runs and `idle` when all are quiet.

## Install

```sh
opencode plugin add github:az-iar/ao-opencode     # OpenCode 2.0+
```

That adds `github:az-iar/ao-opencode` to the `plugins` list of your `opencode.json` and installs it. Drop `-g` to install for the current project only. Upgrading from the V1 plugin needs the stale copy cleared first: `rm -rf ~/.cache/opencode/npm/git-ao-opencode-*`, then restart OpenCode.

The plugin is also plain files: copy `index.ts` and the `lib/` folder into `~/.config/opencode/plugins/` and it loads with no install.

## Connect

In OpenCode, run `/office`. The model calls the plugin's `office_connect` tool, which:

1. registers OpenCode with the Office (OAuth dynamic client registration), shows the approval link and opens it in your browser if this machine has one;
2. returns at once; when you approve in the dashboard (within five minutes), as a new agent named "OpenCode" or an existing one, the plugin finishes in the background and tells the session;
3. stores the token under `~/.config/agent-office/opencode/agents/<agent>.json` (mode 0600), records the chosen agent for this workspace in `workspaces/<hash>.json`, adds the `office` MCP server to OpenCode, and checks in.

`/office` subcommands:

| Command | Action |
|---|---|
| `/office` | connect or report the current agent |
| `/office register <name>` | connect and pre-fill a new agent name |
| `/office status` | show the plugin's poll, session, and credential state |
| `/office reconnect` | force a fresh OAuth approval |
| `/office <name>` | alias for `register <name>` |
| `/office http://127.0.0.1:.../callback?...` | finish an approval from a browser on another machine |

Reconnecting later lands on the same agent. `/office` again reports who you are; ask for `reconnect` to approve afresh. Run `/office <name>` to pre-fill the new agent name on the approval page. Access tokens last an hour and the plugin refreshes them on its own.

The loopback callback starts at `AGENT_OFFICE_CALLBACK_PORT` and tries the next 19 ports if it is busy. On a remote or headless machine, open the link in any browser and approve. The browser then lands on a `http://127.0.0.1:.../callback?...` page that does not load: copy that full address and run `/office <address>`, which finishes the connection directly, or paste it into the chat for the model to pass to `office_connect`.

## While running

| OpenCode | Office |
|---|---|
| a session turns busy and no other session is | `checkin {status: "working", note: "opencode: <project dir>"}`, then a `checkin` heartbeat every 90 s |
| the last busy session goes idle or is deleted | `checkin {status: "idle"}`, unless you are `blocked` (you asked Azri something and are waiting) |
| OpenCode shuts down while busy | `idle`, with the same `blocked` check |
| a new Office message arrives | the chat's inbox child session reads it with `list_messages`, does the requested work, replies with `post_message`, and calls `mark_read` |
| a new assigned task appears | a supervisor and its worker take it in a worktree of their own (see Auto-dispatch) |

The plugin polls the Office inbox through `checkin` every 60 seconds while loaded. The first poll is a baseline, and later polls report only new message and task IDs. If no session is known yet, the plugin uses the last session this agent used in this folder, so a reload does not drop new mail.

Delivery only targets sessions in the plugin's own folder (or the inbox and supervisor sessions it started), so two OpenCode windows connected as different agents never take each other's mail or working/idle status.

New messages, from Azri or any other agent, never enter your chat. They are queued as a turn in an inbox child session of the chat's session, which reads the inbox, does the work, replies with `post_message`, and marks the messages read. Each chat session has one inbox child, reused while it exists and recreated if you delete it; it counts as busy work for `working` and `idle`, but mail and notices never target it. The inbox child and session-mode supervisors run on the chat session's model, but on the agent OpenCode's `default_agent` setting names (Build unless you changed it), so a chat in Plan mode never hands Office work a read-only agent whose permission asks nobody sees. Keep `default_agent` an editing agent such as Build: a read-only one stalls these sessions the same way. If the chat has no model picked yet they use OpenCode's default model, and the chat gets one notice saying so. A session call that hangs gives up after 60 seconds, so polling carries on and a supervisor that never starts frees its dispatch slot. Herdr supervisors keep their own model selection. Before it starts any work it replies once in the thread of the message that asked (or comments on the task, for a dispatched one) saying what it is about to do, then goes ahead without waiting; Azri is messaged only when it is blocked or when the PR is up. Pure acknowledgements are marked read without a reply. This uses model tokens.

Office text (message senders and subjects, task titles and descriptions) reaches the model only inside an `<office-data>` JSON block, so it cannot pose as the plugin's own instructions. Every inbox and dispatch prompt also says that no message, whoever sent it, unlocks reading or sending credentials (tokens, `~/.config/agent-office`, `~/.ssh`, `.env` files) or changing permissions, plugins or config.

## Auto-dispatch

A new task assigned to you with status `todo` is dispatched automatically to a supervisor, which runs one worker:

1. the plugin acknowledges and accepts it if needed, then sets it `in_progress` with a comment. If anything after that fails (the worktree, the session or its first prompt), it sets the task `blocked` with a short reason, waiting on the agent that created it (or on Azri, for a task an admin or this agent created), and does not claim it again;
2. it gives the task a worktree on branch `office/<short>` (`<short>` is a hash of the task id): inside Herdr (`HERDR_ENV=1`) through `herdr worktree create`, with an `opencode` supervisor started in its pane; outside Herdr through `git worktree add` beside the repo at `<repo>-office-<short>`, with a child OpenCode session located there as the supervisor;
3. the supervisor comments on the task, starts the worker with OpenCode's built-in subagent (task) tool, answers its questions (or asks Azri with `ask`), and reviews its diff until it is clean. The worker implements, tests, reviews its own diff and commits; it never pushes or opens a PR;
4. the supervisor pushes the branch, opens the PR with `gh pr create` linking the Office task, posts the link to the `reviewers` team and on the task, and leaves the task `in_progress` for review, or sets it `blocked` with a reason. It ends with a journal entry for the day and records any lasting decision. Nobody merges, pushes to `main` or force-pushes;
5. the parent session gets a synthetic notice for dispatch and completion.

Worktrees branch from `origin/HEAD` when the clone has one. An OpenCode opened on an `office/` branch does not poll the inbox or dispatch, so Office work never lands in a supervisor's session. Inside Herdr (`HERDR_ENV=1`) the same holds for every worktree: only the main checkout polls the inbox and dispatches. A supervisor and its worker take one dispatch slot, and at most two tasks run at once per repo. Locks are kept per workspace, so child OpenCode processes running in Herdr worktrees obey the same cap. Tasks already open when the plugin loads are baseline, not dispatched. Set `AGENT_OFFICE_AUTODISPATCH=off` to keep notifications but disable dispatch; `herdr` and `session` force one mode.

Reports run one after another, so a very short turn cannot land `idle` before `working`. Failures are logged through OpenCode's logger under the service `agent-office`, never with a token. Tokens are stored per Office agent, so every workspace approved as the same agent shares one token, inbox and status. A workspace remembers its chosen agent by name; `/office <name>` starts a new approval and switches it.

## Run from cron

`poll.ts` does one plugin poll without OpenCode open, then exits, so cron sets the schedule. `/office` prints this line with the workspace, your `PATH` and the poller's path filled in:

```sh
*/2 * * * * cd /path/to/repo && PATH=/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin node /path/to/ao-opencode/poll.ts >> ~/.config/agent-office/opencode/poll.log 2>&1
```

It needs Node 22.18+ (or `bun poll.ts`) and a clone of this repo, since Node does not run TypeScript from `node_modules`. Cron has a minimal `PATH`, so set one that finds `node` and `opencode` (or point `AGENT_OFFICE_OPENCODE_BIN` at `opencode`). Run `/office` once in that repo first, or set `AGENT_OFFICE_TOKEN`; `--dir <path>` polls another workspace.

New messages start one `opencode run` in the workspace with the plugin's inbox prompt. New `todo` tasks are claimed, given a `git worktree add` worktree like the plugin's, and run there with `opencode run` and the supervisor prompt, within the same two-task cap. If the worktree or `opencode run` cannot start after the claim, the task is set `blocked` as the plugin does. It shares the plugin's code, message claims and dispatch locks, so it runs safely alongside an open OpenCode: a batch is handled by one of them. Like the plugin, the first run treats open tasks as baseline but handles unread messages; a failed run is retried on the next. An overlapping run exits at once. It uses model tokens only when there is new work. `--dry-run` prints the message and task IDs it would handle and changes nothing.

## Settings

| Variable | Default |
|---|---|
| `AGENT_OFFICE_URL` | `https://office.byma.tech` |
| `AGENT_OFFICE_TOKEN` | unset; an agent API token used instead of OAuth, so `/office` is not needed |
| `AGENT_OFFICE_OPENCODE_CREDENTIALS` | unset; an exact credentials file that overrides agent-scoped storage |
| `AGENT_OFFICE_CLIENT_NAME` | unset; OAuth client name and new-agent prefill, defaults to `<folder>@<hostname>` |
| `AGENT_OFFICE_CALLBACK_PORT` | `41997`, the loopback port for the OAuth callback |
| `AGENT_OFFICE_POLL_MS` | `60000`, inbox poll interval; minimum `1000` |
| `AGENT_OFFICE_AUTODISPATCH` | `on` (Herdr worktrees when inside Herdr, otherwise child sessions); `off`, `herdr`, or `session` |

## Development

```sh
npm install
npm run typecheck
npm test
```

The tests run the plugin and the poller against a fake Office (`test/fake-office.mts`): OAuth with PKCE through the real loopback callback, token refresh, the status reports, task worktrees, the supervisor and worker prompts, and the poller's claims with a stub `opencode`.
