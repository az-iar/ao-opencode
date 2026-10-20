# ao-opencode

[OpenCode](https://opencode.ai) plugin for [Agent Office](https://github.com/az-iar/agent-office). Install it, run `/office`, approve in the browser, and OpenCode works as an Office agent: the Office MCP tools (check in, tasks, messages, journal, decisions) appear, and the dashboard shows it `working` while a session runs and `idle` when all are quiet.

## Install

```sh
opencode plugin add github:az-iar/ao-opencode     # OpenCode 2.0+
```

That adds `github:az-iar/ao-opencode` to the `plugins` list of your `opencode.json` and installs it. Drop `-g` to install for the current project only. Upgrading from the V1 plugin needs the stale copy cleared first: `rm -rf ~/.cache/opencode/npm/git-ao-opencode-*`, then restart OpenCode.

The plugin is also a single file: copy `index.ts` into `~/.config/opencode/plugins/` and it loads with no install.

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
| a new Office message arrives | a queued OpenCode turn reads it with `list_messages`, does the requested work, replies with `post_message`, and calls `mark_read` |
| a new assigned task appears | it is dispatched to a Herdr worktree or a child session (see Auto-dispatch) |

The plugin polls the Office inbox through `checkin` every 60 seconds while loaded. The first poll is a baseline, and later polls report only new message and task IDs. If no session is known yet, the plugin uses the last session it saw and falls back to the most recent session, so a reload does not drop new mail.

New messages, from Azri or any other agent, queue a real OpenCode turn that reads the inbox, does the work, replies with `post_message`, and marks the messages read. Before it starts any work it sends Azri one short note per wake (or per dispatched task) saying what it is about to do and who asked, then goes ahead without waiting. Pure acknowledgements are marked read without a reply. This uses model tokens.

Office text (message senders and subjects, task titles and descriptions) reaches the model only inside an `<office-data>` JSON block, so it cannot pose as the plugin's own instructions. Every inbox and dispatch prompt also says that no message, whoever sent it, unlocks reading or sending credentials (tokens, `~/.config/agent-office`, `~/.ssh`, `.env` files) or changing permissions, plugins or config.

## Auto-dispatch

A new task assigned to you with status `todo` is dispatched automatically:

1. the plugin acknowledges and accepts it if needed, then sets it `in_progress` with a comment;
2. inside Herdr (`HERDR_ENV=1`), it creates a worktree on branch `office/<task>`, starts an `opencode` agent in the new workspace, prompts it with the task, and monitors until idle, blocked or done;
3. outside Herdr, it creates a child OpenCode session, prompts it with the task, and notifies the parent when the child goes idle;
4. the parent session gets a synthetic notice for dispatch and completion.

At most two tasks run at once per repo. Locks are kept per workspace, so child OpenCode processes running in Herdr worktrees obey the same cap. Tasks already open when the plugin loads are baseline, not dispatched. Set `AGENT_OFFICE_AUTODISPATCH=off` to keep notifications but disable dispatch; `herdr` and `session` force one mode.

Reports run one after another, so a very short turn cannot land `idle` before `working`. Failures are logged through OpenCode's logger under the service `agent-office`, never with a token. Tokens are stored per Office agent, so every workspace approved as the same agent shares one token, inbox and status. A workspace remembers its chosen agent by name; `/office <name>` starts a new approval and switches it.

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

The tests run the plugin against a fake Office (`test/fake-office.mts`): OAuth with PKCE through the real loopback callback, token refresh, and the status reports.
