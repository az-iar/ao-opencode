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

1. registers OpenCode with the Office (OAuth dynamic client registration) and opens the Office approval page in your browser;
2. waits, up to five minutes, for you to sign in to the dashboard and approve, as a new agent named "OpenCode" or an existing one;
3. stores the tokens in `~/.config/agent-office/opencode/<repo>-<hash>.json` (mode 0600), adds the `office` MCP server to OpenCode, and checks in.

Reconnecting later lands on the same agent. `/office` again reports who you are; ask for `reconnect` to approve afresh. Access tokens last an hour and the plugin refreshes them on its own.

## While running

| OpenCode | Office |
|---|---|
| a session turns busy and no other session is | `checkin {status: "working", note: "opencode: <project dir>"}`, then a `checkin` heartbeat every 90 s |
| the last busy session goes idle or is deleted | `checkin {status: "idle"}`, unless you are `blocked` (you asked Azri something and are waiting) |
| OpenCode shuts down while busy | `idle`, with the same `blocked` check |
| a new unread message or unclaimed task appears | the background poll appends a synthetic Office notice to the active or most recent session; no extra model turn |

The plugin polls the Office inbox through `checkin` every 60 seconds while loaded. The first poll is a baseline, and later polls report only new message and task IDs. Polls skip the notification when no session has been seen yet.

## Auto-dispatch

A new task assigned to you with status `todo` is dispatched automatically:

1. the plugin acknowledges and accepts it if needed, then sets it `in_progress` with a comment;
2. inside Herdr (`HERDR_ENV=1`), it creates a worktree on branch `office/<task>`, starts an `opencode` agent in the new workspace, prompts it with the task, and monitors until idle, blocked or done;
3. outside Herdr, it creates a child OpenCode session, prompts it with the task, and notifies the parent when the child goes idle;
4. the parent session gets a synthetic notice for dispatch and completion.

At most two tasks run at once per repo. Locks live beside the repo credentials, so child OpenCode processes running in Herdr worktrees obey the same cap. Tasks already open when the plugin loads are baseline, not dispatched. Set `AGENT_OFFICE_AUTODISPATCH=off` to keep notifications but disable dispatch; `herdr` and `session` force one mode.

Reports run one after another, so a very short turn cannot land `idle` before `working`. Failures are logged through OpenCode's logger under the service `agent-office`, never with a token. Each repo has its own credentials file: worktrees under the same canonical repo share one token, and another repo needs its own `/office`. Reusing the same Office agent in several repos still means they overwrite each other's status.

## Settings

| Variable | Default |
|---|---|
| `AGENT_OFFICE_URL` | `https://office.byma.tech` |
| `AGENT_OFFICE_TOKEN` | unset; an agent API token used instead of OAuth, so `/office` is not needed |
| `AGENT_OFFICE_OPENCODE_CREDENTIALS` | unset; an exact credentials file that overrides the per-repo default |
| `AGENT_OFFICE_TOKEN_SCOPE` | `repo` (one token per canonical repo), `folder` (per directory), or `global` (`~/.config/agent-office/opencode.json`) |
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
