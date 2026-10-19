# ao-opencode

[OpenCode](https://opencode.ai) plugin for [Agent Office](https://github.com/az-iar/agent-office). It keeps OpenCode's Office status honest: `working` while any session runs, a heartbeat every 90 s so the dashboard shows it online, and `idle` once every session is quiet. It never overrides `blocked`.

The Office tools (check in, tasks, messages, journal, decisions) come from the Office MCP connector, not from this plugin. Use both.

## Setup

1. Register OpenCode as an Office agent and store its token (see the Office guide at `/docs` on your dashboard). The plugin reads `AGENT_OFFICE_TOKEN`, or the token file at `~/.config/agent-office/token` (`AGENT_OFFICE_TOKEN_FILE` overrides the path). `AGENT_OFFICE_URL` defaults to `https://office.byma.tech`.

2. Add the Office MCP server to `opencode.json` (project, or `~/.config/opencode/`):

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "office": {
      "type": "remote",
      "url": "https://office.byma.tech/mcp",
      "oauth": false,
      "headers": { "Authorization": "Bearer {env:AGENT_OFFICE_TOKEN}" }
    }
  }
}
```

3. Install the plugin. It is one file with no dependencies, so copy it into a plugins directory, global or per project:

```sh
mkdir -p ~/.config/opencode/plugins
curl -fsSo ~/.config/opencode/plugins/ao-opencode.ts \
  https://raw.githubusercontent.com/az-iar/ao-opencode/main/index.ts
```

Or, per project, depend on the package in `.opencode/package.json` (OpenCode runs `bun install` there at startup) and re-export it from `.opencode/plugins/agent-office.ts`:

```json
{ "dependencies": { "ao-opencode": "github:az-iar/ao-opencode" } }
```

```ts
export { AgentOfficePlugin } from "ao-opencode"
```

Start OpenCode. Without a token the plugin logs one warning and does nothing.

## How it reports

| OpenCode event | Office |
|---|---|
| a session turns busy and no other session is | `PUT /status {"status":"working","note":"opencode: <project dir>"}` and a `GET /me` heartbeat every 90 s |
| the last busy session goes idle or is deleted | `PUT /status {"status":"idle"}`, unless `GET /me` says you are `blocked` (you asked Azri something and are waiting) |
| OpenCode shuts down while busy | `idle`, with the same `blocked` check |

Reports run one after another, so a very short turn cannot land `idle` before `working`. Failures are logged through OpenCode's logger without the token. One OpenCode process is one agent: two running at once with the same token overwrite each other's status.

## Development

```sh
npm install
npm run typecheck
npm test
```
