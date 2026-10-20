// Agent Office poller for cron: one checkin, then exit. It shares lib/office.ts
// with the plugin, so it baselines, claims and prompts the same way, and a
// running OpenCode and cron never both handle one batch.
//
//   node poll.ts [--dir <workspace>] [--dry-run]     (or bun poll.ts)
//
//   new messages     one headless `opencode run` with the plugin's inbox prompt
//   new todo tasks   claimed, given a worktree, and run headless as a supervisor
//                    with its worker, up to two tasks at a time per agent
//   --dry-run        prints ids and counts; no Office writes, nothing started
//
//   AGENT_OFFICE_OPENCODE_BIN   the opencode executable, default `opencode`
//   plus the plugin's AGENT_OFFICE_* settings (URL, TOKEN, CREDENTIALS, AUTODISPATCH)
//
// The first run is a baseline for tasks, like the plugin's first poll; unread
// messages are not baselined. Tokens are never logged.

import { spawn, spawnSync } from "node:child_process"
import { mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs"
import { basename, dirname, join, resolve } from "node:path"
import { parseArgs } from "node:util"
import {
  acquireDispatch,
  addWorktree,
  agentStateDir,
  claimNotice,
  claimTask,
  createBearer,
  createClient,
  DISPATCH_STALE_MS,
  fromOthers,
  inboxPrompt,
  itemKey,
  releaseDispatch,
  releaseNotice,
  resolveCredentials,
  supervisorPrompt,
  taskID,
  workspaceKey,
} from "./lib/office.ts"

const OPENCODE = process.env.AGENT_OFFICE_OPENCODE_BIN || "opencode"
const INBOX_TIMEOUT_MS = 60 * 60_000

type Seen = { messages: string[]; assigned: string[]; unclaimed: string[] }

const log = (line: string) => process.stderr.write(`[agent-office poll] ${line}\n`)
const list = (value: unknown): any[] => (Array.isArray(value) ? value : [])

// The repo root a worktree belongs to, as OpenCode's project canonical path.
function canonicalRoot(directory: string) {
  const git = spawnSync("git", ["-C", directory, "rev-parse", "--path-format=absolute", "--git-common-dir"], { encoding: "utf8" })
  const common = git.status === 0 ? git.stdout.trim() : ""
  return basename(common) === ".git" ? dirname(common) : directory
}

function loadSeen(path: string): Seen | undefined {
  try {
    const value = JSON.parse(readFileSync(path, "utf8"))
    return { messages: list(value.messages), assigned: list(value.assigned), unclaimed: list(value.unclaimed) }
  } catch {
    return undefined
  }
}

function saveSeen(path: string, seen: Seen) {
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 })
  writeFileSync(`${path}.tmp`, JSON.stringify(seen, null, 2) + "\n", { mode: 0o600 })
  renameSync(`${path}.tmp`, path)
}

// Overlapping cron runs exit at once; a lock left by a dead poller is taken over.
function acquireRunLock(path: string): (() => void) | undefined {
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 })
  for (let i = 0; i < 2; i++) {
    try {
      writeFileSync(path, JSON.stringify({ pid: process.pid, created: Date.now() }) + "\n", { flag: "wx", mode: 0o600 })
      return () => {
        try {
          unlinkSync(path)
        } catch {}
      }
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw err
      try {
        process.kill(JSON.parse(readFileSync(path, "utf8")).pid, 0)
        return undefined
      } catch (err) {
        if ((err as NodeJS.ErrnoException).code === "EPERM") return undefined
      }
      try {
        unlinkSync(path)
      } catch {}
    }
  }
  return undefined
}

// `opencode run` is OpenCode's non-interactive mode. The prompt is one argv
// entry, never a shell string.
function runOpencode(directory: string, title: string, prompt: string, timeoutMs: number): Promise<number | null> {
  return new Promise((done, fail) => {
    const child = spawn(OPENCODE, ["run", "--dir", directory, "--title", title, prompt], { cwd: directory, stdio: ["ignore", "ignore", "inherit"] })
    const timer = setTimeout(() => child.kill("SIGTERM"), timeoutMs)
    child.on("error", (err) => {
      clearTimeout(timer)
      fail(err)
    })
    child.on("close", (code) => {
      clearTimeout(timer)
      done(code)
    })
  })
}

async function main(): Promise<number> {
  const { values } = parseArgs({ options: { dir: { type: "string" }, "dry-run": { type: "boolean" } } })
  const directory = resolve(values.dir ?? process.cwd())
  const dryRun = Boolean(values["dry-run"])
  const autoDispatch = !["off", "false"].includes((process.env.AGENT_OFFICE_AUTODISPATCH ?? "on").toLowerCase())

  const resolved = resolveCredentials(directory, canonicalRoot(directory))
  const { bearer } = createBearer(() => resolved.credentialsFile, resolved.credentials)
  if (!(await bearer())) {
    log(`not connected for ${directory}: run /office in OpenCode first, or set AGENT_OFFICE_TOKEN`)
    return 2
  }
  const { call, checkin } = createClient(bearer)
  const me = await checkin()
  const agent = me.agent?.slug ?? workspaceKey(directory)
  const pollDir = agentStateDir(resolved.legacyCredentialsFile, "poll", agent)
  const noticeDir = agentStateDir(resolved.legacyCredentialsFile, "notices", agent)
  const dispatchDir = agentStateDir(resolved.legacyCredentialsFile, "dispatch", agent)

  const release = dryRun ? () => {} : acquireRunLock(join(pollDir, "run.lock"))
  if (!release) {
    log("another poll is still running; exiting")
    return 0
  }
  const runs: Promise<void>[] = []
  try {
    const statePath = join(pollDir, "seen.json")
    const saved = loadSeen(statePath)
    const messages = list(me.pending?.unread_messages).filter((item) => fromOthers(item, me.agent))
    const assigned = list(me.pending?.open_tasks)
    const unclaimed = list(me.pending?.unclaimed_team_tasks)
    const seenMessages = new Set(saved?.messages)
    const freshMessages = messages.filter((item) => !seenMessages.has(itemKey(item, "message")))
    const freshAssigned = saved ? assigned.filter((item) => !saved.assigned.includes(itemKey(item, "task"))) : []
    const freshUnclaimed = saved ? unclaimed.filter((item) => !saved.unclaimed.includes(itemKey(item, "task"))) : []
    log(`@${agent}: ${freshMessages.length} new message(s), ${freshAssigned.length} new assigned task(s), ${freshUnclaimed.length} new unclaimed task(s)${saved ? "" : " (task baseline)"}`)

    // A task the cap holds back stays unseen, so the next run tries it again.
    const retry = new Set<string>()
    const todo = freshAssigned.filter((item) => String(item?.status ?? "todo").toLowerCase() === "todo")
    if (todo.length > 0 && !autoDispatch) log(`auto-dispatch is off; not dispatching ${todo.length} task(s)`)
    for (const item of autoDispatch ? todo : []) {
      const id = taskID(item)
      if (dryRun) {
        log(`would dispatch task ${id}`)
        continue
      }
      const acquired = retry.size > 0 ? "full" : acquireDispatch(dispatchDir, id)
      if (acquired === "full") {
        retry.add(itemKey(item, "task"))
        continue
      }
      if (acquired === "busy") {
        log(`task ${id} is already dispatched`)
        continue
      }
      runs.push(
        (async () => {
          try {
            const detail = await call("get_task", { task_id: id }).catch(() => undefined)
            await claimTask(call, id)
            const { path, branch } = await addWorktree(canonicalRoot(directory), id)
            log(`dispatching task ${id} to ${path}`)
            log(`task ${id} run exited ${await runOpencode(path, `Office task ${id}`, supervisorPrompt(id, detail?.task ?? detail, path, branch), DISPATCH_STALE_MS)}`)
          } catch (err) {
            log(`could not dispatch task ${id}: ${String(err)}`)
          } finally {
            releaseDispatch(dispatchDir, id)
          }
        })(),
      )
    }
    if (retry.size > 0) log(`${retry.size} task(s) wait for a free dispatch slot`)

    const handled = new Set<string>()
    if (freshMessages.length > 0) {
      const ids = freshMessages.map((item) => itemKey(item, "message"))
      if (dryRun) {
        log(`would handle message(s) ${ids.join(", ")}`)
      } else if (!claimNotice(noticeDir, ids)) {
        log("another OpenCode is handling these messages")
        for (const id of ids) handled.add(id)
      } else {
        log(`handling ${ids.length} message(s)`)
        const code = await runOpencode(directory, "Office inbox", inboxPrompt(freshMessages), INBOX_TIMEOUT_MS).catch((err) => {
          log(`could not start ${OPENCODE}: ${String(err)}`)
          return null
        })
        if (code === 0) {
          for (const id of ids) handled.add(id)
        } else {
          releaseNotice(noticeDir, ids)
          log(`inbox run failed (exit ${code}); the next run retries`)
        }
      }
    }

    if (!dryRun) {
      saveSeen(statePath, {
        messages: messages.map((item) => itemKey(item, "message")).filter((id) => seenMessages.has(id) || handled.has(id)),
        assigned: assigned.map((item) => itemKey(item, "task")).filter((id) => !retry.has(id)),
        unclaimed: unclaimed.map((item) => itemKey(item, "task")),
      })
    }
  } finally {
    release()
  }
  await Promise.all(runs)
  return 0
}

process.exit(
  await main().catch((err) => {
    log(`poll failed: ${String(err)}`)
    return 1
  }),
)
