import assert from "node:assert/strict"
import { mkdtempSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"
import { acquireDispatch, holdDispatch } from "../lib/office.ts"

// A fake hc whose terminal is already finished: hc get answers state finished
// whatever the name, so a held lock is stale as soon as its worker is gone.
function finishedHc() {
  const dir = mkdtempSync(join(tmpdir(), "ao-hc-"))
  const bin = join(dir, "hc")
  writeFileSync(bin, "#!/bin/sh\necho '{\"state\":\"finished\"}'\nexit 0\n", { mode: 0o755 })
  return bin
}

test("a finished terminal frees the slot even though the holding pid is alive", () => {
  process.env.AGENT_OFFICE_HC_BIN = finishedHc()
  try {
    const dir = mkdtempSync(join(tmpdir(), "ao-locks-"))
    assert.equal(acquireDispatch(dir, "task-a"), "ok")
    // The lock names the hc terminal; the pid would be this long-lived test
    // process, and must not keep the slot alive.
    holdDispatch(dir, "task-a", { terminal: "office-a" })
    assert.equal(acquireDispatch(dir, "task-b"), "ok", "the finished worker's slot is free")
  } finally {
    delete process.env.AGENT_OFFICE_HC_BIN
  }
})

test("a working terminal keeps the slot, and a launch with no terminal yet keeps its pid slot", () => {
  const dir = mkdtempSync(join(tmpdir(), "ao-locks-"))
  const working = join(dir, "hc")
  writeFileSync(working, "#!/bin/sh\necho '{\"state\":\"live\",\"status\":{\"state\":\"working\"}}'\nexit 0\n", { mode: 0o755 })
  process.env.AGENT_OFFICE_HC_BIN = working
  try {
    assert.equal(acquireDispatch(dir, "task-a"), "ok")
    holdDispatch(dir, "task-a", { terminal: "office-a" })
    assert.equal(acquireDispatch(dir, "task-b"), "ok")
    // Two slots now: the working terminal and the new pid lock.
    assert.equal(acquireDispatch(dir, "task-c"), "full")
  } finally {
    delete process.env.AGENT_OFFICE_HC_BIN
  }
})
