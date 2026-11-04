import assert from "node:assert/strict"
import { chmodSync, existsSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { test } from "node:test"

import { browserOpeningAllowed, openBrowser } from "../index.ts"

// az-iar/ao-pi#10/#13: the approval flow shells out to xdg-open, so a suite run
// launched a real browser - hundreds of stray processes on the host, and the
// browser's own request for the approval URL then landed in another test's
// Office call log (this repo's session-identity suite found that the hard way,
// in a Space that has a display). The guard is read at the boundary, and these
// tests keep it honest: a fake xdg-open on PATH records any launch, so the
// first case fails if a real one is attempted, and the second proves the
// recorder works so the first cannot pass by never being reached.
//
// They call openBrowser directly rather than running an approval: the flow
// spawns detached processes and holds a callback server, which is how a
// previous attempt at this test hung for its whole timeout.

// A PATH entry whose xdg-open records that it ran, instead of opening anything.
function fakeOpener(): { dir: string; marker: string } {
  const dir = mkdtempSync(join(tmpdir(), "ao-opener-bin-"))
  const marker = join(dir, "launched")
  writeFileSync(join(dir, "xdg-open"), `#!/bin/sh\necho "$1" >> ${marker}\n`)
  chmodSync(join(dir, "xdg-open"), 0o755)
  return { dir, marker }
}

async function withOpener<T>(opener: { dir: string }, fn: () => Promise<T>): Promise<T> {
  const saved = process.env.PATH
  process.env.PATH = `${opener.dir}:${saved}`
  try {
    return await fn()
  } finally {
    process.env.PATH = saved
  }
}

// The opener spawns detached, so give it a moment to write its marker.
const settle = () => new Promise((resolve) => setTimeout(resolve, 300))

test("a switched-off opener launches nothing, even when xdg-open is on PATH", async () => {
  const opener = fakeOpener()
  const saved = process.env.AGENT_OFFICE_OPEN_BROWSER
  process.env.AGENT_OFFICE_OPEN_BROWSER = "off"
  try {
    assert.equal(browserOpeningAllowed(), false)
    await withOpener(opener, async () => openBrowser("https://office.example/approve?x=1"))
    await settle()
  } finally {
    if (saved === undefined) delete process.env.AGENT_OFFICE_OPEN_BROWSER
    else process.env.AGENT_OFFICE_OPEN_BROWSER = saved
  }
  assert.equal(existsSync(opener.marker), false, "a real opener was attempted while switched off")
})

test("the recorder works when the opener is allowed, so the test above cannot pass by never being reached", async () => {
  const opener = fakeOpener()
  const saved = process.env.AGENT_OFFICE_OPEN_BROWSER
  delete process.env.AGENT_OFFICE_OPEN_BROWSER
  try {
    assert.equal(browserOpeningAllowed(), true)
    await withOpener(opener, async () => openBrowser("https://office.example/approve?x=2"))
    await settle()
  } finally {
    if (saved !== undefined) process.env.AGENT_OFFICE_OPEN_BROWSER = saved
  }
  assert.equal(existsSync(opener.marker), true, "the fake opener never ran, so the guard test proves nothing")
  assert.match(readFileSync(opener.marker, "utf8"), /^https:\/\/office\.example\/approve\?x=2/)
})

test("the switch understands the spellings a runner might use", () => {
  for (const off of ["off", "OFF", "0", "false", "no", " off "]) {
    process.env.AGENT_OFFICE_OPEN_BROWSER = off
    assert.equal(browserOpeningAllowed(), false, `${JSON.stringify(off)} should disable the opener`)
  }
  for (const on of ["", "on", "1", "true"]) {
    process.env.AGENT_OFFICE_OPEN_BROWSER = on
    assert.equal(browserOpeningAllowed(), true, `${JSON.stringify(on)} should leave it enabled`)
  }
  delete process.env.AGENT_OFFICE_OPEN_BROWSER
})
