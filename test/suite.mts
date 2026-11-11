// What every test file in this suite starts with: a temp root this process owns
// and the browser opener switched off.
//
// /tmp is shared, so the suite never deletes by a name pattern - a pattern
// would take a directory another run, or a person, still owns. Each test
// process makes one `ao-opencode-test-<pid>-*` root of its own and every
// mkdtemp goes under it; `removeTempRoot` runs in each file's `after` hook, so a
// test that fails still cleans up, and again on process exit, so a run killed
// mid-file or a hook that never got its turn does not leave the root behind.
// test/tmp-guard.mts counts what a whole run leaves in /tmp.
//
// The opener switch is here because a suite run once launched a real browser
// (az-iar/ao-pi#10, #13): hundreds of stray processes, and the browser's own
// request for the approval URL landed in another test's Office call log. A file
// that forgets the switch is still protected; CI sets the same switch for the
// job as a backstop.
import { mkdtempSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"

// test/tmp-guard.mts runs the suite as a child with a prefix of its own, so the
// entries it counts are only the ones that run made.
export const TEMP_PREFIX = process.env.AO_OPENCODE_TEST_TMP_PREFIX ?? "ao-opencode-test-"

process.env.AGENT_OFFICE_OPEN_BROWSER = "off"

let root: string | undefined
let made = 0

// This process's root, made on the first tempDir: a file that needs no temp dir
// makes none.
export function tempRoot(): string {
  root ??= mkdtempSync(join(tmpdir(), `${TEMP_PREFIX}${process.pid}-`))
  return root
}

// A directory of this process's own, named after what it is for, so a leftover
// would say which file forgot it.
export function tempDir(name: string): string {
  return mkdtempSync(join(tempRoot(), `${name}${++made}-`))
}

export function removeTempRoot(): void {
  if (root === undefined) return
  rmSync(root, { recursive: true, force: true, maxRetries: 3 })
  root = undefined
}

process.once("exit", removeTempRoot)
