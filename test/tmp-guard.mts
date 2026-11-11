// The guard for the temp-dir rule: a run of the suite leaves its temp directory
// as it found it, and a run that fails still removes what it made.
//
// Not named `*.test.mts` on purpose. It runs the suite itself, so the suite's
// glob must not pick it up: two copies of the fixed callback ports some files
// use (45321, 45521) would then be bound by the same run. CI runs it as its own
// step after `npm test`; run it here as `npm run test:tmp`.
//
// Each measured run gets a fresh, empty directory of its own as TMPDIR, TMP and
// TEMP, and the assertion after it is that the directory is still empty. The
// private namespace is what makes the count mean something: the shared temp
// directory holds other runs' entries (and a person's), so counting it races
// them, and matching a name prefix misses a leak that names itself something
// else. Here `os.tmpdir()` resolves inside this directory for the run and every
// process it starts, so anything left in it was left by that run.
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { dirname, join } from "node:path"
import { test } from "node:test"
import { fileURLToPath } from "node:url"

const testDir = dirname(fileURLToPath(import.meta.url))
const repo = dirname(testDir)
const nodeArgs = ["--experimental-strip-types", "--disable-warning=ExperimentalWarning"]
// Every file the suite's glob runs; this guard is not one of them, since its
// name has no `.test.` in it.
const suiteFiles = readdirSync(testDir)
  .filter((file) => file.endsWith(".test.mts"))
  .sort()
  .map((file) => join(testDir, file))

// The guard's own directory in the system temp directory, removed when this
// process ends. Only this exact path is ever removed - never an entry matched by
// a name pattern in the shared directory, which could be another run's.
const root = mkdtempSync(join(tmpdir(), `ao-opencode-guard-${process.pid}-`))
process.once("exit", () => rmSync(root, { recursive: true, force: true }))

let measured = 0

// A fresh temp directory for one measured run, empty so that what is in it
// afterwards belongs to that run alone.
const runDir = (): string => {
  const dir = join(root, `run-${++measured}`)
  mkdirSync(dir)
  return dir
}

const left = (dir: string) => readdirSync(dir).sort()

const child = (args: string[], dir: string) => {
  // A run's whole temp namespace is `dir`: the suite's own root (test/suite.mts)
  // and every fixture, socket and git worktree it makes resolve there.
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    TMPDIR: dir,
    TMP: dir,
    TEMP: dir,
    // The compile cache would put a directory of its own here, which is not the
    // suite's entry to account for.
    NODE_DISABLE_COMPILE_CACHE: "1",
  }
  // The runner hands a test file's process NODE_TEST_CONTEXT; a nested
  // `node --test` that inherits it treats itself as inside a run and skips every
  // file it is given.
  delete env.NODE_TEST_CONTEXT
  delete env.NODE_COMPILE_CACHE
  return spawnSync(process.execPath, [...nodeArgs, ...args], {
    cwd: repo,
    encoding: "utf8",
    // Long enough for the suite on a slow runner, bounded so a hang fails here
    // instead of holding CI.
    timeout: 10 * 60_000,
    env,
  })
}

const report = (run: ReturnType<typeof child>) =>
  `${run.error ? `error: ${run.error.message}\n` : ""}exit ${run.status}\n${run.stdout ?? ""}\n${run.stderr ?? ""}`

test("a whole run leaves its temp directory empty", () => {
  assert.ok(suiteFiles.length > 0, "no test files found to run")
  const dir = runDir()
  assert.deepEqual(left(dir), [], "the measured run starts with an empty temp directory")
  const run = child(["--test", ...suiteFiles], dir)
  // The child's verdict belongs to `npm test`; what this guard needs is that the
  // files really ran - a nested runner that inherits NODE_TEST_CONTEXT reports
  // success over zero tests - and left nothing behind. Both reporters count
  // tests: TAP says `# tests N`, the spec reporter `ℹ tests N`.
  assert.match(run.stdout ?? "", /(?:ℹ tests|# tests) [1-9]\d*/, `the child ran no tests:\n${report(run)}`)
  assert.deepEqual(left(dir), [], `the run left ${left(dir).join(", ")} in its temp directory (child exit ${run.status})`)
})

test("a run that fails still removes the temp dirs it made", () => {
  const dir = runDir()
  const run = child(["--test", join(testDir, "fixtures", "failing.test.mts")], dir)
  assert.equal(run.status, 1, `the fixture was meant to fail:\n${report(run)}`)
  assert.match(run.stdout ?? "", /still leaves its temp dir to the after hook/, "the fixture failed for another reason")
  assert.deepEqual(left(dir), [], `the failed run left ${left(dir).join(", ")} in its temp directory`)
})

test("a child that exits without a hook leaves the temp directory empty", () => {
  // A process that makes a temp dir and ends without a hook: this is the sweep
  // the rule relies on when a file is killed before its after hook runs.
  const dir = runDir()
  const run = child([join(testDir, "fixtures", "print-temp.mts")], dir)
  assert.equal(run.status, 0, report(run))
  const path = (run.stdout ?? "").trim()
  // The path proves the run resolved `os.tmpdir()` inside this directory; an
  // empty count alone would also pass a child that wrote somewhere else.
  assert.ok(path.startsWith(dir), `the child used ${path}, not ${dir}`)
  assert.equal(existsSync(path), false, "the exit sweep left the child's temp dir behind")
  assert.deepEqual(left(dir), [], `the child's temp dir is still in its temp directory`)
})
