// The guard for the temp-dir rule: one run of the suite leaves no new top-level
// /tmp entry of its own behind, and a run that fails still removes what it made.
//
// Not named `*.test.mts` on purpose. It runs the suite itself, so the suite's
// glob must not pick it up: two copies of the fixed callback ports some files
// use (45321, 45521) would then be bound by the same run. CI runs it as its own
// step after `npm test`; run it here as `npm run test:tmp`.
//
// It gives the child run a temp prefix of its own and counts /tmp entries
// carrying that prefix before and after, so the assertion is about the entries
// that run made - never about a directory another run (or a person) left there.
import assert from "node:assert/strict"
import { spawnSync } from "node:child_process"
import { existsSync, readdirSync } from "node:fs"
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

// The top-level entries of /tmp whose name starts with this run's prefix.
const owned = (prefix: string) => readdirSync(tmpdir()).filter((name) => name.startsWith(prefix)).sort()

const child = (args: string[], prefix: string) => {
  const env: NodeJS.ProcessEnv = { ...process.env, AO_OPENCODE_TEST_TMP_PREFIX: prefix }
  // The runner hands a test file's process NODE_TEST_CONTEXT; a nested
  // `node --test` that inherits it treats itself as inside a run and skips every
  // file it is given.
  delete env.NODE_TEST_CONTEXT
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

test("a whole run leaves no new top-level /tmp entry it created", () => {
  const prefix = `ao-opencode-test-guard-${process.pid}-`
  assert.ok(suiteFiles.length > 0, "no test files found to run")
  const before = owned(prefix)
  const run = child(["--test", ...suiteFiles], prefix)
  const after = owned(prefix)
  // The child's verdict belongs to `npm test`; what this guard needs is that the
  // files really ran - a nested runner that inherits NODE_TEST_CONTEXT reports
  // success over zero tests - and left nothing behind. Both reporters count
  // tests: TAP says `# tests N`, the spec reporter `ℹ tests N`.
  assert.match(run.stdout ?? "", /(?:ℹ tests|# tests) [1-9]\d*/, `the child ran no tests:\n${report(run)}`)
  assert.deepEqual(after, before, `the run left ${after.filter((name) => !before.includes(name)).join(", ")} in /tmp (child exit ${run.status})`)
})

test("a run that fails still removes the temp dirs it made", () => {
  const prefix = `ao-opencode-test-guard-fail-${process.pid}-`
  const before = owned(prefix)
  const run = child(["--test", join(testDir, "fixtures", "failing.test.mts")], prefix)
  const after = owned(prefix)
  assert.equal(run.status, 1, `the fixture was meant to fail:\n${report(run)}`)
  assert.match(run.stdout ?? "", /still leaves its temp dir to the after hook/, "the fixture failed for another reason")
  assert.deepEqual(after, before, `the failed run left ${after.filter((name) => !before.includes(name)).join(", ")} in /tmp`)
})

test("the prefix the guard counts is the one the child uses, and exit alone removes the dir", () => {
  // A process that makes a temp dir and ends without a hook: this is the sweep
  // the rule relies on when a file is killed before its after hook runs.
  const prefix = `ao-opencode-test-guard-exit-${process.pid}-`
  const run = child([join(testDir, "fixtures", "print-temp.mts")], prefix)
  assert.equal(run.status, 0, report(run))
  const path = (run.stdout ?? "").trim()
  assert.ok(path.startsWith(join(tmpdir(), prefix)), `the child used ${path}, not the guard's prefix`)
  assert.equal(existsSync(path), false, "the exit sweep left the child's temp dir behind")
  assert.deepEqual(owned(prefix), [], "the child's temp dir is still in /tmp")
})
