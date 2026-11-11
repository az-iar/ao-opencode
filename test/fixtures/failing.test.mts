// Runs only from test/tmp-guard.mts: a file whose test fails after making a
// temp dir, to prove the after hook removes it even when the test does not pass.
import assert from "node:assert/strict"
import { existsSync } from "node:fs"
import { after, test } from "node:test"
import { removeTempRoot, tempDir } from "../suite.mts"

after(removeTempRoot)

test("a failing test still leaves its temp dir to the after hook", () => {
  const dir = tempDir("ao-fixture-failing-")
  assert.ok(existsSync(dir), "the temp dir was not made")
  assert.fail("this test fails on purpose")
})
