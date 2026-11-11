// Runs only from test/tmp-guard.mts: prints a temp dir and exits with no hook,
// so the guard can see that the process-exit sweep removed it and that the
// guard's own temp directory is the one the child used.
import { tempDir } from "../suite.mts"

console.log(tempDir("ao-fixture-print-"))
