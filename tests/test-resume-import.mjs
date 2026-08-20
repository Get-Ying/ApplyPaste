import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const result = spawnSync(process.execPath, [resolve(root, "tests", "test-word-import-smoke.mjs")], {
  cwd: root,
  encoding: "utf8",
  timeout: 90_000
});

if (result.status !== 0) {
  process.stderr.write(result.stderr || result.stdout || "Word import regression failed\n");
  process.exit(result.status || 1);
}

console.log("ApplyPaste public synthetic resume import verified");
