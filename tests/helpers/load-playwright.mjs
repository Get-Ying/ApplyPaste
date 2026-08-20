import { homedir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

let module;
try {
  module = await import("playwright");
} catch {
  const bundled = join(homedir(), ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "node", "node_modules", "playwright", "index.js");
  module = await import(pathToFileURL(bundled).href);
}

export default module.default || module;
