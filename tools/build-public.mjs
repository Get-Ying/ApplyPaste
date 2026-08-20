import { cp, mkdir, rm } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const releaseRoot = resolve(root, "release");
const target = resolve(releaseRoot, "ApplyPaste-v0.1-beta");
if (!target.startsWith(`${releaseRoot}\\`)) throw new Error("release target escaped release directory");

await rm(target, { recursive: true, force: true });
await mkdir(target, { recursive: true });
for (const entry of ["manifest.json", "popup.html", "sidepanel.html", "README.md", "src", "icons", "data", "public-template", "docs"]) {
  await cp(resolve(root, entry), resolve(target, entry), { recursive: true });
}
console.log(target);
