import { access, readFile, readdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const release = resolve(root, "release", "ApplyPaste-v0.1-beta");
const required = ["manifest.json", "popup.html", "sidepanel.html", "README.md", "src/app.js", "src/internal-db.js", "icons/job-paste-128.png", "public-template/data/profile.json", "docs/PUBLIC_BETA_AUDIT.md"];
const digest = value => createHash("sha256").update(value).digest("hex");

for (const file of required) await access(resolve(release, file));
for (const file of ["manifest.json", "popup.html", "sidepanel.html", "src/app.js", "src/resume-docx.js", "src/internal-db.js"]) {
  if (digest(await readFile(resolve(root, file))) !== digest(await readFile(resolve(release, file)))) throw new Error(`stale release file: ${file}`);
}
const entries = new Set(await readdir(release));
for (const forbidden of ["attachments", "records", "exports", "private", "tools", "tests"]) if (entries.has(forbidden)) throw new Error(`forbidden release directory: ${forbidden}`);
const profile = JSON.parse(await readFile(resolve(release, "public-template/data/profile.json"), "utf8"));
if (profile.items.length) throw new Error("release public template is not empty");
await access(resolve(root, "release", "ApplyPaste-v0.1-beta.zip"));
console.log("ApplyPaste public build: PASS");
