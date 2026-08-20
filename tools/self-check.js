import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { extname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const results = [];

function report(level, message, file = "") {
  results.push({ level, message, file });
}

function readJson(path) {
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    report("FAIL", `JSON parse failed: ${error.message}`, rel(path));
    return null;
  }
}

function rel(path) {
  return relative(root, path).replaceAll("\\", "/");
}

function walk(dir, output = []) {
  for (const name of readdirSync(dir)) {
    if (["node_modules", ".git", "release", "archive"].includes(name)) continue;
    const path = join(dir, name);
    const stat = statSync(path);
    if (stat.isDirectory()) walk(path, output);
    else output.push(path);
  }
  return output;
}

const manifestPath = join(root, "manifest.json");
const packagePath = join(root, "package.json");
const popupPath = join(root, "popup.html");
const sidepanelPath = join(root, "sidepanel.html");
const requiredFiles = [manifestPath, packagePath, popupPath, sidepanelPath, join(root, "src", "app.js"), join(root, "src", "internal-db.js")];

for (const path of requiredFiles) {
  if (!existsSync(path)) report("FAIL", "Required file missing", rel(path));
}

const manifest = readJson(manifestPath);
if (manifest) {
  if (manifest.manifest_version !== 3) report("FAIL", "manifest_version is not 3", "manifest.json");
  if (manifest.name !== "ApplyPaste") report("WARN", "Manifest name is not ApplyPaste", "manifest.json");
  if ("key" in manifest) report("FAIL", "Manifest pins extension identity with key", "manifest.json");
  const iconPaths = Object.values(manifest.icons || {});
  const actionIconPaths = Object.values(manifest.action?.default_icon || {});
  for (const path of [...iconPaths, ...actionIconPaths, manifest.action?.default_popup, manifest.side_panel?.default_path].filter(Boolean)) {
    if (!existsSync(join(root, path))) report("FAIL", `Manifest resource missing: ${path}`, "manifest.json");
  }
  if (!manifest.permissions?.includes("sidePanel")) report("WARN", "Missing sidePanel permission", "manifest.json");
}

const htmlFiles = [popupPath, sidepanelPath].filter(existsSync);
for (const path of htmlFiles) {
  const text = readFileSync(path, "utf8");
  for (const match of text.matchAll(/(?:src|href)=["']([^"']+)["']/g)) {
    const target = match[1];
    if (/^(https?:|data:|#)/.test(target)) continue;
    if (!existsSync(join(root, target.split("?")[0]))) report("FAIL", `Referenced file missing: ${target}`, rel(path));
  }
  const ids = [...text.matchAll(/\sid=["']([^"']+)["']/g)].map(match => match[1]);
  const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
  for (const id of new Set(duplicates)) report("FAIL", `Duplicate DOM id: ${id}`, rel(path));
}

for (const path of walk(root)) {
  const ext = extname(path);
  if (![".js", ".mjs", ".html", ".css", ".json", ".md"].includes(ext)) continue;
  const text = readFileSync(path, "utf8");
  if (/console\.log\(/.test(text) && !rel(path).startsWith("tools/") && !rel(path).startsWith("tests/")) {
    report("WARN", "console.log in runtime file", rel(path));
  }
  if (/jobPaste(Profile|Attachments|Records|Settings)|JobPasteInternal|jobpaste-profile/.test(text) && !rel(path).startsWith("tools/") && !rel(path).startsWith("tests/")) {
    report("WARN", "Legacy Job Paste storage reference in runtime/release-adjacent file", rel(path));
  }
  if (/appSecret|App Secret|secret/i.test(text) && !rel(path).startsWith("tools/") && !rel(path).startsWith("tests/")) {
    report("WARN", "Secret-related string present, review export/storage behavior", rel(path));
  }
}

if (!results.some(item => item.level === "FAIL")) report("PASS", "No static FAIL findings");

const counts = results.reduce((acc, item) => {
  acc[item.level] = (acc[item.level] || 0) + 1;
  return acc;
}, {});

for (const item of results) {
  console.log(`${item.level}: ${item.message}${item.file ? ` (${item.file})` : ""}`);
}
console.log(`SUMMARY: PASS=${counts.PASS || 0} WARN=${counts.WARN || 0} FAIL=${counts.FAIL || 0}`);
process.exitCode = counts.FAIL ? 1 : 0;
