import { readdir, readFile, stat } from "node:fs/promises";
import { extname, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const ignoredDirectories = new Set([".git", "node_modules"]);
const binaryExtensions = new Set([".png", ".jpg", ".jpeg", ".webp", ".gif", ".ico", ".pdf", ".docx", ".zip"]);
const allowedPhones = new Set(["13800000000", "13800138000"]);
const allowedEmails = new Set(["test@example.com", "tester@example.com"]);
const forbiddenPrivateMarkers = ["应美玩", "MXene", "复旦大学", "浙江师范大学", "材料研发版", "半导体工艺版", "秋招工作台_2027"];
const findings = [];

async function walk(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  for (const entry of entries) {
    if (entry.isDirectory() && ignoredDirectories.has(entry.name)) continue;
    const path = resolve(directory, entry.name);
    if (entry.isDirectory()) await walk(path);
    else await inspect(path);
  }
}

async function inspect(path) {
  const name = relative(root, path).replaceAll("\\", "/");
  const info = await stat(path);
  if (info.size > 5_000_000 || binaryExtensions.has(extname(path).toLowerCase())) return;
  const text = await readFile(path, "utf8");
  if (name !== "tools/privacy-scan.mjs") for (const marker of forbiddenPrivateMarkers) if (text.includes(marker)) add(name, "PRIVATE_MARKER", marker);
  for (const match of text.matchAll(/(?<!\d)1[3-9]\d{9}(?!\d)/g)) if (!allowedPhones.has(match[0])) add(name, "PHONE", match[0]);
  for (const match of text.matchAll(/[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}/g)) {
    const value = match[0].toLowerCase();
    if (!allowedEmails.has(value) && !value.endsWith("@example.org") && !value.endsWith("@example.net")) add(name, "EMAIL", match[0]);
  }
  for (const match of text.matchAll(/(?<!\d)\d{6}(?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[0-9Xx](?!\d)/g)) add(name, "PRC_ID", match[0]);
  for (const match of text.matchAll(/(?:[A-Za-z]:[\\/](?:Users|My files)[\\/][^\s`'"<>]+)/g)) add(name, "PRIVATE_ABSOLUTE_PATH", match[0]);
  for (const match of text.matchAll(/\b(app[_-]?secret|access[_-]?token|refresh[_-]?token|authorization|bearer)\b\s*[=:]\s*["']([^"']{8,})["']/gi)) {
    if (!/^(?:placeholder|example|test|your[-_]|<|\{)/i.test(match[2])) add(name, "HARDCODED_SECRET", `${match[1]}=<redacted>`);
  }
  if (/^(?:attachments|records|exports|private)\//.test(name) && info.size > 0) add(name, "PRIVATE_DATA_DIRECTORY", `${info.size} bytes`);
  if (/^(?:data|public-template\/data)\/.*\.json$/.test(name)) validatePublicJson(name, text);
}

function validatePublicJson(name, text) {
  try {
    const value = JSON.parse(text);
    if (Array.isArray(value) && value.length) add(name, "NONEMPTY_PUBLIC_DATA", `${value.length} entries`);
    if (Array.isArray(value?.items) && value.items.length) add(name, "NONEMPTY_PUBLIC_PROFILE", `${value.items.length} items`);
    if (Array.isArray(value?.versions) && (value.versions.length !== 1 || value.versions[0]?.id !== "general")) add(name, "PERSONAL_VERSIONS", "expected general only");
  } catch (error) {
    add(name, "INVALID_JSON", error.message);
  }
}

function add(file, code, detail) {
  findings.push({ file, code, detail });
}

await walk(root);
if (findings.length) {
  console.error("ApplyPaste privacy scan: FAIL");
  for (const finding of findings) console.error(`- ${finding.code} ${finding.file}: ${finding.detail}`);
  process.exitCode = 1;
} else {
  console.log("ApplyPaste privacy scan: PASS");
  console.log("No private profile, attachment, record, Feishu credential, personal path, or known source marker found.");
}
