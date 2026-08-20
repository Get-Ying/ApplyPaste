import http from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import playwright from "./helpers/load-playwright.mjs";

const { chromium } = playwright;
const root = fileURLToPath(new URL("../", import.meta.url));
const mime = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png" };

const manifest = JSON.parse(await readFile(join(root, "manifest.json"), "utf8"));
if (manifest.name !== "ApplyPaste") throw new Error("manifest name must be ApplyPaste");
if (manifest.action?.default_title !== "ApplyPaste") throw new Error("action title must be ApplyPaste");
if ("key" in manifest) throw new Error("manifest key must not pin ApplyPaste to the original extension identity");

const appSource = await readFile(join(root, "src", "app.js"), "utf8");
const internalDbSource = await readFile(join(root, "src", "internal-db.js"), "utf8");
for (const forbidden of ["loadBundledProfile", "data/jobpaste-profile.json", "data/attachments.json", "data/settings.json"]) {
  if (appSource.includes(forbidden)) throw new Error(`app startup must not reference ${forbidden}`);
}
for (const required of ["applyPasteUiSettings", "applyPasteFeishuConfig"]) {
  if (!appSource.includes(required)) throw new Error(`missing ApplyPaste namespace ${required}`);
}
if (!internalDbSource.includes("ApplyPasteInternal")) throw new Error("missing ApplyPasteInternal database namespace");

const server = http.createServer(async (request, response) => {
  try {
    const relative = decodeURIComponent(new URL(request.url, "http://localhost").pathname).replace(/^\/+/, "") || "popup.html";
    if (relative === "favicon.ico") { response.writeHead(204); return response.end(); }
    const target = normalize(join(root, relative));
    if (!target.startsWith(normalize(root))) throw new Error("invalid path");
    const data = await readFile(target);
    response.writeHead(200, { "content-type": mime[extname(target)] || "application/octet-stream" });
    response.end(data);
  } catch {
    response.writeHead(404);
    response.end("not found");
  }
});

await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const port = server.address().port;
const browser = await chromium.launch({ headless: true, executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" });
const page = await browser.newPage({ viewport: { width: 650, height: 600 } });
const errors = [];
page.on("pageerror", error => errors.push(error.message));
page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });

await page.addInitScript(({ port }) => {
  const nativeFetch = window.fetch.bind(window);
  const readMemory = () => JSON.parse(localStorage.getItem("applyPasteIsolationStorage") || "{}");
  const writeMemory = memory => localStorage.setItem("applyPasteIsolationStorage", JSON.stringify(memory));
  writeMemory({
    jobPasteProfile: { categories: ["legacy"], items: [{ id: "legacy", title: "legacy", fields: [] }] },
    jobPasteAttachments: [{ id: "legacy-file", title: "legacy attachment" }],
    jobPasteRecords: [{ "记录ID": "legacy-record" }],
    jobPasteSettings: { currentVersion: "legacy" }
  });
  window.__storageGets = [];
  window.__fetchedUrls = [];
  window.chrome = {
    runtime: { getURL: path => `http://127.0.0.1:${port}/${path}`, onMessage: { addListener: () => {} }, sendMessage: async () => {} },
    storage: { local: {
      get: async keys => {
        const requested = Array.isArray(keys) ? keys : [keys];
        window.__storageGets.push(...requested);
        const memory = readMemory();
        return Object.fromEntries(requested.filter(key => key in memory).map(key => [key, memory[key]]));
      },
      set: async values => writeMemory({ ...readMemory(), ...values }),
      remove: async keys => {
        const memory = readMemory();
        for (const key of (Array.isArray(keys) ? keys : [keys])) delete memory[key];
        writeMemory(memory);
      }
    } },
    downloads: { download: async () => 1 },
    tabs: { query: async () => [{ id: 1, url: "https://example.com/job" }], create: async () => ({}) },
    scripting: { executeScript: async () => [] }
  };
  window.fetch = async (input, options) => {
    window.__fetchedUrls.push(String(input));
    return nativeFetch(input, options);
  };
}, { port });

await page.goto(`http://127.0.0.1:${port}/popup.html`, { waitUntil: "networkidle" });
await page.waitForFunction(() => JSON.parse(localStorage.getItem("applyPasteIsolationStorage")).applyPasteUiSettings);

if (errors.length) throw new Error(errors.join("\n"));
if (await page.locator("h1").textContent() !== "ApplyPaste") throw new Error("visible product name must be ApplyPaste");
if (await page.locator(".card").count() !== 0) throw new Error("first install must not render bundled profile items");
if (await page.locator(".attachmentRow").count() !== 0) throw new Error("first install must not render bundled attachments");

const result = await page.evaluate(async () => {
  const memory = JSON.parse(localStorage.getItem("applyPasteIsolationStorage"));
  const databases = await indexedDB.databases?.() || [];
  return { memory, gets: window.__storageGets, fetched: window.__fetchedUrls, databases: databases.map(db => db.name) };
});

for (const key of ["jobPasteProfile", "jobPasteAttachments", "jobPasteRecords", "jobPasteSettings"]) {
  if (!(key in result.memory)) throw new Error(`${key} was deleted or migrated`);
  if (result.gets.includes(key)) throw new Error(`${key} was read during ApplyPaste startup`);
}
if (!("applyPasteUiSettings" in result.memory)) throw new Error("ApplyPaste settings were not written under the new namespace");
if (result.fetched.some(url => /data\/(jobpaste-profile|attachments|settings)\.json/.test(url))) {
  throw new Error("startup fetched bundled Job Paste data");
}
if (!result.databases.includes("ApplyPasteInternal")) throw new Error("ApplyPasteInternal database was not created");
if (result.databases.includes("JobPasteInternal")) throw new Error("JobPasteInternal database was touched");

await browser.close();
server.close();
console.log("ApplyPaste isolation verified");
