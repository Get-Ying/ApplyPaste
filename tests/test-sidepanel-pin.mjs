import http from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import playwright from "./helpers/load-playwright.mjs";

const { chromium } = playwright;
const root = fileURLToPath(new URL("../", import.meta.url));
const mime = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png" };

const server = http.createServer(async (request, response) => {
  try {
    const relative = decodeURIComponent(new URL(request.url, "http://localhost").pathname).replace(/^\/+/, "") || "popup.html";
    if (relative === "favicon.ico") {
      response.writeHead(204);
      response.end();
      return;
    }
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
const page = await browser.newPage({ viewport: { width: 760, height: 720 } });
const errors = [];
page.on("pageerror", error => errors.push(error.message));
page.on("console", message => {
  if (message.type() === "error") errors.push(message.text());
});

await page.addInitScript(({ port }) => {
  const readMemory = () => JSON.parse(localStorage.getItem("applyPasteSidePanelStorage") || "{}");
  const writeMemory = memory => localStorage.setItem("applyPasteSidePanelStorage", JSON.stringify(memory));
  writeMemory({});
  window.__sidePanelCalls = [];
  window.chrome = {
    runtime: { getURL: path => `http://127.0.0.1:${port}/${path}`, onMessage: { addListener: () => {} }, sendMessage: async () => {} },
    storage: {
      local: {
        get: async keys => {
          const requested = Array.isArray(keys) ? keys : [keys];
          const memory = readMemory();
          return Object.fromEntries(requested.filter(key => key in memory).map(key => [key, memory[key]]));
        },
        set: async values => writeMemory({ ...readMemory(), ...values }),
        remove: async keys => {
          const memory = readMemory();
          for (const key of (Array.isArray(keys) ? keys : [keys])) delete memory[key];
          writeMemory(memory);
        }
      }
    },
    downloads: { download: async () => 1 },
    tabs: { query: async () => [{ id: 7, windowId: 42, url: "https://example.com/job" }], create: async () => ({}) },
    scripting: { executeScript: async () => [] },
    sidePanel: {
      setOptions: async options => window.__sidePanelCalls.push({ method: "setOptions", options }),
      open: async options => window.__sidePanelCalls.push({ method: "open", options })
    }
  };
}, { port });

await page.goto(`http://127.0.0.1:${port}/popup.html`, { waitUntil: "networkidle" });
await page.locator("#pinSidePanelBtn").click();
await page.waitForFunction(() => window.__sidePanelCalls.some(call => call.method === "open"));

const calls = await page.evaluate(() => window.__sidePanelCalls);
if (!calls.some(call => call.method === "setOptions" && call.options.path === "sidepanel.html" && call.options.enabled === true)) {
  throw new Error("sidePanel.setOptions was not called with sidepanel.html");
}
if (!calls.some(call => call.method === "open" && call.options.windowId === 42)) {
  throw new Error("sidePanel.open was not called for the current window");
}
if (errors.length) throw new Error(errors.join("\n"));

await browser.close();
server.close();
console.log("ApplyPaste side panel pin verified");
