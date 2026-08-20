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
const page = await browser.newPage({ viewport: { width: 820, height: 760 } });
const errors = [];
page.on("pageerror", error => errors.push(error.message));
page.on("console", message => {
  if (message.type() === "error") errors.push(message.text());
});
page.on("dialog", async dialog => {
  await dialog.accept();
});

await page.addInitScript(({ port }) => {
  const readMemory = () => JSON.parse(localStorage.getItem("applyPasteWordImportStorage") || "{}");
  const writeMemory = memory => localStorage.setItem("applyPasteWordImportStorage", JSON.stringify(memory));
  writeMemory({});
  window.chrome = {
    runtime: {
      getURL: path => `http://127.0.0.1:${port}/${path}`,
      onMessage: { addListener: () => {} },
      sendMessage: async () => {}
    },
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
    tabs: { query: async () => [{ id: 1, url: "https://example.com/job" }], create: async () => ({}) },
    scripting: { executeScript: async () => [] }
  };
  window.showOpenFilePicker = async () => {
    const { createProfileDocx } = await import(`http://127.0.0.1:${port}/src/docx.js`);
    const blob = createProfileDocx({
      categories: ["基础信息", "教育背景", "项目经历", "实习经历", "技能", "荣誉奖项"],
      items: [
        { category: "基础信息", title: "张三", content: "姓名：张三\n手机号码：13800000000\n邮箱：test@example.com", fields: [] },
        { category: "教育背景", title: "XX大学", content: "XX大学\n专业：材料科学与工程\n硕士\n2024.09-2027.06", fields: [] },
        { category: "项目经历", title: "项目A", content: "项目概况：虚构材料项目\n项目职责：完成公开测试\n项目成果：形成测试结果", fields: [] },
        { category: "实习经历", title: "公司A", content: "岗位名称：测试工程师\n2025.01-2025.06", fields: [] },
        { category: "技能", title: "技能", content: "Python\nMatlab", fields: [] },
        { category: "荣誉奖项", title: "一等奖", content: "虚构测试奖项", fields: [] }
      ]
    });
    return [{ getFile: async () => new File([blob], "applypaste-import-smoke.docx", { type: blob.type, lastModified: Date.now() }) }];
  };
}, { port });

await page.goto(`http://127.0.0.1:${port}/popup.html`, { waitUntil: "networkidle" });
await page.locator("#feishuSettingsBtn").click();
await page.locator(".dataImportMenu summary").click();
await page.locator("#resumeImportBtn").click();
await page.locator("#resumeImportDialog[open]").waitFor();
await page.locator("#confirmResumeImportBtn").click();
try {
  await page.waitForFunction(async () => {
    const db = await new Promise((resolve, reject) => {
      const open = indexedDB.open("ApplyPasteInternal", 1);
      open.onerror = () => reject(open.error);
      open.onsuccess = () => resolve(open.result);
    });
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(["data", "attachments"], "readonly");
        const profileRequest = tx.objectStore("data").get("profile");
        const attachmentRequest = tx.objectStore("attachments").getAll();
        tx.oncomplete = () => resolve((profileRequest.result?.items?.length || 0) >= 4 && attachmentRequest.result.length === 1);
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  }, null, { timeout: 10000 });
} catch (error) {
  const debug = await page.evaluate(() => ({
    cards: document.querySelectorAll(".card").length,
    notice: document.querySelector("#notice")?.textContent || "",
    dialogOpen: [...document.querySelectorAll("dialog")].filter(dialog => dialog.open).map(dialog => dialog.id)
  }));
  throw new Error(`Word import did not render cards: ${JSON.stringify(debug)}\n${errors.join("\n")}`);
}

const result = await page.evaluate(async () => {
  const data = await new Promise((resolve, reject) => {
    const open = indexedDB.open("ApplyPasteInternal", 1);
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = db.transaction(["data", "attachments"], "readonly");
      const dataStore = tx.objectStore("data");
      const attachmentStore = tx.objectStore("attachments");
      const profileRequest = dataStore.get("profile");
      const attachmentRequest = attachmentStore.getAll();
      tx.oncomplete = () => {
        db.close();
        resolve({ profile: profileRequest.result, attachments: attachmentRequest.result });
      };
      tx.onerror = () => reject(tx.error);
    };
  });
  return {
    cardCount: document.querySelectorAll(".card").length,
    attachmentCount: data.attachments.length,
    itemTitles: data.profile.items.map(item => item.title),
    categories: data.profile.categories
  };
});

if (errors.length) throw new Error(errors.join("\n"));
if (result.itemTitles.length < 4) throw new Error(`expected imported profile items, got ${result.itemTitles.length}`);
if (result.attachmentCount !== 1) throw new Error(`expected one imported attachment, got ${result.attachmentCount}`);
if (!result.itemTitles.some(title => title.includes("张三"))) throw new Error("basic resume item was not imported");
if (!result.categories.includes("项目经历")) throw new Error("profile categories were not preserved");

await browser.close();
server.close();
console.log("ApplyPaste Word import smoke verified");
