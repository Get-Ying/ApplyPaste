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
    const target = normalize(join(root, relative));
    if (!target.startsWith(normalize(root))) throw new Error("outside root");
    response.writeHead(200, { "content-type": mime[extname(target)] || "application/octet-stream" });
    response.end(await readFile(target));
  } catch {
    if (!response.headersSent) response.writeHead(404);
    response.end("not found");
  }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const port = server.address().port;
const browser = await chromium.launch({ headless: true, executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" });

function fixture() {
  const itemId = "test-item-project";
  return {
    profile: { version: 6, categories: ["项目经历", "荣誉奖项"], items: [{ id: itemId, category: "项目经历", title: "测试项目", tags: ["测试"], fields: [
      { id: "field-name", label: "项目名称", value: "通用项目", versioned: true },
      { id: "field-role", label: "项目角色", value: "通用角色", versioned: true },
      { id: "field-result", label: "成果", value: "通用成果", versioned: false }
    ], content: "项目名称: 通用项目\n项目角色: 通用角色\n成果: 通用成果" }, ...["二", "三", "四"].map((suffix, index) => ({ id: `test-item-${suffix}`, category: "项目经历", title: `测试项目${suffix}`, tags: ["测试"], fields: [{ id: `field-${suffix}`, label: "项目名称", value: `通用项目${suffix}`, versioned: true }], content: `项目名称: 通用项目${suffix}` }))] },
    versions: [{ id: "general", name: "通用版", fieldOverrides: {}, itemOverrides: {} }, { id: "target", name: "目标岗位", fieldOverrides: {}, itemOverrides: {} }],
    records: [],
    attachments: []
  };
}

async function seed(page) {
  await page.waitForTimeout(100);
  await page.evaluate(async data => {
    const db = await import("./src/internal-db.js");
    await db.saveCompleteState(data);
  }, fixture());
}

async function installChromeMock(page) {
  await page.addInitScript(({ testPort }) => {
    const read = () => JSON.parse(localStorage.getItem("p1ChromeStorage") || "{}");
    const write = value => localStorage.setItem("p1ChromeStorage", JSON.stringify(value));
    Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async () => {} } });
    window.chrome = {
      runtime: { getURL: path => `http://127.0.0.1:${testPort}/${path}`, onMessage: { addListener: () => {} }, sendMessage: async () => {} },
      storage: { local: {
        get: async keys => { const data = read(); const list = Array.isArray(keys) ? keys : [keys]; return Object.fromEntries(list.filter(key => key in data).map(key => [key, data[key]])); },
        set: async values => write({ ...read(), ...values }),
        remove: async keys => { const data = read(); for (const key of (Array.isArray(keys) ? keys : [keys])) delete data[key]; write(data); }
      } },
      tabs: { query: async () => [{ id: 1, windowId: 1, url: "https://example.test/job" }], create: async () => ({}) },
      downloads: { download: async () => 1 },
      sidePanel: { open: async () => {}, setOptions: async () => {} }
    };
  }, { testPort: port });
}

async function openPage(context, mode) {
  const page = await context.newPage();
  await installChromeMock(page);
  await page.goto(`http://127.0.0.1:${port}/popup.html?mode=${mode}`, { waitUntil: "networkidle" });
  return page;
}

async function versionTest() {
  const context = await browser.newContext({ viewport: { width: 980, height: 780 } });
  const page = await openPage(context, "workbench");
  await seed(page);
  await page.reload({ waitUntil: "networkidle" });
  await page.locator(`[data-item-id="test-item-project"]`).waitFor();
  if (await page.locator("#differenceBar").isVisible()) throw new Error("通用版不应显示差异控件");
  if (!await page.locator("#currentVersion option[value=general]").count()) throw new Error("通用版不存在");

  await page.locator("#addVersionBtn").click();
  await page.locator("#versionName").fill("测试岗位");
  await page.locator("#versionForm button.primary").click();
  await page.waitForFunction(() => document.querySelector("#currentVersion")?.value !== "general");
  const createdId = await page.locator("#currentVersion").inputValue();
  await page.locator(".versionMore").locator("summary").click();
  if (createdId === "general") throw new Error(`新建岗位版本失败：${await page.locator("#currentVersion").locator("option").allTextContents()}`);
  await page.locator("#renameVersionBtn").click();
  await page.locator("#versionName").fill("重命名岗位");
  await page.locator("#versionForm button.primary").click();
  await page.waitForFunction(id => document.querySelector(`#currentVersion option[value="${id}"]`)?.textContent.includes("重命名岗位"), createdId);
  if (!(await page.locator(`#currentVersion option[value="${createdId}"]`).textContent()).includes("重命名岗位")) throw new Error("重命名失败");
  await page.locator(".versionMore").locator("summary").click();
  await page.locator("#copyVersionBtn").click();
  await page.locator("#versionName").fill("复制岗位");
  await page.locator("#versionForm button.primary").click();
  await page.waitForFunction(id => document.querySelector("#currentVersion")?.value !== id, createdId);
  const copiedId = await page.locator("#currentVersion").inputValue();
  if (copiedId === createdId) throw new Error("复制版本失败");
  page.once("dialog", dialog => dialog.accept());
  await page.locator(".versionMore").locator("summary").click();
  await page.locator("#deleteVersionBtn").click();
  await page.locator(".versionMore").locator("summary").click();
  if (await page.locator(`#currentVersion option[value="${copiedId}"]`).count()) throw new Error("删除普通版本失败");
  await page.locator("#currentVersion").selectOption(createdId);
  await page.locator(".versionMore").evaluate(menu => { menu.open = false; });

  const row = page.locator(`[data-item-id="test-item-project"] .fieldRow`).filter({ has: page.locator(".fieldLabel", { hasText: "项目名称" }) });
  await row.locator(".fieldEdit").click();
  await page.locator("#fieldEditValue").fill("岗位项目");
  await page.locator("#fieldEditorForm button.primary").click();
  await page.locator("#fieldEditorDialog").waitFor({ state: "hidden" });
  if (!await row.locator(".fieldScope.override").count()) throw new Error("当前版本 override 未显示");
  if (await page.locator("#showDifferencesBtn").textContent() !== "本版本差异 1") throw new Error("差异数量不正确");
  await page.locator("#showDifferencesBtn").click();
  if (await page.locator("#list .card").count() !== 1 || !await page.locator("#list .fieldScope.override").count()) throw new Error("差异过滤不正确");
  await page.locator("#showAllBtn").click();
  await row.locator(".fieldEdit").click();
  await page.locator("#useCommonFieldAnswerBtn").click();
  await page.waitForTimeout(100);
  if (!(await page.locator("#showDifferencesBtn").textContent()).includes("0")) throw new Error(`使用通用答案未删除 override：${await page.locator("#showDifferencesBtn").textContent()} / ${JSON.stringify(await page.evaluate(async () => (await (await import("./src/internal-db.js")).loadInternalState()).versions))}`);
  await page.reload({ waitUntil: "networkidle" });
  if (await page.locator("#currentVersion").inputValue() !== createdId) throw new Error("刷新后版本未保留");
  await page.locator("#currentVersion").selectOption("general");
  await page.locator(".versionMore").locator("summary").click();
  if (!await page.locator("#deleteVersionBtn").isDisabled()) throw new Error("通用版删除按钮未禁用");
  if (!await page.locator("#currentVersion option[value=general]").count()) throw new Error("通用版被删除");
  await context.close();
  return "版本管理 14 项场景通过";
}

async function sidePanelTest() {
  const context = await browser.newContext({ viewport: { width: 420, height: 760 } });
  const page = await openPage(context, "sidepanel");
  await seed(page);
  await page.reload({ waitUntil: "networkidle" });
  await page.locator(`[data-item-id="test-item-project"]`).waitFor();
  await page.locator("#currentVersion").selectOption("target");
  for (const selector of ["#currentVersion", "#showAllBtn", "#showDifferencesBtn", "#searchInput", "#sidepanelCategory", "#quickNavigationSelect"]) if (!await page.locator(selector).isVisible()) throw new Error(`Side Panel 控件未初始化：${selector}`);
  const sticky = page.locator("#sidepanelStickyNavigation");
  if (!await sticky.isVisible()) throw new Error("sticky 导航未显示");
  await page.locator("#sidepanelCategory").selectOption("项目经历");
  await page.locator("#searchInput").fill("项目");
  await page.locator(".card .copyBtn").first().click();
  if (await page.locator("#searchInput").inputValue() !== "项目") throw new Error("复制后搜索词丢失");
  await page.locator("#searchInput").fill("");
  if (await page.locator("#showDifferencesBtn").isEnabled()) {
    await page.locator("#showDifferencesBtn").click();
    await page.locator("#showAllBtn").click();
  }
  await page.locator("#quickNavigationSelect").selectOption("test-item-project");
  const sideField = page.locator(`[data-item-id="test-item-project"] .fieldRow`).filter({ has: page.locator(".fieldLabel", { hasText: "项目名称" }) });
  await sideField.locator(".fieldEdit").click();
  await page.locator("#fieldEditValue").fill("侧边栏编辑");
  await page.locator("#fieldEditorForm button.primary").click();
  await page.locator("#fieldEditorDialog").waitFor({ state: "hidden" });
  if (!await sideField.locator(".fieldValue").textContent().then(text => text.includes("侧边栏编辑"))) throw new Error("Side Panel 字段编辑失败");
  await page.locator("#contentPane").evaluate(node => node.scrollTop = node.scrollHeight);
  for (const selector of ["#currentVersion", "#showAllBtn", "#showDifferencesBtn", "#searchInput", "#sidepanelCategory", "#quickNavigationSelect"]) if (!await page.locator(selector).isVisible()) throw new Error(`滚到底部后控件消失：${selector}`);
  const audit = await page.evaluate(() => ({ horizontal: document.documentElement.scrollWidth > document.documentElement.clientWidth + 1, bodyScroll: document.body.scrollHeight > document.body.clientHeight + 1, paneScroll: document.querySelector("#contentPane").scrollHeight > document.querySelector("#contentPane").clientHeight + 1 }));
  if (audit.horizontal || audit.bodyScroll || !audit.paneScroll) throw new Error(`滚动容器异常：${JSON.stringify(audit)}`);
  const footer = await page.locator(".submissionBar").boundingBox();
  if (!footer || Math.abs(footer.y + footer.height - 760) > 3) throw new Error("底部投递栏未固定");
  const projectPosition = await page.locator("#contentPane").evaluate(node => node.scrollTop);
  await page.locator("#sidepanelCategory").selectOption("荣誉奖项");
  await page.locator("#contentPane").evaluate(node => node.scrollTop = 0);
  await page.locator("#sidepanelCategory").selectOption("项目经历");
  if (Math.abs((await page.locator("#contentPane").evaluate(node => node.scrollTop)) - projectPosition) > 8) throw new Error("分类滚动位置恢复异常");
  const workbench = await openPage(context, "workbench");
  await workbench.locator(`[data-item-id="test-item-project"]`).waitFor();
  if (!await workbench.locator(`[data-item-id="test-item-project"] .fieldValue`).first().textContent().then(text => text.includes("侧边栏编辑"))) throw new Error("工作台未读取 Side Panel 编辑结果");
  await workbench.close();
  await context.close();
  return "Side Panel 16 项场景通过";
}

try {
  console.log(await versionTest());
  console.log(await sidePanelTest());
} finally {
  await browser.close();
  server.close();
}
