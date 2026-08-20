import http from "node:http";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";
import { homedir, tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import playwright from "./helpers/load-playwright.mjs";

const { chromium } = playwright;
const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const playwrightRoot = join(homedir(), "AppData", "Local", "ms-playwright");
const chromiumFolder = (await readdir(playwrightRoot)).filter(name => /^chromium-\d+$/.test(name)).sort().at(-1);
const browserExecutable = join(playwrightRoot, chromiumFolder, "chrome-win64", "chrome.exe");
const profile = await mkdtemp(join(tmpdir(), "applypaste-public-e2e-"));
const errors = [];

const server = http.createServer(async (request, response) => {
  try {
    const relative = decodeURIComponent(new URL(request.url, "http://localhost").pathname).replace(/^\/+/, "") || "tests/fixtures/job-page-standard.html";
    const target = normalize(join(root, relative));
    if (!target.startsWith(normalize(root))) throw new Error("outside root");
    const body = await readFile(target);
    response.writeHead(200, { "content-type": extname(target) === ".html" ? "text/html; charset=utf-8" : "application/octet-stream" });
    response.end(body);
  } catch {
    response.writeHead(404); response.end("not found");
  }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const port = server.address().port;
const appSource = await readFile(join(root, "src", "app.js"), "utf8");
const extractorSource = appSource.slice(appSource.indexOf("function extractStructuredJobPosting"), appSource.indexOf("\nfunction renderRecordSuggestions", appSource.indexOf("function extractStructuredJobPosting")));

async function extensionId(context) {
  const page = context.pages()[0] || await context.newPage();
  await page.goto("chrome://extensions/");
  await page.waitForTimeout(1200);
  const readItems = () => page.evaluate(() => {
    const all = [];
    const visit = root => {
      for (const item of root.querySelectorAll?.("extensions-item") || []) all.push(item);
      for (const element of root.querySelectorAll?.("*") || []) if (element.shadowRoot) visit(element.shadowRoot);
    };
    visit(document);
    return all.map(item => ({ id: item.id, name: item.shadowRoot?.querySelector("#name")?.textContent?.trim() || "" }));
  });
  let items = await readItems();
  let extension = items.find(item => item.name === "ApplyPaste");
  if (!extension) {
    try {
      const toggleFound = await page.evaluate(() => {
        const deep = (root, selector) => {
          const found = root.querySelector?.(selector); if (found) return found;
          for (const element of root.querySelectorAll?.("*") || []) if (element.shadowRoot) { const nested = deep(element.shadowRoot, selector); if (nested) return nested; }
          return null;
        };
        const toggle = deep(document, "#devMode");
        if (toggle && !toggle.checked) toggle.click();
        return Boolean(toggle);
      });
      if (!toggleFound) throw new Error("developer mode toggle not found");
      await page.waitForTimeout(500);
      const chooserPromise = page.waitForEvent("filechooser", { timeout: 5000 });
      const loadFound = await page.evaluate(() => {
        const deep = (root, selector) => {
          const found = root.querySelector?.(selector); if (found) return found;
          for (const element of root.querySelectorAll?.("*") || []) if (element.shadowRoot) { const nested = deep(element.shadowRoot, selector); if (nested) return nested; }
          return null;
        };
        const button = deep(document, "#loadUnpacked"); button?.click(); return Boolean(button);
      });
      if (!loadFound) throw new Error("load unpacked button not found");
      const chooser = await chooserPromise;
      await chooser.setFiles(root);
      await page.waitForTimeout(1000);
      items = await readItems();
      extension = items.find(item => item.name === "ApplyPaste");
    } catch (error) { errors.push(`manual Edge load: ${error.message}`); }
  }
  if (!extension) {
    try {
      const preferences = JSON.parse(await readFile(join(profile, "Default", "Preferences"), "utf8"));
      const entry = Object.entries(preferences.extensions?.settings || {}).find(([, value]) => normalize(value.path || "") === normalize(root));
      if (entry) extension = { id: entry[0], name: "ApplyPaste" };
    } catch {}
  }
  if (!extension) {
    await page.goto("chrome://version/");
    const commandLine = await page.locator("#command_line").textContent().catch(() => "unavailable");
    throw new Error(`ApplyPaste was not loaded: ${JSON.stringify(items)}; ${errors.join("; ")}; command=${commandLine}`);
  }
  return extension.id;
}

async function installPicker(page) {
  await page.evaluate(() => {
    window.__applyPasteFileKind = "pdf";
    window.showOpenFilePicker = async () => {
      const kind = window.__applyPasteFileKind;
      if (kind === "png") {
        const bytes = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="), char => char.charCodeAt(0));
        return [{ getFile: async () => new File([bytes], "test-image.png", { type: "image/png" }) }];
      }
      if (kind === "docx") {
        const { createProfileDocx } = await import("./src/docx.js");
        const blob = createProfileDocx({ categories: ["基础信息"], items: [{ category: "基础信息", title: "公开测试", content: "姓名：张三" }] });
        return [{ getFile: async () => new File([blob], "test-resume.docx", { type: blob.type }) }];
      }
      return [{ getFile: async () => new File(["%PDF-1.4\n1 0 obj<</Type /Page>>endobj\n%%EOF"], "test-document.pdf", { type: "application/pdf" }) }];
    };
  });
}

async function addAttachment(page, kind, category) {
  await page.evaluate(value => { window.__applyPasteFileKind = value; }, kind);
  await page.locator("#addAttachmentBtn").click();
  await page.locator("#attachmentDialog[open]").waitFor();
  await page.locator("#attachmentCategory").selectOption(category);
  await page.locator("#attachmentForm button.primary").click();
  await page.locator("#attachmentDialog").waitFor({ state: "hidden" });
}

async function readState(page) {
  return page.evaluate(async () => (await import("./src/internal-db.js")).loadInternalState());
}

let context;
try {
  context = await chromium.launchPersistentContext(profile, {
    headless: false,
    executablePath: browserExecutable,
    ignoreDefaultArgs: ["--disable-extensions"],
    viewport: { width: 1040, height: 820 },
    args: ["--window-position=-32000,-32000", `--disable-extensions-except=${root}`, `--load-extension=${root}`]
  });
  const id = await extensionId(context);
  for (const fixture of ["job-page-standard.html", "job-page-chinese.html", "job-page-minimal.html", "job-page-dynamic.html"]) {
    const fixturePage = await context.newPage();
    await fixturePage.goto(`http://127.0.0.1:${port}/tests/fixtures/${fixture}`, { waitUntil: "networkidle" });
    const recognized = await fixturePage.evaluate(source => (0, eval)(`(${source})`)(), extractorSource);
    if (!recognized.job || !recognized.jd) throw new Error(`job page recognition failed for ${fixture}: ${JSON.stringify(recognized)}`);
    if (fixture !== "job-page-minimal.html" && (!recognized.company || !recognized.location)) throw new Error(`job page fields incomplete for ${fixture}: ${JSON.stringify(recognized)}`);
    await fixturePage.close();
  }
  const page = await context.newPage();
  page.on("pageerror", error => errors.push(error.message));
  await page.goto(`chrome-extension://${id}/popup.html?mode=workbench`, { waitUntil: "networkidle" });

  if (await page.locator(".card").count()) throw new Error("fresh install contains profile cards");
  const initial = await readState(page);
  if (initial.attachments.length || initial.records.length || initial.versions.length !== 1 || initial.versions[0].id !== "general") throw new Error("fresh install is not empty/general-only");
  const local = await page.evaluate(() => chrome.storage.local.get(null));
  if (Object.values(local).some(value => JSON.stringify(value).includes("appSecret"))) throw new Error("fresh install contains Feishu secret data");

  await page.locator("#emptyNewItemBtn").click();
  await page.locator("#editCategory").selectOption("项目经历");
  await page.locator("#editTitle").fill("公开测试项目");
  await page.locator("#editContent").fill("项目名称：公开测试项目\n项目描述：通用描述\n项目成果：测试通过");
  await page.locator("#editorForm button.primary").click();
  await page.locator(".card").first().waitFor();

  await page.locator("#addVersionBtn").click();
  await page.locator("#versionName").fill("测试岗位版");
  await page.locator("#versionForm button.primary").click();
  await page.waitForFunction(() => document.querySelector("#currentVersion")?.value !== "general");
  const versionId = await page.locator("#currentVersion").inputValue();
  if (versionId === "general") throw new Error("version creation failed");
  const description = page.locator(".fieldRow").filter({ has: page.locator(".fieldLabel", { hasText: "项目描述" }) });
  await description.locator(".fieldEdit").click();
  await page.locator("#fieldVersioned").check();
  await page.locator("#fieldEditValue").fill("岗位版描述");
  await page.locator("#fieldEditorForm button.primary").click();
  await page.locator("#fieldEditorDialog").waitFor({ state: "hidden" });
  await page.waitForFunction(() => /[1-9]/.test(document.querySelector("#showDifferencesBtn")?.textContent || ""));
  if (!(await page.locator("#showDifferencesBtn").textContent()).includes("1")) throw new Error("version override/diff failed");

  const side = await context.newPage();
  side.on("pageerror", error => errors.push(error.message));
  await side.goto(`chrome-extension://${id}/popup.html?mode=sidepanel`, { waitUntil: "networkidle" });
  for (const selector of ["#currentVersion", "#differenceBar", "#searchInput", "#sidepanelCategory", "#quickNavigationSelect"]) if (!await side.locator(selector).isVisible()) throw new Error(`side panel control missing: ${selector}`);
  await side.locator("#searchInput").fill("公开测试项目");
  if (await side.locator(".card").count() !== 1) throw new Error("side panel search failed");
  await side.locator("#searchInput").fill("");
  await side.locator("#sidepanelCategory").selectOption("项目经历");
  await side.locator("#quickNavigationSelect").selectOption(await side.locator(".card").first().getAttribute("data-item-id"));

  await installPicker(side);
  await side.locator("#sidepanelCategory").selectOption("附件");
  await addAttachment(side, "pdf", "其他");
  await addAttachment(side, "png", "其他");
  await addAttachment(side, "docx", "简历");
  if (await side.locator(".attachmentRow").count() !== 3) throw new Error("attachment upload failed");
  const attachmentAudit = await side.evaluate(async () => (await (await import("./src/internal-db.js")).loadInternalState()).attachments.map(item => ({ blob: item.blob instanceof Blob, size: item.size, mimeType: item.mimeType })));
  if (attachmentAudit.some(item => !item.blob || !item.size || !item.mimeType)) throw new Error(`attachment Blob/metadata failed: ${JSON.stringify(attachmentAudit)}`);
  const rowByFile = file => side.locator(".attachmentRow").filter({ hasText: file }).first();
  await rowByFile("test-document.pdf").getByRole("button", { name: "预览", exact: true }).click();
  await side.locator('[data-preview-type="pdf"]').waitFor(); await side.locator("#closeAttachmentPreviewBtn").click();
  const originalImageId = await rowByFile("test-image.png").getAttribute("data-attachment-id");
  await rowByFile("test-image.png").getByRole("button", { name: "预览", exact: true }).click();
  await side.locator('[data-preview-type="image"]').waitFor(); await side.locator("#closeAttachmentPreviewBtn").click();
  await rowByFile("test-resume.docx").getByRole("button", { name: "预览", exact: true }).click();
  await side.locator('[data-preview-type="docx"]').waitFor(); await side.locator("#closeAttachmentPreviewBtn").click();
  await side.evaluate(() => { window.__applyPasteFileKind = "png"; });
  await rowByFile("test-document.pdf").getByRole("button", { name: "替换", exact: true }).click();
  await side.locator("#attachmentReplaceDialog[open]").waitFor(); await side.locator("#confirmReplacementBtn").click();
  await side.locator("#attachmentReplaceDialog").waitFor({ state: "hidden" });
  side.once("dialog", dialog => dialog.accept());
  const originalImageRow = side.locator(`[data-attachment-id="${originalImageId}"]`); await originalImageRow.locator("summary").click(); await originalImageRow.getByRole("button", { name: "删除附件", exact: true }).click();
  await side.waitForFunction(() => document.querySelectorAll(".attachmentRow").length === 2);
  const resumeRow = rowByFile("test-resume.docx"); await resumeRow.locator("summary").click(); await resumeRow.getByRole("button", { name: "设置为当前版本默认附件", exact: true }).click();

  await page.bringToFront();
  await page.reload({ waitUntil: "networkidle" });
  await page.locator("#completeBtn").click();
  await page.locator("#completeCompany").fill("示例科技");
  await page.locator("#completeJob").fill("测试工程师");
  await page.locator("#completeLocation").fill("测试城市");
  if (await page.locator("#completeResume").inputValue() !== "test-resume") throw new Error("default resume was not used by application flow");
  await page.locator("#completeUrl").fill(`http://127.0.0.1:${port}/tests/fixtures/job-page-standard.html`);
  await page.locator("#completeJd").fill("负责公开测试与质量验证。");
  await page.locator("#completeSubmitBtn").click();
  await page.locator("#completeDialog").waitFor({ state: "hidden" });
  const saved = await readState(page);
  if (saved.records.length !== 1 || saved.records[0]["状态"] !== "已投递" || !saved.records[0]["记录ID"] || saved.records[0]["飞书同步状态"] !== "未同步") throw new Error("local application record failed without Feishu");

  await context.close(); context = null;
  context = await chromium.launchPersistentContext(profile, { headless: false, executablePath: browserExecutable, ignoreDefaultArgs: ["--disable-extensions"], viewport: { width: 1040, height: 820 }, args: ["--window-position=-32000,-32000", `--disable-extensions-except=${root}`, `--load-extension=${root}`] });
  const restartedId = await extensionId(context);
  const restarted = await context.newPage();
  await restarted.goto(`chrome-extension://${restartedId}/popup.html?mode=workbench`, { waitUntil: "networkidle" });
  const persisted = await readState(restarted);
  if (!persisted.profile.items.length || persisted.versions.length < 2 || persisted.attachments.length !== 2 || persisted.records.length !== 1) throw new Error("restart persistence failed");
  if (errors.length) throw new Error(errors.join("\n"));
  console.log("E2E PASS: unpacked ApplyPaste loaded in a fresh Chromium extension profile");
  console.log("EDGE AUTOMATION WARN: branded Edge blocks command-line extension loading; use the manual Edge checklist for the final surface check");
  console.log("PERSISTENCE PASS: profile, versions, attachments, records, and settings survived browser restart");
} finally {
  await context?.close().catch(() => {});
  server.close();
  await rm(profile, { recursive: true, force: true });
}
