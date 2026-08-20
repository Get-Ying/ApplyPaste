import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const root = resolve(fileURLToPath(new URL("../", import.meta.url)));
const failures = [];
const structural = [];
const behavior = [];

const read = path => readFile(resolve(root, path), "utf8");
const check = (condition, label, group = structural) => condition ? group.push(label) : failures.push(label);

const manifest = JSON.parse(await read("manifest.json"));
const popup = await read("popup.html");
const app = await read("src/app.js");
const db = await read("src/internal-db.js");
const template = JSON.parse(await read("public-template/data/profile.json"));
const settings = JSON.parse(await read("public-template/data/settings.json"));

check(manifest.manifest_version === 3, "Manifest V3");
check(manifest.name === "ApplyPaste" && manifest.action?.default_title === "ApplyPaste", "ApplyPaste brand");
check(manifest.side_panel?.default_path === "sidepanel.html" && manifest.permissions?.includes("sidePanel"), "Side Panel registration");
check(manifest.permissions?.includes("storage") && manifest.permissions?.includes("unlimitedStorage"), "storage permissions");
check(/ApplyPasteInternal/.test(db) && !/indexedDB\.open\(["']JobPasteInternal/.test(db), "isolated IndexedDB");
check(/applyPasteViewState/.test(app) && !/jobPasteViewState/.test(app), "isolated UI storage");
for (const id of ["currentVersion", "differenceBar", "searchInput", "sidepanelCategory", "quickNavigationSelect", "completeBtn", "laterBtn", "abortBtn"]) check(popup.includes(`id="${id}"`), `DOM #${id}`);
for (const fn of ["handleVersionSubmit", "setDifferenceMode", "useCommonFieldAnswer", "navigateToContent", "handleSearchInput", "confirmResumeImport", "previewAttachment", "handleCompleteSubmit", "syncRecord"]) check(app.includes(`function ${fn}`) || app.includes(`async function ${fn}`), `logic ${fn}`);
check(template.items.length === 0, "public profile is empty");
check(settings.versions.length === 1 && settings.versions[0].id === "general", "general version only");
check(["基础信息", "教育背景", "成绩课程", "项目经历", "附件"].every(category => category === "附件" || template.categories.includes(category)), "recommended public categories");
check(!template.categories.includes("基本信息"), "category naming converged");

for (const item of template.items) {
  for (const field of item.fields || []) {
    const multilineSafe = /开放题|描述|职责|成果|概况|回答|原文|备注/.test(`${item.category} ${field.label}`);
    if (!multilineSafe && /(?:^|\n)\s*[^\n：:]{1,30}[：:]\s*\S+/m.test(String(field.value || ""))) failures.push(`POTENTIAL_FIELD_GRANULARITY_ERROR ${item.title}/${field.label}`);
    if (["select", "multi-select"].includes(field.type) && (!Array.isArray(field.options) || !field.options.length)) failures.push(`INVALID_SELECT_OPTIONS ${item.title}/${field.label}`);
  }
}

for (const test of ["test-applypaste-isolation.mjs", "test-field-guide.mjs", "test-p1-core.mjs", "test-sidepanel-pin.mjs", "test-word-import-smoke.mjs"]) {
  const result = spawnSync(process.execPath, [resolve(root, "tests", test)], { cwd: root, encoding: "utf8", timeout: 90_000 });
  if (result.status === 0) behavior.push(test);
  else failures.push(`${test}: ${(result.stderr || result.stdout || "failed").trim().split(/\r?\n/).slice(-2).join(" ")}`);
}

if (failures.length) {
  console.error("ApplyPaste feature parity: FAIL");
  failures.forEach(item => console.error(`- ${item}`));
  process.exitCode = 1;
} else {
  console.log(`STRUCTURAL PASS (${structural.length})`);
  console.log(`BEHAVIOR PASS (${behavior.length} browser suites)`);
  console.log("E2E delegated to tests/test-public-release.mjs");
}
