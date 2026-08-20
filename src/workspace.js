import { toMarkdownLibrary } from "./markdown.js";

const DB_NAME = "applyPasteWorkspace";
const STORE_NAME = "handles";
const HANDLE_KEY = "root";

const DIRECTORY_TREE = [
  "data",
  "attachments/resumes",
  "attachments/transcripts",
  "attachments/certificates",
  "attachments/publications",
  "attachments/other",
  "exports",
  "history/deleted_attachments",
  "records",
  "imports"
];

export async function chooseWorkspace() {
  if (!window.showDirectoryPicker) throw new Error("当前 Edge 不支持本地工作区目录授权");
  const handle = await window.showDirectoryPicker({ mode: "readwrite", id: "apply-paste-workspace" });
  await ensurePermission(handle, true);
  await setHandle(handle);
  await ensureWorkspace(handle);
  return handle;
}

export async function restoreWorkspace() {
  const handle = await getHandle();
  if (!handle) return null;
  return (await ensurePermission(handle, false)) ? handle : null;
}

export async function requestWorkspacePermission(handle) {
  if (!handle) return false;
  const granted = await ensurePermission(handle, true);
  if (granted) await ensureWorkspace(handle);
  return granted;
}

export async function initializeWorkspace(handle, state) {
  await ensureWorkspace(handle);
  await writeJson(handle, "data/profile.json", state.profile);
  await writeJson(handle, "data/settings.json", state.settings);
  await writeJson(handle, "data/attachments.json", state.attachments);
  await writeLatestExports(handle, exportableProfile(state));
}

export async function persistWorkspace(handle, state, change, previousState = null, options = {}) {
  if (!handle || !(await ensurePermission(handle, false))) return false;
  await ensureWorkspace(handle);
  const stamp = timestamp();
  if (previousState) {
    const snapshot = {
      createdAt: new Date().toISOString(),
      change,
      profile: previousState.profile,
      settings: previousState.settings,
      attachments: previousState.attachments
    };
    const snapshotName = `profile_${stamp}.json`;
    await writeJson(handle, `history/${snapshotName}`, snapshot);
    const index = await readJson(handle, "history/index.json", []);
    index.unshift({ file: snapshotName, createdAt: snapshot.createdAt, change });
    await writeJson(handle, "history/index.json", index.slice(0, 200));
  }
  await writeJson(handle, "data/profile.json", state.profile);
  await writeJson(handle, "data/settings.json", state.settings);
  await writeJson(handle, "data/attachments.json", state.attachments);
  if (options.exportLatest) await writeLatestExports(handle, exportableProfile(state));
  return true;
}

export async function exportWorkspaceProfile(handle, profile, settings = null) {
  if (!handle || !(await ensurePermission(handle, false))) return false;
  await writeLatestExports(handle, settings ? { ...profile, fillingVersions: settings.versions } : profile);
  return true;
}

function exportableProfile(state) {
  return { ...state.profile, fillingVersions: state.settings.versions };
}

export async function readWorkspaceState(handle) {
  if (!handle || !(await ensurePermission(handle, false))) return null;
  return {
    profile: await readJson(handle, "data/profile.json", null),
    settings: await readJson(handle, "data/settings.json", null),
    attachments: await readJson(handle, "data/attachments.json", null)
  };
}

export async function readWorkspaceRecords(handle) {
  try {
    const file = await getFileHandle(handle, "records/applications.csv");
    const text = (await (await file.getFile()).text()).replace(/^\ufeff/, "");
    const rows = parseCsv(text);
    if (rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).filter(row => row.some(Boolean)).map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] || ""])));
  } catch (error) {
    if (error.name === "NotFoundError") return [];
    throw error;
  }
}

export async function listHistory(handle) {
  return readJson(handle, "history/index.json", []);
}

export async function readHistorySnapshot(handle, fileName) {
  return readJson(handle, `history/${safeName(fileName)}`, null);
}

export async function importAttachmentFile(handle, file, folder, mode = "keep") {
  const dir = await getDirectory(handle, `attachments/${folder}`, true);
  let fileName = safeName(file.name);
  const exists = await fileExists(dir, fileName);
  if (exists && mode === "cancel") return null;
  if (exists && mode === "keep") fileName = uniqueName(fileName);
  const target = await dir.getFileHandle(fileName, { create: true });
  const writable = await target.createWritable();
  await writable.write(file);
  await writable.close();
  return { fileName, path: `attachments/${folder}/${fileName}` };
}

export async function seedBundledAttachment(handle, item) {
  if (!item.path) return item;
  const folder = attachmentFolder(item.category);
  const dir = await getDirectory(handle, `attachments/${folder}`, true);
  if (!(await fileExists(dir, item.fileName))) {
    const response = await fetch(chrome.runtime.getURL(item.path));
    if (response.ok) {
      const target = await dir.getFileHandle(item.fileName, { create: true });
      const writable = await target.createWritable();
      await writable.write(await response.blob());
      await writable.close();
    }
  }
  return { ...item, localPath: `attachments/${folder}/${item.fileName}`, source: "workspace" };
}

export async function readAttachmentBlob(handle, item) {
  if (item.localPath) {
    if (!handle || !(await ensurePermission(handle, false))) throw new Error("请重新授权 ApplyPaste 工作区后再读取该附件");
    const fileHandle = await getFileHandle(handle, item.localPath);
    return fileHandle.getFile();
  }
  if (!item.path) throw new Error("附件没有可读取的本地文件");
  const response = await fetch(chrome.runtime.getURL(item.path));
  if (!response.ok) throw new Error("附件读取失败");
  return response.blob();
}

export async function removeAttachmentFile(handle, item) {
  if (!handle || !item.localPath) return;
  const blob = await readAttachmentBlob(handle, item);
  const archived = `${timestamp()}_${safeName(item.fileName)}`;
  const archiveDir = await getDirectory(handle, "history/deleted_attachments", true);
  const archiveHandle = await archiveDir.getFileHandle(archived, { create: true });
  const writable = await archiveHandle.createWritable();
  await writable.write(blob);
  await writable.close();
  const parts = item.localPath.split("/");
  const fileName = parts.pop();
  const parent = await getDirectory(handle, parts.join("/"), false);
  await parent.removeEntry(fileName);
}

export async function moveAttachmentFile(handle, item, nextFolder) {
  if (!handle || !item.localPath) return item;
  const blob = await readAttachmentBlob(handle, item);
  const result = await importAttachmentFile(handle, new File([blob], item.fileName, { type: item.mimeType }), nextFolder, "keep");
  const parts = item.localPath.split("/");
  const oldName = parts.pop();
  const oldDir = await getDirectory(handle, parts.join("/"), false);
  await oldDir.removeEntry(oldName);
  return { ...item, fileName: result.fileName, localPath: result.path };
}

export function attachmentFolder(category) {
  if (/简历/.test(category)) return "resumes";
  if (/成绩/.test(category)) return "transcripts";
  if (/论文|专利/.test(category)) return "publications";
  if (/证明|证件|身份|语言/.test(category)) return "certificates";
  return "other";
}

async function writeLatestExports(handle, profile) {
  const name = "applypaste-profile-latest";
  await writeJson(handle, `exports/${name}.json`, profile);
  await writeText(handle, `exports/${name}.md`, toMarkdownLibrary(profile));
}

async function ensureWorkspace(handle) {
  for (const path of DIRECTORY_TREE) await getDirectory(handle, path, true);
}

async function ensurePermission(handle, request) {
  const options = { mode: "readwrite" };
  if ((await handle.queryPermission(options)) === "granted") return true;
  return request && (await handle.requestPermission(options)) === "granted";
}

async function getDirectory(root, path, create) {
  let current = root;
  for (const part of path.split("/").filter(Boolean)) {
    current = await current.getDirectoryHandle(part, { create });
  }
  return current;
}

async function getFileHandle(root, path) {
  const parts = path.split("/");
  const fileName = parts.pop();
  const dir = await getDirectory(root, parts.join("/"), false);
  return dir.getFileHandle(fileName);
}

async function writeJson(root, path, value) {
  await writeText(root, path, JSON.stringify(value, null, 2));
}

async function writeText(root, path, text) {
  const parts = path.split("/");
  const fileName = parts.pop();
  const dir = await getDirectory(root, parts.join("/"), true);
  const file = await dir.getFileHandle(fileName, { create: true });
  const writable = await file.createWritable();
  await writable.write(text);
  await writable.close();
}

async function readJson(root, path, fallback) {
  try {
    const file = await getFileHandle(root, path);
    return JSON.parse(await (await file.getFile()).text());
  } catch (error) {
    if (error.name === "NotFoundError") return fallback;
    throw error;
  }
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE_NAME);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function setHandle(handle) {
  const db = await openDb();
  await new Promise((resolve, reject) => {
    const request = db.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).put(handle, HANDLE_KEY);
    request.onsuccess = resolve;
    request.onerror = () => reject(request.error);
  });
  db.close();
}

async function getHandle() {
  const db = await openDb();
  const result = await new Promise((resolve, reject) => {
    const request = db.transaction(STORE_NAME).objectStore(STORE_NAME).get(HANDLE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return result;
}

async function fileExists(dir, name) {
  try {
    await dir.getFileHandle(name);
    return true;
  } catch (error) {
    if (error.name === "NotFoundError") return false;
    throw error;
  }
}

function timestamp() {
  return new Date().toISOString().replace(/[-:]/g, "").replace("T", "_").replace(".", "_").replace("Z", "");
}

function safeName(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, "_");
}

function uniqueName(name) {
  const dot = name.lastIndexOf(".");
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : "";
  return `${base}_${timestamp()}${ext}`;
}

function parseCsv(text) {
  const rows = []; let row = [], value = "", quoted = false;
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') { value += '"'; index++; }
      else if (char === '"') quoted = false;
      else value += char;
    } else if (char === '"') quoted = true;
    else if (char === ",") { row.push(value); value = ""; }
    else if (char === "\n") { row.push(value.replace(/\r$/, "")); rows.push(row); row = []; value = ""; }
    else value += char;
  }
  if (value || row.length) { row.push(value.replace(/\r$/, "")); rows.push(row); }
  return rows;
}
