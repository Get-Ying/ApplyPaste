import { parseMarkdownLibrary, toMarkdownLibrary } from "./markdown.js";
import { createProfileDocx } from "./docx.js";

const STORAGE_KEY = "applyPasteProfile";
const FILE_META_KEY = "applyPasteFileMeta";

export async function loadBundledProfile() {
  const response = await fetch(chrome.runtime.getURL("data/applypaste-profile.json"));
  if (!response.ok) throw new Error("默认信息库读取失败");
  return response.json();
}

export async function loadStoredProfile() {
  const stored = await chrome.storage.local.get([STORAGE_KEY, FILE_META_KEY]);
  return {
    profile: stored[STORAGE_KEY] || null,
    fileMeta: stored[FILE_META_KEY] || null
  };
}

export async function saveStoredProfile(profile, fileMeta = null) {
  const next = { ...profile, updatedAt: new Date().toISOString() };
  await chrome.storage.local.set({
    [STORAGE_KEY]: next,
    ...(fileMeta ? { [FILE_META_KEY]: fileMeta } : {})
  });
  return next;
}

export async function openProfileFile() {
  const [handle] = await window.showOpenFilePicker({
    multiple: false,
    types: [
      {
        description: "ApplyPaste 信息库",
        accept: {
          "application/json": [".json"],
          "text/markdown": [".md"],
          "text/plain": [".txt"]
        }
      }
    ]
  });
  const file = await handle.getFile();
  const profile = parseProfileText(await file.text(), file.name);
  return { profile, handle, fileMeta: { name: file.name, type: detectType(file.name) } };
}

export async function saveProfileToHandle(profile, handle, preferredType = "json") {
  const writable = await handle.createWritable();
  const text = preferredType === "md" ? toMarkdownLibrary(profile) : JSON.stringify(profile, null, 2);
  await writable.write(text);
  await writable.close();
}

export function downloadProfile(profile, type = "json") {
  if (type === "docx") return downloadBlob(createProfileDocx(profile), "applypaste-profile.docx");
  const text = type === "md" ? toMarkdownLibrary(profile) : JSON.stringify(profile, null, 2);
  const blob = new Blob([text], { type: type === "md" ? "text/markdown" : "application/json" });
  downloadBlob(blob, `applypaste-profile.${type}`);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  chrome.downloads.download({
    url,
    filename,
    saveAs: true
  });
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

export function parseProfileText(text, fileName = "") {
  if (detectType(fileName) === "json") {
    const profile = JSON.parse(text);
    if (!Array.isArray(profile.categories) || !Array.isArray(profile.items)) {
      throw new Error("JSON 信息库缺少 categories 或 items");
    }
    return profile;
  }
  return parseMarkdownLibrary(text);
}

export function detectType(fileName) {
  return /\.json$/i.test(fileName) ? "json" : "md";
}
