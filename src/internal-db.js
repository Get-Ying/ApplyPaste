const DB_NAME = "ApplyPasteInternal";
const DB_VERSION = 1;
const DATA_STORE = "data";
const ATTACHMENT_STORE = "attachments";

export async function loadInternalState() {
  const db = await openDb();
  try {
    const transaction = db.transaction([DATA_STORE, ATTACHMENT_STORE], "readonly");
    const data = transaction.objectStore(DATA_STORE);
    const attachments = transaction.objectStore(ATTACHMENT_STORE);
    const [profile, versions, records, attachmentItems] = await Promise.all([
      request(data.get("profile")),
      request(data.get("versions")),
      request(data.get("records")),
      request(attachments.getAll())
    ]);
    await transactionDone(transaction);
    return { profile: profile || null, versions: versions || null, records: records || [], attachments: attachmentItems || [] };
  } finally {
    db.close();
  }
}

export async function saveCoreData(profile, versions) {
  const db = await openDb();
  try {
    const transaction = db.transaction(DATA_STORE, "readwrite");
    const store = transaction.objectStore(DATA_STORE);
    store.put(profile, "profile");
    store.put(versions, "versions");
    await transactionDone(transaction);
  } finally {
    db.close();
  }
}

export async function saveRecords(records) {
  const db = await openDb();
  try {
    const transaction = db.transaction(DATA_STORE, "readwrite");
    transaction.objectStore(DATA_STORE).put(records, "records");
    await transactionDone(transaction);
  } finally {
    db.close();
  }
}

export async function saveAttachments(attachments) {
  const db = await openDb();
  try {
    const transaction = db.transaction(ATTACHMENT_STORE, "readwrite");
    const store = transaction.objectStore(ATTACHMENT_STORE);
    store.clear();
    for (const attachment of attachments) store.put(attachment);
    await transactionDone(transaction);
  } finally {
    db.close();
  }
}

export async function saveCompleteState({ profile, versions, records, attachments }) {
  const db = await openDb();
  try {
    const transaction = db.transaction([DATA_STORE, ATTACHMENT_STORE], "readwrite");
    const data = transaction.objectStore(DATA_STORE);
    const files = transaction.objectStore(ATTACHMENT_STORE);
    data.put(profile, "profile");
    data.put(versions, "versions");
    data.put(records, "records");
    files.clear();
    for (const attachment of attachments) files.put(attachment);
    await transactionDone(transaction);
  } finally {
    db.close();
  }
}

export function storageErrorMessage(error) {
  if (error?.name === "QuotaExceededError") return "浏览器存储空间不足，请删除不需要的附件或先导出完整备份。";
  return error?.message || "内部存储失败";
}

function openDb() {
  return new Promise((resolve, reject) => {
    const dbRequest = indexedDB.open(DB_NAME, DB_VERSION);
    dbRequest.onupgradeneeded = () => {
      const db = dbRequest.result;
      if (!db.objectStoreNames.contains(DATA_STORE)) db.createObjectStore(DATA_STORE);
      if (!db.objectStoreNames.contains(ATTACHMENT_STORE)) db.createObjectStore(ATTACHMENT_STORE, { keyPath: "id" });
    };
    dbRequest.onsuccess = () => resolve(dbRequest.result);
    dbRequest.onerror = () => reject(dbRequest.error);
  });
}

function request(value) {
  return new Promise((resolve, reject) => {
    value.onsuccess = () => resolve(value.result);
    value.onerror = () => reject(value.error);
  });
}

function transactionDone(transaction) {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error);
    transaction.onabort = () => reject(transaction.error || new Error("IndexedDB transaction aborted"));
  });
}
