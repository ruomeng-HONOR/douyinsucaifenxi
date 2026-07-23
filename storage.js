const DB_NAME = "video-material-insights-local";
const DB_VERSION = 2;
const STORE = "daily-metrics";
const MONTHLY_STORE = "monthly-materials";
let databasePromise;

function openDatabase() {
  if (databasePromise) return databasePromise;
  databasePromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      const daily = db.objectStoreNames.contains(STORE)
        ? request.transaction.objectStore(STORE)
        : db.createObjectStore(STORE, { keyPath: "key" });
      if (!daily.indexNames.contains("materialKey")) daily.createIndex("materialKey", "materialKey", { unique: false });
      if (!db.objectStoreNames.contains(MONTHLY_STORE)) db.createObjectStore(MONTHLY_STORE, { keyPath: "key" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => {
      databasePromise = null;
      reject(request.error);
    };
  });
  return databasePromise;
}

export async function readRows() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

async function saveToStore(storeName, rows, onProgress) {
  const db = await openDatabase();
  const batchSize = 5000;
  for (let offset = 0; offset < rows.length; offset += batchSize) {
    await new Promise((resolve, reject) => {
      const transaction = db.transaction(storeName, "readwrite");
      const store = transaction.objectStore(storeName);
      rows.slice(offset, offset + batchSize).forEach((row) => store.put(row));
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error || new Error("本地数据写入被中止"));
    });
    onProgress?.(Math.min(offset + batchSize, rows.length), rows.length);
  }
  return rows.length;
}

export async function saveRows(rows, onProgress) {
  return saveToStore(STORE, rows, onProgress);
}

export async function readMonthlyRows() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = db.transaction(MONTHLY_STORE, "readonly").objectStore(MONTHLY_STORE).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

export async function saveMonthlyRows(rows, onProgress) {
  return saveToStore(MONTHLY_STORE, rows, onProgress);
}

export async function readMaterialRows(material, start, end) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = db.transaction(STORE, "readonly").objectStore(STORE).index("materialKey").getAll(material.materialKey);
    request.onsuccess = () => resolve((request.result || []).filter((row) => row.date >= start && row.date <= end));
    request.onerror = () => reject(request.error);
  });
}

export async function clearRows() {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([STORE, MONTHLY_STORE], "readwrite");
    transaction.objectStore(STORE).clear();
    transaction.objectStore(MONTHLY_STORE).clear();
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error);
  });
}
