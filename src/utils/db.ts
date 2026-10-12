const DB_NAME = '179_Notice_Generator_DB';
const DB_VERSION = 1;
const STORE_NAME = 'master_template';
const TEMPLATE_KEY = '179_bnss_ncmec_notice_master';
const CDR_TEMPLATE_KEY = 'ncmec_cdr_caf_sdr_master';
const IO_DETAILS_KEY = '179_bnss_ncmec_io_details';

export interface StoredTemplate {
  id: string;
  fileName: string;
  updatedAt: number;
  buffer: ArrayBuffer;
}

export interface StoredIoDetails {
  id: string;
  ioName: string;
  ioMobile: string;
  updatedAt: number;
}

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function saveMasterTemplate(file: File): Promise<StoredTemplate> {
  const buffer = await file.arrayBuffer();
  const db = await openDB();
  const record: StoredTemplate = {
    id: TEMPLATE_KEY,
    fileName: file.name,
    updatedAt: Date.now(),
    buffer,
  };

  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(record);

    req.onsuccess = () => resolve(record);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

export async function getMasterTemplate(): Promise<StoredTemplate | null> {
  const db = await openDB();

  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get(TEMPLATE_KEY);

    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

export async function saveCdrMasterTemplate(file: File): Promise<StoredTemplate> {
  const buffer = await file.arrayBuffer();
  const db = await openDB();
  const record: StoredTemplate = {
    id: CDR_TEMPLATE_KEY,
    fileName: file.name,
    updatedAt: Date.now(),
    buffer,
  };

  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(record);

    req.onsuccess = () => resolve(record);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

export async function getCdrMasterTemplate(): Promise<StoredTemplate | null> {
  const db = await openDB();

  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get(CDR_TEMPLATE_KEY);

    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

export async function saveIoDetails(ioName: string, ioMobile: string): Promise<StoredIoDetails> {
  const db = await openDB();
  const record: StoredIoDetails = {
    id: IO_DETAILS_KEY,
    ioName,
    ioMobile,
    updatedAt: Date.now(),
  };

  try {
    localStorage.setItem(IO_DETAILS_KEY, JSON.stringify({ ioName, ioMobile }));
  } catch {
    // Ignore localStorage quota/privacy errors
  }

  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const req = store.put(record);

    req.onsuccess = () => resolve(record);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

export async function getIoDetails(): Promise<{ ioName: string; ioMobile: string }> {
  try {
    const db = await openDB();
    const result = await new Promise<StoredIoDetails | null>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const req = store.get(IO_DETAILS_KEY);

      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
      tx.oncomplete = () => db.close();
    });

    if (result) {
      return { ioName: result.ioName || '', ioMobile: result.ioMobile || '' };
    }
  } catch {
    // Fallback to localStorage if IndexedDB fails
  }

  try {
    const raw = localStorage.getItem(IO_DETAILS_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return { ioName: parsed.ioName || '', ioMobile: parsed.ioMobile || '' };
    }
  } catch {
    // Ignore parse errors
  }

  return { ioName: '', ioMobile: '' };
}
