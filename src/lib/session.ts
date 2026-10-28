import type { ColumnMapping } from './matching';
import type { CellEdit } from './spreadsheet';

export interface SavedSession {
  version: 1;
  savedAt: string;
  workbook: { name: string; bytes: Uint8Array } | null;
  images: File[];
  sheetId: string;
  mapping: ColumnMapping;
  confirmed: boolean;
  edits: CellEdit[];
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('graphic-archive-workspace', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('sessions');
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
  });
}

export async function saveSession(session: SavedSession): Promise<void> {
  const db = await openDatabase();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('sessions', 'readwrite');
      tx.objectStore('sessions').put(session, 'current');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error ?? new Error('저장이 중단되었습니다.'));
    });
  } finally { db.close(); }
}

export async function readSession(): Promise<SavedSession | undefined> {
  const db = await openDatabase();
  try {
    return await new Promise<SavedSession | undefined>((resolve, reject) => {
      const request = db.transaction('sessions', 'readonly').objectStore('sessions').get('current');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  } finally { db.close(); }
}
