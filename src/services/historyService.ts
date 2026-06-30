import { HistoryItem } from '../types';

export class HistoryService {
  private dbName = 'ExcelAiAddInDb';
  private storeName = 'history';
  private dbVersion = 1;

  constructor() {
    this.initDb().catch(err => {
      console.error('Không thể khởi tạo IndexedDB cho Lịch sử:', err);
    });
  }

  /**
   * Khởi tạo cơ sở dữ liệu IndexedDB
   */
  private initDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.dbVersion);

      request.onerror = () => {
        reject(new Error('Không thể mở IndexedDB'));
      };

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onupgradeneeded = (_event) => {
        const db = request.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          const store = db.createObjectStore(this.storeName, { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
        }
      };
    });
  }

  /**
   * Thêm một mục lịch sử mới vào database
   */
  async addHistoryItem(item: Omit<HistoryItem, 'id' | 'timestamp'>): Promise<HistoryItem> {
    const db = await this.initDb();
    const historyItem: HistoryItem = {
      ...item,
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
    };

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.add(historyItem);

      request.onsuccess = () => {
        resolve(historyItem);
      };

      request.onerror = () => {
        reject(new Error('Không thể thêm lịch sử vào IndexedDB'));
      };
    });
  }

  /**
   * Lấy toàn bộ danh sách lịch sử, sắp xếp theo thời gian mới nhất trước
   */
  async getAllHistory(): Promise<HistoryItem[]> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readonly');
      const store = transaction.objectStore(this.storeName);
      const request = store.getAll();

      request.onsuccess = () => {
        const results = request.result as HistoryItem[];
        // Sắp xếp giảm dần theo thời gian (mới nhất lên đầu)
        results.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
        resolve(results);
      };

      request.onerror = () => {
        reject(new Error('Không thể truy vấn lịch sử từ IndexedDB'));
      };
    });
  }

  /**
   * Xóa toàn bộ lịch sử
   */
  async clearHistory(): Promise<void> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.clear();

      request.onsuccess = () => {
        resolve();
      };

      request.onerror = () => {
        reject(new Error('Không thể xóa sạch IndexedDB'));
      };
    });
  }
}
