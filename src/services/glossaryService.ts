import { GlossaryItem } from '../types';

export class GlossaryService {
  private dbName = 'ExcelAiAddInDb';
  private storeName = 'glossary';
  private dbVersion = 2;

  constructor() {
    if (typeof indexedDB !== 'undefined') {
      this.initDb().catch(err => {
        console.error('Không thể khởi tạo IndexedDB cho Glossary:', err);
      });
    }
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
        if (!db.objectStoreNames.contains('history')) {
          const store = db.createObjectStore('history', { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
        }
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName, { keyPath: 'japanese' });
        }
      };
    });
  }

  /**
   * Thêm hoặc cập nhật một thuật ngữ
   */
  async addGlossaryItem(item: GlossaryItem): Promise<void> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.put(item); // Dùng put để tự động đè nếu trùng key 'japanese'

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Không thể lưu thuật ngữ vào IndexedDB'));
    });
  }

  /**
   * Thêm hàng loạt thuật ngữ (Batch Import)
   */
  async addGlossaryItems(items: GlossaryItem[]): Promise<number> {
    if (items.length === 0) return 0;
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);

      let successCount = 0;
      transaction.oncomplete = () => {
        resolve(successCount);
      };

      transaction.onerror = () => {
        reject(new Error('Lỗi xảy ra trong quá trình import thuật ngữ'));
      };

      items.forEach(item => {
        const request = store.put(item);
        request.onsuccess = () => {
          successCount++;
        };
      });
    });
  }

  /**
   * Xóa một thuật ngữ theo từ gốc tiếng Nhật
   */
  async deleteGlossaryItem(japanese: string): Promise<void> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.delete(japanese);

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Không thể xóa thuật ngữ khỏi IndexedDB'));
    });
  }

  /**
   * Lấy toàn bộ danh sách thuật ngữ
   */
  async getGlossaryItems(): Promise<GlossaryItem[]> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readonly');
      const store = transaction.objectStore(this.storeName);
      const request = store.getAll();

      request.onsuccess = () => {
        resolve(request.result || []);
      };

      request.onerror = () => {
        reject(new Error('Không thể đọc dữ liệu thuật ngữ'));
      };
    });
  }

  /**
   * Xóa toàn bộ thuật ngữ
   */
  async clearGlossary(): Promise<void> {
    const db = await this.initDb();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction([this.storeName], 'readwrite');
      const store = transaction.objectStore(this.storeName);
      const request = store.clear();

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Không thể xóa sạch bảng thuật ngữ'));
    });
  }

  /**
   * Phân tích tệp JSON sang GlossaryItem[]
   */
  parseJSON(content: string): GlossaryItem[] {
    const parsed = JSON.parse(content);
    if (Array.isArray(parsed)) {
      return parsed
        .filter(item => item && typeof item.japanese === 'string' && typeof item.vietnamese === 'string')
        .map(item => ({
          japanese: item.japanese.trim(),
          vietnamese: item.vietnamese.trim(),
          note: item.note ? String(item.note).trim() : undefined
        }));
    } else if (typeof parsed === 'object') {
      return Object.keys(parsed).map(key => ({
        japanese: key.trim(),
        vietnamese: String(parsed[key]).trim()
      }));
    }
    throw new Error('Định dạng JSON không hợp lệ. Phải là mảng [{japanese, vietnamese}] hoặc đối tượng {từ_gốc: dịch_nghĩa}.');
  }

  /**
   * Phân tích tệp CSV sang GlossaryItem[]
   */
  parseCSV(content: string): GlossaryItem[] {
    const lines = content.split(/\r?\n/);
    const items: GlossaryItem[] = [];

    lines.forEach(line => {
      const trimmedLine = line.trim();
      if (!trimmedLine) return;

      // Xác định delimiter là phẩy (,) hoặc chấm phẩy (;)
      const delimiter = trimmedLine.includes(';') ? ';' : ',';
      
      // Split đơn giản nhưng bỏ qua các cột trống
      const parts = trimmedLine.split(delimiter).map(p => p.trim().replace(/^["']|["']$/g, ''));
      if (parts.length >= 2 && parts[0] && parts[1]) {
        // Bỏ qua dòng tiêu đề nếu có
        if (parts[0].toLowerCase() === 'japanese' || parts[0] === 'Từ gốc' || parts[0] === 'tiếng nhật') {
          return;
        }
        items.push({
          japanese: parts[0],
          vietnamese: parts[1],
          note: parts[2] || undefined
        });
      }
    });

    return items;
  }

  /**
   * Phân tích tệp MD (Markdown) sang GlossaryItem[]
   * Hỗ trợ các định dạng:
   * - | Japanese | Vietnamese | (Markdown table)
   * - - **JP**: VN (Markdown list)
   * - - JP | VN
   */
  parseMarkdown(content: string): GlossaryItem[] {
    const lines = content.split(/\r?\n/);
    const items: GlossaryItem[] = [];

    lines.forEach(line => {
      const trimmedLine = line.trim();
      if (!trimmedLine) return;

      // 1. Phân tích bảng: | JP | VN | [Note] |
      if (trimmedLine.startsWith('|')) {
        const parts = trimmedLine.split('|').map(p => p.trim());
        // Bỏ qua dòng rìa ngoài của split
        if (parts.length >= 4) {
          const jp = parts[1];
          const vn = parts[2];
          const note = parts[3];

          // Bỏ qua dòng tiêu đề hoặc dòng gạch ngang ngăn cách (|---|---|)
          if (!jp || jp.startsWith('-') || jp.toLowerCase() === 'japanese' || jp === 'Từ gốc') {
            return;
          }

          items.push({
            japanese: jp,
            vietnamese: vn,
            note: note || undefined
          });
        }
        return;
      }

      // 2. Phân tích list: - **JP**: VN hoặc - JP: VN
      const listMatch = trimmedLine.match(/^[*+-]\s+(?:\*\*(.*?)\*\*|(.*?))\s*[:|-]\s*(.*)$/);
      if (listMatch) {
        const jp = (listMatch[1] || listMatch[2] || '').trim();
        const vnAndNote = (listMatch[3] || '').trim();
        
        if (jp && vnAndNote) {
          // Phân tách Note nếu có dạng "VN (Note)"
          const noteMatch = vnAndNote.match(/^(.*?)\s*\((.*?)\)$/);
          const vn = noteMatch ? noteMatch[1].trim() : vnAndNote;
          const note = noteMatch ? noteMatch[2].trim() : undefined;

          items.push({
            japanese: jp,
            vietnamese: vn,
            note: note
          });
        }
      }
    });

    return items;
  }
}
