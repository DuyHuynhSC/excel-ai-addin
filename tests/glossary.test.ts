import { describe, it, expect } from 'vitest';
import { GlossaryService } from '../src/services/glossaryService';

describe('GlossaryService File Parsing Tests', () => {
  const glossaryService = new GlossaryService();

  describe('parseJSON', () => {
    it('should parse an array of glossary items correctly', () => {
      const jsonContent = JSON.stringify([
        { japanese: '画面カスタマイズ', vietnamese: 'Tùy biến màn hình', note: 'Giao diện' },
        { japanese: '出力フォルダ', vietnamese: 'Thư mục đầu ra' }
      ]);
      const result = glossaryService.parseJSON(jsonContent);
      expect(result).toHaveLength(2);
      expect(result[0].japanese).toBe('画面カスタマイズ');
      expect(result[0].vietnamese).toBe('Tùy biến màn hình');
      expect(result[0].note).toBe('Giao diện');
      expect(result[1].japanese).toBe('出力フォルダ');
      expect(result[1].vietnamese).toBe('Thư mục đầu ra');
      expect(result[1].note).toBeUndefined();
    });

    it('should parse a key-value object correctly', () => {
      const jsonContent = JSON.stringify({
        '日本語': 'Tiếng Nhật',
        'ベトナム語': 'Tiếng Việt'
      });
      const result = glossaryService.parseJSON(jsonContent);
      expect(result).toHaveLength(2);
      expect(result[0].japanese).toBe('日本語');
      expect(result[0].vietnamese).toBe('Tiếng Nhật');
      expect(result[1].japanese).toBe('ベトナム語');
      expect(result[1].vietnamese).toBe('Tiếng Việt');
    });
  });

  describe('parseCSV', () => {
    it('should parse comma-separated values correctly', () => {
      const csvContent = `japanese,vietnamese,note\n画面,Màn hình,UI\n出力,Đầu ra`;
      const result = glossaryService.parseCSV(csvContent);
      expect(result).toHaveLength(2);
      expect(result[0].japanese).toBe('画面');
      expect(result[0].vietnamese).toBe('Màn hình');
      expect(result[0].note).toBe('UI');
      expect(result[1].japanese).toBe('出力');
      expect(result[1].vietnamese).toBe('Đầu ra');
      expect(result[1].note).toBeUndefined();
    });

    it('should parse semicolon-separated values correctly', () => {
      const csvContent = `Từ gốc;Dịch nghĩa\n入力;Đầu vào\n保存;Lưu`;
      const result = glossaryService.parseCSV(csvContent);
      expect(result).toHaveLength(2);
      expect(result[0].japanese).toBe('入力');
      expect(result[0].vietnamese).toBe('Đầu vào');
      expect(result[1].japanese).toBe('保存');
      expect(result[1].vietnamese).toBe('Lưu');
    });
  });

  describe('parseMarkdown', () => {
    it('should parse markdown tables correctly', () => {
      const mdContent = `
| Japanese | Vietnamese | Note |
|---|---|---|
| 画面カスタマイズ | Tùy biến màn hình | Giao diện |
| 出力フォルダ | Thư mục đầu ra | |
`;
      const result = glossaryService.parseMarkdown(mdContent);
      expect(result).toHaveLength(2);
      expect(result[0].japanese).toBe('画面カスタマイズ');
      expect(result[0].vietnamese).toBe('Tùy biến màn hình');
      expect(result[0].note).toBe('Giao diện');
      expect(result[1].japanese).toBe('出力フォルダ');
      expect(result[1].vietnamese).toBe('Thư mục đầu ra');
    });

    it('should parse markdown list format correctly', () => {
      const mdContent = `
* **日本語**: Tiếng Nhật (Ngôn ngữ)
- **ベトナム語**: Tiếng Việt
* 画面: Màn hình (Giao diện UI)
`;
      const result = glossaryService.parseMarkdown(mdContent);
      expect(result).toHaveLength(3);
      expect(result[0].japanese).toBe('日本語');
      expect(result[0].vietnamese).toBe('Tiếng Nhật');
      expect(result[0].note).toBe('Ngôn ngữ');

      expect(result[1].japanese).toBe('ベトナム語');
      expect(result[1].vietnamese).toBe('Tiếng Việt');
      expect(result[1].note).toBeUndefined();

      expect(result[2].japanese).toBe('画面');
      expect(result[2].vietnamese).toBe('Màn hình');
      expect(result[2].note).toBe('Giao diện UI');
    });
  });
});
