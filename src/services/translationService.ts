import { ApiFactory, UnifiedGenerateOptions } from '../api/apiFactory';
import { ApiResponse } from '../types';

export interface TranslationOptions {
  text: string;
  sourceLang: string; // ví dụ: "Auto" hoặc "Tiếng Anh", "Tiếng Việt"
  targetLang: string; // ví dụ: "Tiếng Việt", "Tiếng Anh"
  tone?: 'formal' | 'informal' | 'neutral';
  stream?: boolean;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

import { GlossaryService } from './glossaryService';

export class TranslationService {
  private apiFactory: ApiFactory;
  private glossaryService: GlossaryService;

  constructor(apiFactory: ApiFactory) {
    this.apiFactory = apiFactory;
    this.glossaryService = new GlossaryService();
  }

  /**
   * Thực hiện dịch văn bản qua LLM
   */
  async translate(options: TranslationOptions): Promise<ApiResponse> {
    const toneText = options.tone === 'formal' 
      ? 'Trang trọng, lịch sự' 
      : options.tone === 'informal' 
      ? 'Thân mật, tự nhiên' 
      : 'Trung lập';

    // Nạp và so khớp các thuật ngữ Glossary chuyên ngành (chỉ áp dụng cho tiếng Nhật - tiếng Việt)
    let matchingGlossaryText = '';
    try {
      const isJaToVi = (options.sourceLang === 'Tiếng Nhật' || options.sourceLang === 'Tự động phát hiện') && options.targetLang === 'Tiếng Việt';
      const isViToJa = options.sourceLang === 'Tiếng Việt' && options.targetLang === 'Tiếng Nhật';
      
      if (isJaToVi || isViToJa) {
        const glossaryItems = await this.glossaryService.getGlossaryItems();
        const matches = glossaryItems.filter(item => 
          options.text.includes(item.japanese) || options.text.includes(item.vietnamese)
        );

        if (matches.length > 0) {
          matchingGlossaryText = '\n\nDưới đây là danh sách thuật ngữ chuyên ngành/chuẩn công ty BẮT BUỘC tuân thủ (Glossary):\n' +
            matches.map(item => `- "${item.japanese}" tương đương "${item.vietnamese}"${item.note ? ` (Ngữ cảnh: ${item.note})` : ''}`).join('\n') +
            '\n\nHãy bắt buộc dịch các thuật ngữ này đúng chính xác như trên, không tự ý thay đổi hay biến tấu.';
        }
      }
    } catch (err) {
      console.warn('Không thể nạp thuật ngữ cho dịch thuật:', err);
    }

    const systemPrompt = `Bạn là một Dịch giả Chuyên Nghiệp. Nhiệm vụ của bạn là dịch chính xác nội dung được cung cấp từ ngôn ngữ nguồn sang ngôn ngữ đích mà không thêm bất kỳ nhận xét, phân tích hay giải thích nào bên ngoài. Hãy giữ nguyên định dạng của văn bản gốc (như dấu xuống dòng, danh sách, dấu câu).${matchingGlossaryText}`;

    const userPrompt = `Dịch đoạn văn bản sau:
---
Ngôn ngữ nguồn: ${options.sourceLang}
Ngôn ngữ đích: ${options.targetLang}
Giọng văn (Tone): ${toneText}

Nội dung cần dịch:
"""
${options.text}
"""
---

Constraints:
- Chỉ trả về bản dịch hoàn chỉnh duy nhất.
- Không thêm câu dẫn như "Bản dịch là:", "Đây là bản dịch:".
- Không tự ý thêm giải thích hay chú thích chân trang.
- Đảm bảo bản dịch mượt mà và tự nhiên đối với ngôn ngữ đích.

Bản dịch:`;

    const generateOptions: UnifiedGenerateOptions = {
      prompt: userPrompt,
      system: systemPrompt,
      stream: options.stream ?? false,
      temperature: 0.3, // Nhiệt độ thấp cho kết quả dịch chuẩn xác, ít sáng tạo linh tinh
      onChunk: options.onChunk,
      signal: options.signal,
    };

    return await this.apiFactory.generate(generateOptions, 'translation');
  }
}
