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

export class TranslationService {
  private apiFactory: ApiFactory;

  constructor(apiFactory: ApiFactory) {
    this.apiFactory = apiFactory;
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

    const systemPrompt = `Bạn là một Dịch giả Chuyên Nghiệp. Nhiệm vụ của bạn là dịch chính xác nội dung được cung cấp từ ngôn ngữ nguồn sang ngôn ngữ đích mà không thêm bất kỳ nhận xét, phân tích hay giải thích nào bên ngoài. Hãy giữ nguyên định dạng của văn bản gốc (như dấu xuống dòng, danh sách, dấu câu).`;

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
