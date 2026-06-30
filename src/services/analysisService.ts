import { ApiFactory, UnifiedGenerateOptions } from '../api/apiFactory';
import { ApiResponse } from '../types';

export interface AnalysisOptions {
  text: string;
  focusArea: 'general' | 'sentiment' | 'entities' | 'data_cleanup';
  customInstructions?: string;
  stream?: boolean;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

export class AnalysisService {
  private apiFactory: ApiFactory;

  constructor(apiFactory: ApiFactory) {
    this.apiFactory = apiFactory;
  }

  /**
   * Thực hiện phân tích nội dung văn bản hoặc dữ liệu ô Excel
   */
  async analyze(options: AnalysisOptions): Promise<ApiResponse> {
    const focusDescriptions = {
      general: 'Phân tích toàn diện: Tóm tắt ý chính, đưa ra các nhận xét quan trọng và đề xuất tiếp theo.',
      sentiment: 'Phân tích sắc thái (Sentiment Analysis): Xác định thái độ (Tích cực, Tiêu cực, Trung lập) kèm luận điểm chi tiết.',
      entities: 'Trích xuất thực thể (Entity Extraction): Liệt kê các danh từ riêng như tên người, công ty, địa danh, ngày tháng, email, số điện thoại.',
      data_cleanup: 'Đánh giá chất lượng dữ liệu: Tìm lỗi sai chính tả, thông tin thiếu sót, hoặc định dạng không nhất quán và đề xuất làm sạch.'
    };

    const systemPrompt = `Bạn là một Nhà Phân Tích Dữ Liệu cao cấp. Hãy phân tích văn bản/dữ liệu được cung cấp một cách khách quan, khoa học và trình bày kết quả rõ ràng dưới dạng Markdown.`;

    let userPrompt = `Hãy phân tích nội dung sau:
---
Mục tiêu phân tích: ${focusDescriptions[options.focusArea]}
Ngôn ngữ đầu ra: Tiếng Việt (Tiếng Việt chuẩn công sở)
`;

    if (options.customInstructions) {
      userPrompt += `Yêu cầu bổ sung từ người dùng: ${options.customInstructions}\n`;
    }

    userPrompt += `
Nội dung cần phân tích:
"""
${options.text}
"""
---

Constraints:
- Trả về kết quả phân tích có cấu trúc Markdown rõ ràng (sử dụng các thẻ tiêu đề #, ##, danh sách gạch đầu dòng, bảng nếu cần).
- Phân tích ngắn gọn nhưng sâu sắc, tránh viết dài dòng sáo rỗng.
- Nếu dữ liệu không đủ để phân tích sắc thái/thực thể, hãy báo cáo rõ ràng thay vì tự bịa ra thông tin.

Kết quả phân tích:`;

    const generateOptions: UnifiedGenerateOptions = {
      prompt: userPrompt,
      system: systemPrompt,
      stream: options.stream ?? false,
      temperature: 0.2, // Nhiệt độ thấp giúp mô hình bám sát dữ liệu và phân tích logic
      onChunk: options.onChunk,
      signal: options.signal,
    };

    return await this.apiFactory.generate(generateOptions, 'analysis');
  }
}
