import { ApiFactory, UnifiedGenerateOptions } from '../api/apiFactory';
import { ApiResponse } from '../types';

export interface SummarizationOptions {
  text: string;
  format: 'paragraph' | 'bullets' | 'key_actions';
  length: 'concise' | 'detailed';
  stream?: boolean;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

export class SummarizationService {
  private apiFactory: ApiFactory;

  constructor(apiFactory: ApiFactory) {
    this.apiFactory = apiFactory;
  }

  /**
   * Thực hiện tóm tắt văn bản
   */
  async summarize(options: SummarizationOptions): Promise<ApiResponse> {
    const formatText = options.format === 'bullets' 
      ? 'Danh sách các gạch đầu dòng (bullet points) tóm tắt các ý chính'
      : options.format === 'key_actions' 
      ? 'Danh sách các hành động/đầu việc cần thực hiện tiếp theo (Action items)' 
      : 'Một đoạn văn ngắn gọn, mạch lạc';

    const lengthText = options.length === 'concise' ? 'Cực kỳ ngắn gọn (không quá 150 từ)' : 'Chi tiết, đầy đủ ngữ cảnh (300-500 từ)';

    const systemPrompt = `Bạn là một Nhà Viết Tóm Tắt Chuyên Nghiệp. Nhiệm vụ của bạn là chắt lọc nội dung cốt lõi của văn bản gốc và trình bày lại theo yêu cầu mà không làm mất đi các thông tin quan trọng nhất.`;

    const userPrompt = `Hãy tóm tắt đoạn văn bản sau:
---
Định dạng yêu cầu: ${formatText}
Độ dài mong muốn: ${lengthText}
Ngôn ngữ đầu ra: Tiếng Việt

Nội dung cần tóm tắt:
"""
${options.text}
"""
---

Constraints:
- Chỉ trả về phần tóm tắt trực tiếp, không có lời giới thiệu ("Dưới đây là tóm tắt...").
- Bám sát sự thật có trong văn bản, tuyệt đối không suy diễn hoặc thêm thông tin mới không có trong bản gốc.
- Đảm bảo cấu trúc câu dễ đọc, mạch lạc.

Bản tóm tắt:`;

    const generateOptions: UnifiedGenerateOptions = {
      prompt: userPrompt,
      system: systemPrompt,
      stream: options.stream ?? false,
      temperature: 0.4, // Cân bằng giữa bám sát thực tế và cách diễn đạt lưu loát
      onChunk: options.onChunk,
      signal: options.signal,
    };

    return await this.apiFactory.generate(generateOptions, 'summarization');
  }
}
