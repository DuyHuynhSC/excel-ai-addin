/* global CustomFunctions */
declare const CustomFunctions: any;
import { ApiFactory } from './api/apiFactory';
import { loadConfig } from './utils/configLoader';
import { TranslationService } from './services/translationService';
import { SummarizationService } from './services/summarizationService';
import { AnalysisService } from './services/analysisService';
import { validateInputText } from './utils/validators';

// Khởi tạo ApiFactory dùng chung khi chạy custom functions
let apiFactoryInstance: ApiFactory | null = null;

function getApiFactory(): ApiFactory {
  const config = loadConfig();
  if (!apiFactoryInstance) {
    apiFactoryInstance = new ApiFactory(config);
  } else {
    // Cập nhật cấu hình runtime mới nhất từ LocalStorage
    apiFactoryInstance.updateConfig(config);
  }
  return apiFactoryInstance;
}

/**
 * Dịch một ô dữ liệu văn bản sang ngôn ngữ mong muốn trực tiếp trong ô Excel.
 * @customfunction TRANSLATE
 * @param text Nội dung văn bản cần dịch.
 * @param targetLang Ngôn ngữ cần dịch sang (ví dụ: "Tiếng Anh", "Tiếng Việt").
 * @param sourceLang Ngôn ngữ của văn bản gốc (mặc định: "Auto").
 * @returns Bản dịch của văn bản.
 */
export async function translate(
  text: string,
  targetLang: string,
  sourceLang: string = 'Auto'
): Promise<string> {
  const validation = validateInputText(text);
  if (!validation.isValid) {
    return `❌ ${validation.message}`;
  }

  try {
    const factory = getApiFactory();
    const service = new TranslationService(factory);
    const response = await service.translate({
      text: String(text),
      sourceLang,
      targetLang,
      tone: 'neutral',
      stream: false
    });
    return response.result;
  } catch (error) {
    return `❌ Lỗi: ${error instanceof Error ? error.message : String(error)}`;
  }
}

/**
 * Tóm tắt văn bản trực tiếp trong ô Excel.
 * @customfunction SUMMARIZE
 * @param text Nội dung văn bản cần tóm tắt.
 * @param format Định dạng đầu ra: "paragraph" (mặc định) hoặc "bullets" (gạch đầu dòng).
 * @param length Độ dài tóm tắt: "concise" (mặc định) hoặc "detailed".
 * @returns Bản tóm tắt của văn bản.
 */
export async function summarize(
  text: string,
  format: 'paragraph' | 'bullets' = 'paragraph',
  length: 'concise' | 'detailed' = 'concise'
): Promise<string> {
  const validation = validateInputText(text);
  if (!validation.isValid) {
    return `❌ ${validation.message}`;
  }

  try {
    const factory = getApiFactory();
    const service = new SummarizationService(factory);
    const response = await service.summarize({
      text: String(text),
      format,
      length,
      stream: false
    });
    return response.result;
  } catch (error) {
    return `❌ Lỗi: ${error instanceof Error ? error.message : String(error)}`;
  }
}

/**
 * Phân tích nội dung ô dữ liệu Excel.
 * @customfunction ANALYZE
 * @param text Nội dung cần phân tích.
 * @param focusArea Lĩnh vực phân tích: "general" (tổng quan), "sentiment" (sắc thái), hoặc "entities" (thực thể).
 * @returns Kết quả phân tích chi tiết.
 */
export async function analyze(
  text: string,
  focusArea: 'general' | 'sentiment' | 'entities' = 'general'
): Promise<string> {
  const validation = validateInputText(text);
  if (!validation.isValid) {
    return `❌ ${validation.message}`;
  }

  try {
    const factory = getApiFactory();
    const service = new AnalysisService(factory);
    const response = await service.analyze({
      text: String(text),
      focusArea,
      stream: false
    });
    return response.result;
  } catch (error) {
    return `❌ Lỗi: ${error instanceof Error ? error.message : String(error)}`;
  }
}

/**
 * Thực thi một câu lệnh/yêu cầu tự chọn của bạn lên ô dữ liệu được cung cấp.
 * @customfunction ASK
 * @param prompt Câu hỏi hoặc yêu cầu xử lý (Ví dụ: "Viết lại bằng thơ").
 * @param contextText Nội dung ô dữ liệu làm ngữ cảnh.
 * @returns Câu trả lời từ AI.
 */
export async function ask(prompt: string, contextText: string): Promise<string> {
  const promptValidation = validateInputText(prompt);
  if (!promptValidation.isValid) {
    return `❌ Hướng dẫn: ${promptValidation.message}`;
  }

  const contextValidation = validateInputText(contextText);
  if (!contextValidation.isValid) {
    return `❌ Ngữ cảnh: ${contextValidation.message}`;
  }

  try {
    const factory = getApiFactory();
    const fullPrompt = `Hãy thực hiện yêu cầu sau dựa trên dữ liệu ngữ cảnh được cung cấp.
Yêu cầu: ${prompt}

Dữ liệu ngữ cảnh:
"""
${contextText}
"""`;

    const response = await factory.generate({
      prompt: fullPrompt,
      system: 'Bạn là một Trợ lý AI đắc lực trong bảng tính Excel. Hãy trả lời ngắn gọn, trực diện và chính xác vào câu hỏi.',
      stream: false,
      temperature: 0.5
    }, 'custom');

    return response.result;
  } catch (error) {
    return `❌ Lỗi: ${error instanceof Error ? error.message : String(error)}`;
  }
}

// Đăng ký các hàm này vào danh sách custom functions của Office
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('TRANSLATE', translate);
  CustomFunctions.associate('SUMMARIZE', summarize);
  CustomFunctions.associate('ANALYZE', analyze);
  CustomFunctions.associate('ASK', ask);
}
