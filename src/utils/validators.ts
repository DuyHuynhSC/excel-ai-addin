/**
 * Bộ kiểm tra tính hợp lệ của dữ liệu đầu vào.
 */

// Giới hạn độ dài ký tự tối đa của một ô Excel đầu vào (xấp xỉ 4,000 tokens)
export const MAX_INPUT_CHARS = 16000;

export interface ValidationResult {
  isValid: boolean;
  message?: string;
}

/**
 * Kiểm tra tính hợp lệ của văn bản đầu vào trước khi gửi tới API.
 * @param text Văn bản từ ô Excel
 */
export function validateInputText(text: unknown): ValidationResult {
  if (text === undefined || text === null) {
    return {
      isValid: false,
      message: 'Không có dữ liệu đầu vào (ô trống).'
    };
  }

  const normalized = String(text).trim();
  if (normalized.length === 0) {
    return {
      isValid: false,
      message: 'Nội dung ô trống hoặc chỉ chứa khoảng trắng.'
    };
  }

  if (normalized.length > MAX_INPUT_CHARS) {
    return {
      isValid: false,
      message: `Văn bản quá dài. Giới hạn là ${MAX_INPUT_CHARS} ký tự (~4,000 tokens). Hiện tại có ${normalized.length} ký tự.`
    };
  }

  return { isValid: true };
}

/**
 * Làm sạch văn bản chống Prompt Injection cơ bản.
 * Loại bỏ các chỉ dẫn trực tiếp phá vỡ hệ thống nếu cần.
 * @param text Văn bản đầu vào từ ô
 */
export function sanitizeInput(text: string): string {
  // Loại bỏ các thẻ HTML nguy hiểm và làm sạch
  return text
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    // Chặn một số chuỗi điều khiển cố gắng bypass hệ thống
    .replace(/(system prompt|ignore previous instructions|bỏ qua hướng dẫn)/gi, '[bị chặn]');
}

/**
 * Kiểm tra cấu hình API
 * @param url Địa chỉ API URL
 */
export function validateApiUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:';
  } catch {
    return false;
  }
}
