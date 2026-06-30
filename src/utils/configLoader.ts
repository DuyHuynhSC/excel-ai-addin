import { AppConfig } from '../types';

const DEFAULT_CONFIG: AppConfig = {
  ollamaUrl: import.meta.env.VITE_OLLAMA_API_URL || 'http://localhost:11434',
  customApiUrl: import.meta.env.VITE_CUSTOM_API_URL || 'https://api.company.com/v1',
  customApiKey: import.meta.env.VITE_CUSTOM_API_KEY || '',
  activeService: 'ollama',
  fallbackEnabled: true,
  models: {
    translation: import.meta.env.VITE_MODEL_TRANSLATION || 'phi:2b',
    analysis: import.meta.env.VITE_MODEL_ANALYSIS || 'llama2:7b-chat',
    summarization: import.meta.env.VITE_MODEL_SUMMARIZATION || 'mistral:7b',
    custom: 'mistral:7b'
  }
};

/**
 * Tải cấu hình từ localStorage. Nếu chưa có, trả về cấu hình mặc định.
 */
export function loadConfig(): AppConfig {
  try {
    const savedConfigStr = localStorage.getItem('excel_ai_addin_config');
    if (savedConfigStr) {
      const parsed = JSON.parse(savedConfigStr);
      return {
        ...DEFAULT_CONFIG,
        ...parsed,
        models: {
          ...DEFAULT_CONFIG.models,
          ...(parsed.models || {})
        }
      };
    }
  } catch (e) {
    console.error('Không thể đọc cấu hình từ localStorage:', e);
  }
  return DEFAULT_CONFIG;
}

/**
 * Lưu cấu hình vào localStorage.
 */
export function saveConfig(config: AppConfig): void {
  try {
    localStorage.setItem('excel_ai_addin_config', JSON.stringify(config));
  } catch (e) {
    console.error('Không thể lưu cấu hình vào localStorage:', e);
  }
}
