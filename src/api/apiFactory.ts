import { OllamaClient } from './ollamaClient';
import { CustomApiClient } from './customApiClient';
import { AppConfig, ApiResponse, TaskType } from '../types';
import { decryptText } from '../utils/cryptoUtils';

export interface UnifiedGenerateOptions {
  prompt: string;
  system?: string;
  stream?: boolean;
  temperature?: number;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

export class ApiFactory {
  private config: AppConfig;
  private ollamaClient: OllamaClient;
  private customApiClient: CustomApiClient | null = null;
  private customClientInitialized = false;

  constructor(config: AppConfig) {
    this.config = config;
    this.ollamaClient = new OllamaClient(config.ollamaUrl);
  }

  private async getCustomClient(): Promise<CustomApiClient | null> {
    if (!this.customClientInitialized) {
      try {
        const decryptedKey = await decryptText(this.config.customApiKey);
        if (decryptedKey) {
          this.customApiClient = new CustomApiClient(this.config.customApiUrl, decryptedKey);
        } else {
          this.customApiClient = null;
        }
      } catch (e) {
        console.error('Không thể giải mã Custom API Key:', e);
        this.customApiClient = null;
      }
      this.customClientInitialized = true;
    }
    return this.customApiClient;
  }

  /**
   * Cập nhật cấu hình runtime mới
   */
  async updateConfig(newConfig: AppConfig) {
    this.config = newConfig;
    this.ollamaClient = new OllamaClient(newConfig.ollamaUrl);
    this.customApiClient = null;
    this.customClientInitialized = false;
  }

  /**
   * Gọi sinh văn bản dựa trên dịch vụ active, hỗ trợ thử lại và tự động chuyển đổi dự phòng (fallback)
   */
  async generate(
    options: UnifiedGenerateOptions,
    taskType: TaskType
  ): Promise<ApiResponse> {
    const activeService = this.config.activeService;
    const fallbackEnabled = this.config.fallbackEnabled;

    try {
      if (activeService === 'custom') {
        return await this.executeWithRetry(() => this.callCustomApi(options, taskType));
      } else {
        return await this.executeWithRetry(() => this.callOllama(options, taskType));
      }
    } catch (error) {
      console.warn(`Lỗi khi thực hiện dịch vụ chính (${activeService}):`, error);

      if (fallbackEnabled) {
        const secondaryService = activeService === 'custom' ? 'ollama' : 'custom';
        console.log(`Đang tự động chuyển sang dịch vụ dự phòng: ${secondaryService}`);

        if (options.onChunk) {
          options.onChunk('\n\n[🔄 Hệ thống: Đang tự động chuyển sang dịch vụ dự phòng...]\n\n');
        }

        try {
          if (secondaryService === 'custom') {
            return await this.executeWithRetry(() => this.callCustomApi(options, taskType));
          } else {
            return await this.executeWithRetry(() => this.callOllama(options, taskType));
          }
        } catch (fallbackError) {
          console.error('Cả hai dịch vụ chính và dự phòng đều thất bại:', fallbackError);
          throw new Error(`Tất cả các dịch vụ AI đều không khả dụng. Chi tiết lỗi gốc: ${error instanceof Error ? error.message : String(error)}`);
        }
      }

      throw error;
    }
  }

  /**
   * Thực hiện gọi Ollama Client
   */
  private async callOllama(options: UnifiedGenerateOptions, taskType: TaskType): Promise<ApiResponse> {
    const model = this.config.models[taskType] || this.config.models.custom;
    return await this.ollamaClient.generate(
      {
        model,
        prompt: options.prompt,
        system: options.system,
        stream: options.stream,
        temperature: options.temperature,
        onChunk: options.onChunk,
        signal: options.signal,
      },
      taskType
    );
  }

  /**
   * Thực hiện gọi Custom API Client
   */
  private async callCustomApi(options: UnifiedGenerateOptions, taskType: TaskType): Promise<ApiResponse> {
    const client = await this.getCustomClient();
    if (!client) {
      throw new Error('Dịch vụ Custom API chưa được cấu hình API Key hoặc lỗi giải mã khóa.');
    }
    const model = this.config.models[taskType] || this.config.models.custom;
    return await client.generate(
      {
        model,
        prompt: options.prompt,
        system: options.system,
        stream: options.stream,
        temperature: options.temperature,
        onChunk: options.onChunk,
        signal: options.signal,
      },
      taskType
    );
  }

  /**
   * Thực thi hàm với cơ chế Exponential Backoff
   * Thử lại tối đa 3 lần: lần 1 chờ 1s, lần 2 chờ 2s, lần 3 chờ 4s
   */
  private async executeWithRetry<T>(fn: () => Promise<T>, attempts = 3, delay = 1000): Promise<T> {
    try {
      // Thiết lập timeout 120s cho mỗi request
      const timeoutPromise = new Promise<never>((_, reject) =>
        setTimeout(() => reject(new Error('TIMEOUT_ERROR')), 120000)
      );
      return await Promise.race([fn(), timeoutPromise]);
    } catch (error) {
      if (error instanceof Error && error.message === 'TIMEOUT_ERROR') {
        throw new Error('Yêu cầu hết thời gian xử lý (quá 120 giây).');
      }

      if (attempts <= 1) {
        throw error;
      }
      console.warn(`Lỗi xảy ra, đang thử lại lần tiếp theo sau ${delay}ms...`, error);
      await new Promise(resolve => setTimeout(resolve, delay));
      return this.executeWithRetry(fn, attempts - 1, delay * 2);
    }
  }

  /**
   * Lấy danh sách model có sẵn của service hoạt động hiện tại
   */
  async getActiveModels(): Promise<string[]> {
    if (this.config.activeService === 'custom') {
      const client = await this.getCustomClient();
      if (client) {
        return await client.getModels();
      }
    }
    return await this.ollamaClient.getModels();
  }

  /**
   * Kiểm tra tình trạng kết nối của cả hai dịch vụ
   */
  async checkHealth(): Promise<{ ollama: boolean; custom: boolean }> {
    const ollamaPromise = this.ollamaClient.testConnection();
    
    const client = await this.getCustomClient();
    const customPromise = client
      ? client.testConnection()
      : Promise.resolve(false);

    // Timeout kết nối là 15 giây
    const timeoutPromise = (serviceName: string) =>
      new Promise<boolean>(resolve =>
        setTimeout(() => {
          console.warn(`Kiểm tra kết nối ${serviceName} bị quá thời gian (15 giây).`);
          resolve(false);
        }, 15000)
      );

    const [ollamaResult, customResult] = await Promise.all([
      Promise.race([ollamaPromise, timeoutPromise('Ollama')]),
      Promise.race([customPromise, timeoutPromise('Custom API')])
    ]);

    return {
      ollama: ollamaResult,
      custom: customResult
    };
  }
}
