import { ApiResponse, TaskType } from '../types';

export interface CustomApiGenerateOptions {
  model: string;
  prompt: string;
  system?: string;
  stream?: boolean;
  temperature?: number;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

export class CustomApiClient {
  private baseUrl: string;
  private apiKey: string;
  private proxyDisabled: boolean;

  constructor(baseUrl: string, apiKey: string, proxyDisabled: boolean = false) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
    this.apiKey = apiKey;
    this.proxyDisabled = proxyDisabled;
  }

  /**
   * Định tuyến URL để vượt qua lỗi CORS bằng cách sử dụng Proxy trung gian khi gọi server khác Origin
   */
  private getRequestUrl(endpointPath: string): string {
    const cleanPath = endpointPath.replace(/^\//, '');
    const currentOrigin = typeof window !== 'undefined' ? window.location.origin : '';

    // Nếu URL đích trỏ tới một máy chủ ngoài khác origin của giao diện Add-in,
    // tự động định tuyến thông qua proxy /api-proxy (hỗ trợ cả Local Vite và Vercel Serverless)
    if (!this.proxyDisabled && currentOrigin && !this.baseUrl.startsWith(currentOrigin)) {
      return `${currentOrigin}/api-proxy/${cleanPath}`;
    }
    
    return `${this.baseUrl}/${cleanPath}`;
  }

  /**
   * Tạo headers yêu cầu kèm đính kèm địa chỉ đích thực tế X-Target-Url khi đi qua proxy
   */
  private getHeaders(extraHeaders: Record<string, string> = {}): Record<string, string> {
    const headers: Record<string, string> = {
      'Authorization': `Bearer ${this.apiKey}`,
      'Accept': 'application/json',
      ...extraHeaders
    };

    const currentOrigin = typeof window !== 'undefined' ? window.location.origin : '';
    if (!this.proxyDisabled && currentOrigin && !this.baseUrl.startsWith(currentOrigin)) {
      headers['X-Target-Url'] = this.baseUrl;
    }

    return headers;
  }

  /**
   * Kiểm tra kết nối tới Custom API.
   * Thử nghiệm gọi /models và dự phòng tới /chat/completions.
   */
  async testConnection(): Promise<boolean> {
    const errors: string[] = [];
    
    // Thử thách 1: Gọi thử endpoint /models
    try {
      const response = await fetch(this.getRequestUrl('models'), {
        method: 'GET',
        headers: this.getHeaders(),
      });
      if (response.ok) return true;
      errors.push(`/models trả về status ${response.status}`);
    } catch (e) {
      errors.push(`Gọi /models thất bại: ${e instanceof Error ? e.message : String(e)}`);
    }

    // Thử thách 2: Dự phòng gọi thử /chat/completions với cấu hình tối thiểu
    try {
      const response = await fetch(this.getRequestUrl('chat/completions'), {
        method: 'POST',
        headers: this.getHeaders({ 'Content-Type': 'application/json' }),
        body: JSON.stringify({
          model: 'gpt-3.5-turbo',
          messages: [{ role: 'user', content: 'ping' }],
          max_tokens: 1
        })
      });
      if (response.ok) return true;

      // Nếu trả về lỗi khác 401/403 (ví dụ 404 model not found) tức là key vẫn đúng và kết nối được
      if (response.status !== 401 && response.status !== 403) {
        return true;
      }
      errors.push(`/chat/completions trả về status ${response.status}`);
    } catch (e) {
      errors.push(`Gọi /chat/completions thất bại: ${e instanceof Error ? e.message : String(e)}`);
    }

    throw new Error(errors.join(' | '));
  }

  /**
   * Lấy danh sách các model hiện có tại API gateway.
   */
  async getModels(): Promise<string[]> {
    try {
      const response = await fetch(this.getRequestUrl('models'), {
        method: 'GET',
        headers: this.getHeaders(),
      });
      if (!response.ok) return [];
      const data = await response.json();
      return data.data ? data.data.map((m: { id: string }) => m.id) : [];
    } catch {
      return [];
    }
  }

  /**
   * Gửi yêu cầu sinh văn bản theo cấu trúc OpenAI Chat Completions.
   */
  async generate(
    options: CustomApiGenerateOptions,
    taskType: TaskType = 'custom'
  ): Promise<ApiResponse> {
    const startTime = Date.now();
    const isStream = options.stream ?? false;

    const messages = [];
    if (options.system) {
      messages.push({ role: 'system', content: options.system });
    }
    messages.push({ role: 'user', content: options.prompt });

    const body = {
      model: options.model,
      messages: messages,
      stream: isStream,
      temperature: options.temperature ?? 0.7,
    };

    const response = await fetch(this.getRequestUrl('chat/completions'), {
      method: 'POST',
      headers: this.getHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify(body),
      signal: options.signal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Custom API Error (${response.status}): ${errorText}`);
    }

    if (isStream && response.body) {
      let accumulatedText = '';
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let inputTokens = 0;
      let outputTokens = 0;

      try {
        let buffer = '';
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          // Giữ dòng cuối trong buffer
          buffer = lines.pop() || '';

          for (const line of lines) {
            const cleanedLine = line.trim();
            if (cleanedLine === '') continue;
            if (cleanedLine.startsWith('data: [DONE]')) continue;

            if (cleanedLine.startsWith('data: ')) {
              const jsonStr = cleanedLine.slice(6);
              try {
                const parsed = JSON.parse(jsonStr);
                const delta = parsed.choices?.[0]?.delta?.content;
                if (delta) {
                  accumulatedText += delta;
                  if (options.onChunk) {
                    options.onChunk(delta);
                  }
                }
                if (parsed.usage) {
                  inputTokens = parsed.usage.prompt_tokens ?? 0;
                  outputTokens = parsed.usage.completion_tokens ?? 0;
                }
              } catch (err) {
                // Đôi khi dòng SSE bị cắt nửa chừng, bỏ qua lỗi phân tích cục bộ
              }
            }
          }
        }
      } finally {
        reader.releaseLock();
      }

      return {
        success: true,
        task: taskType,
        sourceText: options.prompt,
        result: accumulatedText,
        model: options.model,
        tokens: {
          input: inputTokens || Math.round(options.prompt.length / 4),
          output: outputTokens || Math.round(accumulatedText.length / 4),
        },
        processingTime: (Date.now() - startTime) / 1000,
        timestamp: new Date().toISOString(),
      };
    } else {
      // Non-streaming response
      const data = await response.json();
      return {
        success: true,
        task: taskType,
        sourceText: options.prompt,
        result: data.choices[0].message.content,
        model: options.model,
        tokens: {
          input: data.usage?.prompt_tokens ?? Math.round(options.prompt.length / 4),
          output: data.usage?.completion_tokens ?? Math.round(data.choices[0].message.content.length / 4),
        },
        processingTime: (Date.now() - startTime) / 1000,
        timestamp: new Date().toISOString(),
      };
    }
  }
}
