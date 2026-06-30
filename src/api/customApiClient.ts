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

  constructor(baseUrl: string, apiKey: string) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
    this.apiKey = apiKey;
  }

  /**
   * Kiểm tra kết nối tới Custom API.
   * Thường gọi một endpoint nhỏ như lấy danh sách model.
   */
  async testConnection(): Promise<boolean> {
    try {
      const response = await fetch(`${this.baseUrl}/models`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`,
          'Accept': 'application/json'
        },
      });
      return response.ok;
    } catch (e) {
      console.warn('Lỗi kết nối tới Custom API:', e);
      return false;
    }
  }

  /**
   * Lấy danh sách các model hiện có tại API gateway.
   */
  async getModels(): Promise<string[]> {
    try {
      const response = await fetch(`${this.baseUrl}/models`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`,
          'Accept': 'application/json'
        },
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

    const response = await fetch(`${this.baseUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${this.apiKey}`,
      },
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
                // Nếu API trả kèm thông tin token usage ở chunk cuối cùng
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
