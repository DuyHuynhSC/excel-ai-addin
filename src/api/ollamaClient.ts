import { ApiResponse, TaskType } from '../types';

export interface OllamaGenerateOptions {
  model: string;
  prompt: string;
  system?: string;
  stream?: boolean;
  temperature?: number;
  onChunk?: (chunk: string) => void;
  signal?: AbortSignal;
}

export class OllamaClient {
  private baseUrl: string;

  constructor(baseUrl: string) {
    this.baseUrl = baseUrl.replace(/\/$/, '');
  }

  /**
   * Kiểm tra kết nối tới Ollama.
   * Trả về true nếu kết nối thành công.
   */
  async testConnection(): Promise<boolean> {
    try {
      const response = await fetch(`${this.baseUrl}/api/tags`, {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
      });
      if (response.ok) return true;
      throw new Error(`Ollama trả về status ${response.status}`);
    } catch (e) {
      throw new Error(`Ollama connection failed: ${e instanceof Error ? e.message : String(e)}`);
    }
  }

  /**
   * Lấy danh sách các model hiện có trong Ollama local.
   */
  async getModels(): Promise<string[]> {
    try {
      const response = await fetch(`${this.baseUrl}/api/tags`, {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
      });
      if (!response.ok) return [];
      const data = await response.json();
      return data.models ? data.models.map((m: { name: string }) => m.name) : [];
    } catch {
      return [];
    }
  }

  /**
   * Gửi yêu cầu sinh văn bản tới Ollama (hỗ trợ streaming).
   */
  async generate(
    options: OllamaGenerateOptions,
    taskType: TaskType = 'custom'
  ): Promise<ApiResponse> {
    const startTime = Date.now();
    const isStream = options.stream ?? false;
    const body = {
      model: options.model,
      prompt: options.prompt,
      system: options.system,
      stream: isStream,
      options: {
        temperature: options.temperature ?? 0.7,
      }
    };

    const response = await fetch(`${this.baseUrl}/api/generate`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
      signal: options.signal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Ollama API Error (${response.status}): ${errorText}`);
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
          // Giữ lại dòng cuối cùng chưa hoàn chỉnh trong buffer
          buffer = lines.pop() || '';

          for (const line of lines) {
            if (line.trim() === '') continue;
            try {
              const parsed = JSON.parse(line);
              if (parsed.response) {
                accumulatedText += parsed.response;
                if (options.onChunk) {
                  options.onChunk(parsed.response);
                }
              }
              if (parsed.done) {
                inputTokens = parsed.prompt_eval_count ?? 0;
                outputTokens = parsed.eval_count ?? 0;
              }
            } catch (err) {
              console.warn('Lỗi phân tích dòng stream JSON:', err);
            }
          }
        }

        // Xử lý nốt phần buffer còn lại
        if (buffer.trim() !== '') {
          try {
            const parsed = JSON.parse(buffer);
            if (parsed.response) {
              accumulatedText += parsed.response;
              if (options.onChunk) {
                options.onChunk(parsed.response);
              }
            }
          } catch {}
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
        result: data.response,
        model: options.model,
        tokens: {
          input: data.prompt_eval_count ?? Math.round(options.prompt.length / 4),
          output: data.eval_count ?? Math.round(data.response.length / 4),
        },
        processingTime: (Date.now() - startTime) / 1000,
        timestamp: new Date().toISOString(),
      };
    }
  }
}
