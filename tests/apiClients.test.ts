import { describe, it, expect, vi, beforeEach } from 'vitest';
import { OllamaClient } from '../src/api/ollamaClient';
import { CustomApiClient } from '../src/api/customApiClient';

describe('API Clients Unit Tests', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  describe('OllamaClient', () => {
    it('should test connection successfully', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
      });
      vi.stubGlobal('fetch', mockFetch);

      const client = new OllamaClient('http://localhost:11434');
      const connectionOk = await client.testConnection();
      
      expect(connectionOk).toBe(true);
      expect(mockFetch).toHaveBeenCalledWith('http://localhost:11434/api/tags', expect.anything());
    });

    it('should return list of models from tags', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          models: [
            { name: 'phi:2b' },
            { name: 'mistral:7b' }
          ]
        })
      });
      vi.stubGlobal('fetch', mockFetch);

      const client = new OllamaClient('http://localhost:11434');
      const models = await client.getModels();
      
      expect(models).toEqual(['phi:2b', 'mistral:7b']);
    });

    it('should generate text successfully (non-streaming)', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          response: 'Bản dịch tiếng Việt',
          prompt_eval_count: 10,
          eval_count: 20
        })
      });
      vi.stubGlobal('fetch', mockFetch);

      const client = new OllamaClient('http://localhost:11434');
      const res = await client.generate({
        model: 'phi:2b',
        prompt: 'Translate hello',
        stream: false
      });

      expect(res.success).toBe(true);
      expect(res.result).toBe('Bản dịch tiếng Việt');
      expect(res.tokens?.input).toBe(10);
      expect(res.tokens?.output).toBe(20);
    });
  });

  describe('CustomApiClient', () => {
    it('should query completions successfully', async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          choices: [
            {
              message: {
                content: 'Bản dịch từ Custom API'
              }
            }
          ],
          usage: {
            prompt_tokens: 15,
            completion_tokens: 30
          }
        })
      });
      vi.stubGlobal('fetch', mockFetch);

      const client = new CustomApiClient('https://api.company.com/v1', 'mock_key');
      const res = await client.generate({
        model: 'gpt-3.5-turbo',
        prompt: 'Translate standard text',
        stream: false
      });

      expect(res.success).toBe(true);
      expect(res.result).toBe('Bản dịch từ Custom API');
      expect(res.tokens?.input).toBe(15);
      expect(res.tokens?.output).toBe(30);
    });
  });
});
