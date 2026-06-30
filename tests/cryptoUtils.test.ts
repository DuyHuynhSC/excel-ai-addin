import { describe, it, expect, beforeAll } from 'vitest';

// Mock LocalStorage for Node testing environment
const store: Record<string, string> = {};
beforeAll(() => {
  globalThis.localStorage = {
    getItem: (key: string) => store[key] || null,
    setItem: (key: string, value: string) => { store[key] = value; },
    removeItem: (key: string) => { delete store[key]; },
    clear: () => { for (const k in store) delete store[k]; },
    length: 0,
    key: () => null,
  } as unknown as Storage;
});

// Import after setting up mock
import { encryptText, decryptText } from '../src/utils/cryptoUtils';

describe('CryptoUtils Unit Tests', () => {
  it('should encrypt and decrypt back to the original text', async () => {
    const originalText = 'sk-company-api-gateway-key-12345';
    const encrypted = await encryptText(originalText);
    
    expect(encrypted).toBeDefined();
    expect(encrypted).toContain(':');
    expect(encrypted).not.toBe(originalText);

    const decrypted = await decryptText(encrypted);
    expect(decrypted).toBe(originalText);
  });

  it('should handle empty values gracefully', async () => {
    expect(await encryptText('')).toBe('');
    expect(await decryptText('')).toBe('');
  });

  it('should bypass decryption and return text directly if no colon is present', async () => {
    const plainTextKey = 'random-text-without-colon';
    const decrypted = await decryptText(plainTextKey);
    expect(decrypted).toBe(plainTextKey);
  });
});
