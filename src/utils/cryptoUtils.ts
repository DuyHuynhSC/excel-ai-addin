/**
 * Tiện ích mã hóa để bảo vệ các API key lưu trữ tại LocalStorage.
 * Sử dụng Web Crypto API (AES-GCM 256-bit).
 */

const KEY_STORAGE_KEY = 'excel_addin_sec_k';

// Lấy hoặc sinh mới khóa mã hóa từ kho lưu trữ
async function getOrCreateKey(): Promise<CryptoKey> {
  const existingKeyJwk = localStorage.getItem(KEY_STORAGE_KEY);
  if (existingKeyJwk) {
    try {
      const jwk = JSON.parse(existingKeyJwk);
      return await globalThis.crypto.subtle.importKey(
        'jwk',
        jwk,
        { name: 'AES-GCM', length: 256 },
        true,
        ['encrypt', 'decrypt']
      );
    } catch (e) {
      console.error('Lỗi khi khôi phục khóa mã hóa, sẽ tạo mới:', e);
    }
  }

  // Tạo khóa mới
  const key = await globalThis.crypto.subtle.generateKey(
    { name: 'AES-GCM', length: 256 },
    true,
    ['encrypt', 'decrypt']
  );

  const exported = await globalThis.crypto.subtle.exportKey('jwk', key);
  localStorage.setItem(KEY_STORAGE_KEY, JSON.stringify(exported));
  return key;
}

/**
 * Mã hóa chuỗi văn bản
 * @param text Văn bản cần mã hóa (API key)
 */
export async function encryptText(text: string): Promise<string> {
  if (!text) return '';
  try {
    const key = await getOrCreateKey();
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
    const encoder = new TextEncoder();
    const encodedText = encoder.encode(text);

    const encryptedContent = await globalThis.crypto.subtle.encrypt(
      { name: 'AES-GCM', iv },
      key,
      encodedText
    );

    // Chuyển kết quả và IV sang dạng Hex/Base64 để lưu trữ
    const encryptedArray = new Uint8Array(encryptedContent);
    const ivHex = Array.from(iv).map(b => b.toString(16).padStart(2, '0')).join('');
    const contentHex = Array.from(encryptedArray).map(b => b.toString(16).padStart(2, '0')).join('');

    return `${ivHex}:${contentHex}`;
  } catch (error) {
    console.error('Lỗi mã hóa dữ liệu:', error);
    throw new Error('Không thể mã hóa thông tin bảo mật.');
  }
}

/**
 * Giải mã chuỗi văn bản
 * @param encryptedText Chuỗi đã được mã hóa ở dạng iv:content
 */
export async function decryptText(encryptedText: string): Promise<string> {
  if (!encryptedText) return '';
  try {
    if (!encryptedText.includes(':')) {
      // Khóa dạng plain text (ví dụ nạp từ tệp .env), trả về trực tiếp
      return encryptedText;
    }
    const parts = encryptedText.split(':');
    if (parts.length !== 2) {
      throw new Error('Định dạng chuỗi mã hóa không hợp lệ.');
    }

    const [ivHex, contentHex] = parts;
    const iv = new Uint8Array(ivHex.match(/.{1,2}/g)!.map(byte => parseInt(byte, 16)));
    const content = new Uint8Array(contentHex.match(/.{1,2}/g)!.map(byte => parseInt(byte, 16)));

    const key = await getOrCreateKey();
    const decryptedContent = await globalThis.crypto.subtle.decrypt(
      { name: 'AES-GCM', iv },
      key,
      content
    );

    const decoder = new TextDecoder();
    return decoder.decode(decryptedContent);
  } catch (error) {
    console.error('Lỗi giải mã dữ liệu:', error);
    // Nếu giải mã thất bại do khóa đã bị đổi, trả về chuỗi rỗng
    return '';
  }
}
