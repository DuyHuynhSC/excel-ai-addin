import { defineConfig } from 'vite';
import { resolve } from 'path';
import fs from 'fs';
import os from 'os';
import { HttpsProxyAgent } from 'https-proxy-agent';
import * as https from 'https';

const homedir = os.homedir();
const certPath = resolve(homedir, '.office-addin-dev-certs/localhost.crt');
const keyPath = resolve(homedir, '.office-addin-dev-certs/localhost.key');

const proxyUrl = process.env.HTTPS_PROXY || process.env.HTTP_PROXY || process.env.https_proxy || process.env.http_proxy;

// Hàm kiểm tra xem URL đích có được bypass qua proxy không (dành cho mạng nội bộ/intranet)
function shouldBypassProxy(targetUrl: string): boolean {
  if (!targetUrl) return true;
  
  try {
    const url = new URL(targetUrl);
    const host = url.hostname.toLowerCase();
    
    // Mặc định bỏ qua proxy cho localhost
    if (host === 'localhost' || host === '127.0.0.1') {
      return true;
    }
    
    // Mặc định bỏ qua proxy cho các dải IP Private (mạng nội bộ)
    const privateIpRegex = /^(?:10\.\d+\.\d+\.\d+|172\.(?:1[6-9]|2\d|3[0-1])\.\d+\.\d+|192\.168\.\d+\.\d+)$/;
    if (privateIpRegex.test(host)) {
      return true;
    }
    
    const noProxy = process.env.NO_PROXY || process.env.no_proxy || '';
    if (!noProxy) return false;
    
    const bypassList = noProxy.split(',').map(item => item.trim().toLowerCase());
    for (const bypass of bypassList) {
      if (!bypass) continue;
      if (host === bypass || 
          (bypass.startsWith('.') && host.endsWith(bypass)) || 
          host.endsWith('.' + bypass) ||
          bypass === '*') {
        return true;
      }
    }
  } catch (e) {
    return true; // Nếu URL lỗi định dạng, đi trực tiếp
  }
  return false;
}

// Tạo Custom Agent để định tuyến động giữa kết nối trực tiếp và đi qua proxy của công ty
class DynamicAgent extends https.Agent {
  private proxyAgent: any;
  private directAgent: any;

  constructor(pUrl: string | undefined) {
    super({ keepAlive: true, rejectUnauthorized: false });
    this.directAgent = new https.Agent({ keepAlive: true, rejectUnauthorized: false });
    
    if (pUrl) {
      try {
        const parsed = new URL(pUrl);
        this.proxyAgent = new HttpsProxyAgent({
          protocol: parsed.protocol,
          host: parsed.hostname,
          port: parsed.port,
          auth: parsed.username && parsed.password ? `${parsed.username}:${parsed.password}` : undefined,
          rejectUnauthorized: false // Bỏ qua kiểm tra chứng chỉ tự ký của proxy công ty khi kết nối ngoài
        });
      } catch (e) {
        console.warn('Lỗi phân tích cú pháp proxyUrl, chuyển sang agent thô:', e);
        this.proxyAgent = new HttpsProxyAgent(pUrl);
      }
    } else {
      this.proxyAgent = null;
    }
  }

  addRequest(req: any, options: any) {
    const targetUrl = req.getHeader('x-target-url') || req.getHeader('X-Target-Url') || '';
    if (shouldBypassProxy(targetUrl)) {
      return this.directAgent.addRequest(req, options);
    }
    if (this.proxyAgent) {
      return this.proxyAgent.addRequest(req, options);
    }
    return this.directAgent.addRequest(req, options);
  }
}

const proxyAgent = new DynamicAgent(proxyUrl);

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'taskpane.html'),
      },
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: '[name].js',
        assetFileNames: '[name].[ext]',
      }
    },
    outDir: 'dist',
    emptyOutDir: true,
  },
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    },
    proxy: {
      '/api-proxy': {
        target: 'https://api.company.com/v1', // Giá trị mặc định dự phòng
        changeOrigin: true,
        secure: false, // Hỗ trợ bỏ qua kiểm tra chứng chỉ tự ký của gateway nội bộ
        agent: proxyAgent, // Định tuyến qua proxy hoặc đi trực tiếp động
        router: (req: any) => {
          const targetHeader = req.headers['x-target-url'];
          if (typeof targetHeader === 'string' && targetHeader.trim() !== '') {
            return targetHeader;
          }
          return 'https://api.company.com/v1';
        },
        rewrite: (path: string) => path.replace(/^\/api-proxy/, ''),
        configure: (proxy: any) => {
          proxy.on('error', (err: any, _req: any, res: any) => {
            console.error('[Vite Proxy Error]:', err.message);
            res.writeHead(500, {
              'Content-Type': 'text/plain; charset=utf-8',
            });
            res.end(`PROXY_CONNECTION_ERROR: ${err.message}`);
          });
          proxy.on('proxyReq', (_proxyReq: any, req: any, _res: any) => {
            const targetUrl = req.headers['x-target-url'] || 'https://api.company.com/v1';
            console.log(`[Proxy Request] ${req.method} ${req.url} -> Target: ${targetUrl}`);
          });
        }
      } as any
    }
  }
});
