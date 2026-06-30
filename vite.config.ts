import { defineConfig } from 'vite';
import { resolve } from 'path';
import fs from 'fs';
import os from 'os';
import { HttpsProxyAgent } from 'https-proxy-agent';

const homedir = os.homedir();
const certPath = resolve(homedir, '.office-addin-dev-certs/localhost.crt');
const keyPath = resolve(homedir, '.office-addin-dev-certs/localhost.key');

const proxyUrl = process.env.HTTPS_PROXY || process.env.HTTP_PROXY || process.env.https_proxy || process.env.http_proxy;
const proxyAgent = proxyUrl ? new HttpsProxyAgent(proxyUrl) : undefined;

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
        agent: proxyAgent, // Định tuyến qua proxy hệ thống (HTTP_PROXY / HTTPS_PROXY) nếu có
        router: (req: any) => {
          const targetHeader = req.headers['x-target-url'];
          if (typeof targetHeader === 'string' && targetHeader.trim() !== '') {
            return targetHeader;
          }
          return 'https://api.company.com/v1';
        },
        rewrite: (path: string) => path.replace(/^\/api-proxy/, ''),
      } as any
    }
  }
});
