import { defineConfig } from 'vite';
import { resolve } from 'path';
import fs from 'fs';
import os from 'os';

const homedir = os.homedir();
const certPath = resolve(homedir, '.office-addin-dev-certs/localhost.crt');
const keyPath = resolve(homedir, '.office-addin-dev-certs/localhost.key');

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
    }
  }
});
