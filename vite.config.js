import { defineConfig } from 'vite';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { dirname, resolve } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 使用官方開發者憑證啟動 HTTPS
export default defineConfig({
  base: '/latex2svg/',
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync('C:\\Users\\ASUS\\.office-addin-dev-certs\\localhost.key'),
      cert: fs.readFileSync('C:\\Users\\ASUS\\.office-addin-dev-certs\\localhost.crt')
    }
  },
  build: {
    rollupOptions: {
      input: {
        commands: resolve(__dirname, 'commands.html')
      }
    }
  }
});
