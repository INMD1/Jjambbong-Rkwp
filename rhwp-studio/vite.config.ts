import { defineConfig } from 'vite';
import { resolve } from 'path';
import { readFileSync } from 'fs';

const pkg = JSON.parse(readFileSync(resolve(__dirname, 'package.json'), 'utf-8'));

export default defineConfig({
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
  },
  optimizeDeps: {
    include: ['pako', 'saxes'],
  },
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
      '@wasm': resolve(__dirname, '..', 'pkg'),
      'hwpkit-extension': resolve(__dirname, '..', 'hwpkit-extension'),
      'pako': resolve(__dirname, 'node_modules', 'pako'),
      'saxes': resolve(__dirname, 'node_modules', 'saxes'),
    },
  },
  server: {
    host: '0.0.0.0',
    port: 7700,
    allowedHosts: true,
    fs: {
      allow: ['..'],
    },
  },
});
