import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      // Run against the packages' TypeScript source — no build step needed.
      '@cobuildx.ai/office-viewer/react': resolve(__dirname, '../packages/core/src/react/index.ts'),
      '@cobuildx.ai/office-viewer': resolve(__dirname, '../packages/core/src/index.ts'),
    },
  },
});
