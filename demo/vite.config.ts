import { resolve } from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  resolve: {
    alias: {
      // Develop the demo directly against the package's TypeScript source (no build step).
      '@cobuildx.ai/office-viewer': resolve(__dirname, '../packages/core/src/index.ts'),
    },
  },
});
