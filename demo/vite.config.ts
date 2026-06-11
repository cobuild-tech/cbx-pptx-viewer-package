import { resolve } from 'node:path';
import { defineConfig } from 'vite';

export default defineConfig({
  resolve: {
    alias: {
      // Develop the demo directly against core's TypeScript source (no build step).
      '@pptx-viewer/core': resolve(__dirname, '../packages/core/src/index.ts'),
    },
  },
});
