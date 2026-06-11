import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      // Run against the packages' TypeScript source — no build step needed.
      '@pptx-viewer/core': resolve(__dirname, '../packages/core/src/index.ts'),
      '@pptx-viewer/react': resolve(__dirname, '../packages/react/src/index.ts'),
    },
  },
});
