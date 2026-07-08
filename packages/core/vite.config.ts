import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import dts from 'vite-plugin-dts';

export default defineConfig({
  build: {
    lib: {
      entry: {
        index: resolve(__dirname, 'src/index.ts'),
        react: resolve(__dirname, 'src/react/index.ts'),
      },
      formats: ['es'],
      fileName: (_format, entryName) => `${entryName}.js`,
    },
    rollupOptions: {
      external: ['react', 'react-dom', 'react/jsx-runtime'],
    },
    sourcemap: true,
    target: 'es2022',
  },
  plugins: [dts({ rollupTypes: true, include: ['src'] })],
  test: {
    globals: true,
    environment: 'node',
  },
});
