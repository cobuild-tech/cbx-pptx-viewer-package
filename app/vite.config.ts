import { resolve } from 'node:path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';

// Set USE_LOCAL_PACKAGE=true (in app/.env or the shell) to develop the app
// directly against packages/core's TypeScript source — hot-reloads on every
// edit, no build/publish needed. Leave it unset/false to consume the real
// published @cobuildx.ai/office-viewer from npm, exactly as an external app does.
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '');
  const useLocal = env.USE_LOCAL_PACKAGE === 'true';

  return {
    plugins: [react()],
    resolve: {
      alias: useLocal
        ? [
            // Order matters: the more specific "/react" entry must precede the base one.
            {
              find: '@cobuildx.ai/office-viewer/react',
              replacement: resolve(__dirname, '../packages/core/src/react/index.ts'),
            },
            {
              find: '@cobuildx.ai/office-viewer',
              replacement: resolve(__dirname, '../packages/core/src/index.ts'),
            },
          ]
        : [],
    },
    // In local mode the core source pulls in these deps directly; pre-bundle them
    // upfront so Vite doesn't re-optimize mid-session (which invalidates hashed
    // URLs and 504s the browser). No-op when consuming the pre-built npm package.
    optimizeDeps: useLocal ? { include: ['fflate', 'fast-xml-parser'] } : {},
  };
});
