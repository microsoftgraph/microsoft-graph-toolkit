import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import viteTsconfigPaths from 'vite-tsconfig-paths';

export default defineConfig(() => {
  return {
    build: {
      outDir: 'build'
    },
    plugins: [react(), viteTsconfigPaths()],
    server: {
      port: 3000,
      strictPort: true
    }
  };
});
