import { defineConfig } from 'vite';

export default defineConfig({
  base: '/pptx-web/',
  build: {
    // The renderer is already isolated behind a dynamic import.
    chunkSizeWarningLimit: 1600,
  },
});
