import { defineConfig } from 'vite';

// JSX (React) cez vstavaný esbuild, bez ďalšieho pluginu.
export default defineConfig({
  esbuild: { jsx: 'automatic' },
});
