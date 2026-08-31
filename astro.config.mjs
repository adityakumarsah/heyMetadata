// @ts-check
import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// https://astro.build/config
export default defineConfig({
  site: 'https://heymetadata.com',
  integrations: [sitemap()],
  vite: {
    optimizeDeps: {
      exclude: ['@jsquash/avif'],
    },
    assetsInclude: ['**/*.wasm'],
    server: {
      headers: {
        'Cross-Origin-Opener-Policy': 'same-origin',
        'Cross-Origin-Embedder-Policy': 'require-corp',
      },
    },
  },
});