// @ts-check
import { defineConfig } from 'astro/config';
import tailwindcss from '@tailwindcss/vite';

// GitHub Pages 專案站：base 需為 /<repo-name>/
// Vercel：維持 base = '/'（建置時不要設 BASE）
const base = process.env.BASE || '/';

// https://astro.build/config
export default defineConfig({
  site: process.env.SITE || 'https://sixshoes.github.io',
  base,
  vite: {
    plugins: [tailwindcss()],
  },
});
