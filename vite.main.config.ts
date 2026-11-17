import { defineConfig } from 'vite';

// Hlavný proces. Knižnice (imapflow, mailparser, nodemailer) sú v devDependencies,
// aby ich Vite zabalil priamo do .vite/build/main.js (inštalátor nepotrebuje node_modules).
// node:sqlite je vstavaný modul Electronu (Node 24); Vite ho nesmie zabaliť.
export default defineConfig({
  build: { rollupOptions: { external: ['node:sqlite'] } },
});
