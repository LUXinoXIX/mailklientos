import { defineConfig } from 'vite';

// Hlavný proces. Knižnice (imapflow, mailparser, nodemailer) sú v devDependencies,
// aby ich Vite zabalil priamo do .vite/build/main.js (inštalátor nepotrebuje node_modules).
export default defineConfig({});
