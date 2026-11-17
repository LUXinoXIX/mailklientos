// Knižnice bez vlastných typov (alebo s typmi, ktoré tu nepotrebujeme presne).
declare module 'imapflow';
declare module 'mailparser';
declare module 'nodemailer';
declare module 'nodemailer/lib/mail-composer';
declare module '*.css';
declare module '*.svg' { const url: string; export default url; }
