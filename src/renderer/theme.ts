import { createDarkTheme, createLightTheme, type BrandVariants, type Theme } from '@fluentui/react-components';
import type { CSSProperties } from 'react';
import type { ThemeDef } from '../core/settings';

export type { ThemeDef };
export type ThemeKey = string;

/** Vstavané témy (vlastné si používateľ vytvorí v Zobrazenie → Vlastné témy). top = null → horný pruh vo farbe konta. */
export const BUILTIN_THEMES: ThemeDef[] = [
  { id: 'sage', name: 'Sivozelená', dark: false, accent: '#2F6F5E', top: null, rail: '#e3e6ea', pane: '#f5f6f7', read: '#ffffff', text: '#242424', radius: 4 },
  { id: 'blue', name: 'Modrá (MS365)', dark: false, accent: '#0F6CBD', top: null, rail: '#e3e6ea', pane: '#f5f5f5', read: '#ffffff', text: '#242424', radius: 4 },
  { id: 'gray', name: 'Sivá', dark: false, accent: '#4B5563', top: null, rail: '#e3e6ea', pane: '#f4f4f5', read: '#ffffff', text: '#242424', radius: 4 },
  { id: 'graphite', name: 'Grafit (tmavá)', dark: true, accent: '#4fb3a0', top: '#1f2326', rail: '#16191b', pane: '#202427', read: '#262b2e', text: '#e6e8ea', radius: 6 },
  { id: 'midnight', name: 'Polnočná modrá (tmavá)', dark: true, accent: '#5aa2ff', top: '#13233f', rail: '#0d1830', pane: '#132140', read: '#18284a', text: '#e3ebf7', radius: 8 },
  { id: 'nord', name: 'Nord', dark: false, accent: '#5e81ac', top: '#4c566a', rail: '#d8dee9', pane: '#e5e9f0', read: '#eceff4', text: '#2e3440', radius: 4 },
  { id: 'paper', name: 'Papier', dark: false, accent: '#9a5b2c', top: '#7a4a26', rail: '#ece3d4', pane: '#f4ede1', read: '#fbf7ef', text: '#3b2f22', radius: 8 },
  { id: 'contrast', name: 'Vysoký kontrast', dark: false, accent: '#0050c8', top: '#000000', rail: '#ffffff', pane: '#ffffff', read: '#ffffff', text: '#000000', radius: 0 },
];

/** Staré kľúče tém (blue/sage/gray) aj vlastné témy. */
export function themeById(id: string, custom: ThemeDef[] = []): ThemeDef {
  return custom.find((t) => t.id === id) ?? BUILTIN_THEMES.find((t) => t.id === id) ?? BUILTIN_THEMES[0]!;
}

/** Kompatibilita so staršou verziou: zoznam tém s akcentom. */
export const THEMES: Record<string, { label: string; accent: string }> = Object.fromEntries(BUILTIN_THEMES.map((t) => [t.id, { label: t.name, accent: t.accent }]));

const hex = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
const toHex = (c: number[]) => `#${c.map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('')}`;
const mix = (a: number[], b: number[], t: number) => a.map((v, i) => v + (b[i]! - v) * t);
const mixHex = (a: string, b: string, t: number) => toHex(mix(hex(a), hex(b), t));
export const luminance = (h: string) => {
  const [r, g, b] = hex(h);
  return (0.299 * r! + 0.587 * g! + 0.114 * b!) / 255;
};
/** Čitateľná farba textu na danom pozadí. */
export const onColor = (bg: string) => (luminance(bg) > 0.6 ? '#1b1b1b' : '#ffffff');

/** Fluent potrebuje 16 odtieňov značkovej farby (10 najtmavší … 160 najsvetlejší, 80 = akcent). */
export function brandRamp(accent: string): BrandVariants {
  const base = hex(accent);
  const out: Record<number, string> = {};
  for (let i = 1; i <= 16; i++) {
    const key = i * 10;
    out[key] =
      key < 80
        ? toHex(mix([0, 0, 0], base, 0.25 + (0.75 * (key - 10)) / 70))
        : toHex(mix(base, [255, 255, 255], ((key - 80) / 80) * 0.92));
  }
  return out as unknown as BrandVariants;
}

const cache = new Map<string, Theme>();
/** Fluent téma z definície: pozadia, text, okraje a zaoblenie podľa témy. */
export function getTheme(key: ThemeKey | ThemeDef, custom: ThemeDef[] = []): Theme {
  const t = typeof key === 'string' ? themeById(key, custom) : key;
  const ck = JSON.stringify(t);
  const hit = cache.get(ck);
  if (hit) return hit;
  const base = t.dark ? createDarkTheme(brandRamp(t.accent)) : createLightTheme(brandRamp(t.accent));
  const r = `${t.radius}px`;
  const theme: Theme = {
    ...base,
    colorNeutralBackground1: t.read,
    colorNeutralBackground1Hover: mixHex(t.read, t.text, 0.06),
    colorNeutralBackground1Pressed: mixHex(t.read, t.text, 0.1),
    colorNeutralBackground1Selected: mixHex(t.read, t.text, 0.08),
    colorNeutralBackground2: t.pane,
    colorNeutralBackground3: mixHex(t.pane, t.text, 0.06),
    colorNeutralBackground4: mixHex(t.pane, t.text, 0.1),
    colorNeutralForeground1: t.text,
    colorNeutralForeground2: mixHex(t.text, t.read, 0.18),
    colorNeutralForeground3: mixHex(t.text, t.read, 0.4),
    colorNeutralForeground4: mixHex(t.text, t.read, 0.55),
    colorNeutralStroke1: mixHex(t.read, t.text, 0.28),
    colorNeutralStroke2: mixHex(t.read, t.text, 0.16),
    colorNeutralStroke3: mixHex(t.read, t.text, 0.09),
    colorSubtleBackgroundHover: mixHex(t.read, t.text, 0.06),
    colorSubtleBackgroundPressed: mixHex(t.read, t.text, 0.1),
    borderRadiusSmall: `${Math.max(0, t.radius - 2)}px`,
    borderRadiusMedium: r,
    borderRadiusLarge: `${t.radius + 2}px`,
    borderRadiusXLarge: `${t.radius + 4}px`,
  };
  cache.set(ck, theme);
  return theme;
}

/** CSS premenné pre vlastné časti rozhrania (pás kont, horný pruh). */
export function themeVars(t: ThemeDef): CSSProperties {
  return {
    ['--accRailBg' as string]: t.rail,
    ['--railFg' as string]: luminance(t.rail) > 0.55 ? '#374151' : '#d6dadf',
    ['--radius' as string]: `${t.radius}px`,
    colorScheme: t.dark ? 'dark' : 'light',
  };
}

/** Farby kont (pás kont, avatary, horný pruh). Tmavšie odtiene, aby biely text mal dobrý kontrast. */
export const ACCOUNT_COLORS = ['#2F6F5E', '#1d4ed8', '#6d28d9', '#b45309', '#be123c', '#0e7490', '#4B5563', '#3f6212'];

/** Staré svetlé farby z kroku 2 premapované na nové tmavšie (tyrkysová → sivozelená). */
const LEGACY: Record<string, string> = {
  '#14b8a6': '#2F6F5E', '#8b5cf6': '#6d28d9', '#3b82f6': '#1d4ed8', '#f59e0b': '#b45309',
  '#ef4444': '#be123c', '#10b981': '#2F6F5E', '#ec4899': '#be123c',
};
export const accountColor = (c: string | undefined) => (c ? LEGACY[c.toLowerCase()] ?? c : '#4B5563');
