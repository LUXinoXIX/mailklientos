import { createLightTheme, type BrandVariants, type Theme } from '@fluentui/react-components';

export type ThemeKey = 'blue' | 'sage' | 'gray';

export const THEMES: Record<ThemeKey, { label: string; accent: string }> = {
  blue: { label: 'Modrá (MS365)', accent: '#0F6CBD' },
  sage: { label: 'Sivozelená', accent: '#2F6F5E' },
  gray: { label: 'Sivá', accent: '#4B5563' },
};

const hex = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
const toHex = (c: number[]) => `#${c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('')}`;
const mix = (a: number[], b: number[], t: number) => a.map((v, i) => v + (b[i] - v) * t);

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

const cache = new Map<ThemeKey, Theme>();
export function getTheme(key: ThemeKey): Theme {
  if (!cache.has(key)) cache.set(key, createLightTheme(brandRamp(THEMES[key].accent)));
  return cache.get(key)!;
}

/** Farby kont (pás kont, avatary, horný pruh). Tmavšie odtiene, aby biely text mal dobrý kontrast. */
export const ACCOUNT_COLORS = ['#2F6F5E', '#1d4ed8', '#6d28d9', '#b45309', '#be123c', '#0e7490', '#4B5563', '#3f6212'];

/** Staré svetlé farby z kroku 2 premapované na nové tmavšie (tyrkysová → sivozelená). */
const LEGACY: Record<string, string> = {
  '#14b8a6': '#2F6F5E', '#8b5cf6': '#6d28d9', '#3b82f6': '#1d4ed8', '#f59e0b': '#b45309',
  '#ef4444': '#be123c', '#10b981': '#2F6F5E', '#ec4899': '#be123c',
};
export const accountColor = (c: string | undefined) => (c ? LEGACY[c.toLowerCase()] ?? c : '#4B5563');
