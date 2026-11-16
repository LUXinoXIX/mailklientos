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

/** Farby kont (bodky, avatary). */
export const ACCOUNT_COLORS = ['#14b8a6', '#8b5cf6', '#3b82f6', '#f59e0b', '#ef4444', '#10b981', '#ec4899'];
