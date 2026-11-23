// Počasie pre dlaždicu Súhrnu – Open-Meteo (bez kľúča, bez registrácie): vyhľadanie miesta a predpoveď.
import { net } from 'electron';
import type { WeatherData, WeatherPlace } from '../core/api';
import { UserError } from './mail';

const cache = new Map<string, { at: number; data: WeatherData }>();

async function getJson(url: string): Promise<unknown> {
  const r = await net.fetch(url, { headers: { Accept: 'application/json' } });
  if (!r.ok) throw new UserError(`Služba počasia vrátila chybu ${r.status}.`);
  return r.json();
}

/** Vyhľadanie miesta podľa názvu (slovenské názvy). */
export async function searchPlace(q: string): Promise<WeatherPlace[]> {
  const name = q.trim();
  if (name.length < 2) return [];
  const j = (await getJson(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(name)}&count=8&language=sk&format=json`)) as {
    results?: { name: string; latitude: number; longitude: number; country?: string; admin1?: string }[];
  };
  return (j.results ?? []).map((r) => ({ name: r.name, lat: r.latitude, lon: r.longitude, detail: [r.admin1, r.country].filter(Boolean).join(', ') }));
}

/** Aktuálne počasie a predpoveď na 5 dní (výsledok sa pamätá 30 minút). */
export async function forecast(lat: number, lon: number): Promise<WeatherData> {
  const key = `${lat.toFixed(3)},${lon.toFixed(3)}`;
  const c = cache.get(key);
  if (c && Date.now() - c.at < 30 * 60_000) return c.data;
  const url =
    `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}` +
    '&current=temperature_2m,apparent_temperature,weather_code,wind_speed_10m,relative_humidity_2m,is_day' +
    '&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max' +
    '&timezone=auto&forecast_days=5';
  const j = (await getJson(url)) as {
    current?: { temperature_2m: number; apparent_temperature: number; weather_code: number; wind_speed_10m: number; relative_humidity_2m: number; is_day: number; time: string };
    daily?: { time: string[]; weather_code: number[]; temperature_2m_max: number[]; temperature_2m_min: number[]; precipitation_probability_max?: (number | null)[] };
  };
  if (!j.current || !j.daily) throw new UserError('Služba počasia nevrátila údaje.');
  const data: WeatherData = {
    current: {
      temp: j.current.temperature_2m,
      feels: j.current.apparent_temperature,
      code: j.current.weather_code,
      wind: j.current.wind_speed_10m,
      humidity: j.current.relative_humidity_2m,
      day: j.current.is_day === 1,
      time: j.current.time,
    },
    days: j.daily.time.map((d, i) => ({
      date: d,
      code: j.daily!.weather_code[i] ?? 0,
      max: j.daily!.temperature_2m_max[i] ?? 0,
      min: j.daily!.temperature_2m_min[i] ?? 0,
      rain: j.daily!.precipitation_probability_max?.[i] ?? null,
    })),
  };
  cache.set(key, { at: Date.now(), data });
  return data;
}
