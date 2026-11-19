// Čítanie RSS 2.0 / Atom / RDF kanálov v rozhraní (DOMParser) a nájdenie kanála na bežnej stránke.
export interface NewsItem {
  title: string;
  link: string;
  date?: string;
  image?: string;
}
export interface Feed {
  title: string;
  items: NewsItem[];
}

const txt = (el: Element | null | undefined) => (el?.textContent ?? '').replace(/\s+/g, ' ').trim();
const abs = (u: string, base: string) => {
  try {
    return new URL(u, base).toString();
  } catch {
    return u;
  }
};

export function parseFeed(text: string, base: string): Feed | null {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.querySelector('parsererror')) return null;
  const root = doc.documentElement;
  const name = root.localName.toLowerCase();
  if (name === 'feed') {
    // Atom
    const items = [...root.getElementsByTagNameNS('*', 'entry')].map((e) => {
      const links = [...e.getElementsByTagNameNS('*', 'link')];
      const l = links.find((x) => (x.getAttribute('rel') ?? 'alternate') === 'alternate') ?? links[0];
      const thumb = e.getElementsByTagNameNS('*', 'thumbnail')[0]?.getAttribute('url') ?? undefined;
      return { title: txt(e.getElementsByTagNameNS('*', 'title')[0]), link: abs(l?.getAttribute('href') ?? '', base), date: txt(e.getElementsByTagNameNS('*', 'published')[0]) || txt(e.getElementsByTagNameNS('*', 'updated')[0]) || undefined, image: thumb };
    });
    return { title: txt(root.getElementsByTagNameNS('*', 'title')[0]), items };
  }
  if (name === 'rss' || name === 'rdf') {
    const items = [...root.getElementsByTagNameNS('*', 'item')].map((e) => {
      const enc = [...e.getElementsByTagName('enclosure')].find((x) => /^image\//.test(x.getAttribute('type') ?? ''))?.getAttribute('url');
      const media = e.getElementsByTagNameNS('*', 'thumbnail')[0]?.getAttribute('url') ?? e.getElementsByTagNameNS('*', 'content')[0]?.getAttribute('url');
      const date = txt(e.getElementsByTagName('pubDate')[0]) || txt(e.getElementsByTagNameNS('*', 'date')[0]);
      return { title: txt(e.getElementsByTagName('title')[0]), link: abs(txt(e.getElementsByTagName('link')[0]) || txt(e.getElementsByTagName('guid')[0]), base), date: date || undefined, image: enc || media || undefined };
    });
    const ch = root.getElementsByTagName('channel')[0];
    return { title: txt(ch?.getElementsByTagName('title')[0]), items };
  }
  return null;
}

/** Na bežnej stránke nájde odkaz na kanál (<link rel="alternate" type="application/rss+xml">). */
export function discoverFeed(html: string, base: string): string | null {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const l = doc.querySelector('link[rel~="alternate"][type="application/rss+xml"], link[rel~="alternate"][type="application/atom+xml"]');
  const href = l?.getAttribute('href');
  return href ? abs(href, base) : null;
}

/** Stiahne kanál (aj zo stránky, ktorá naň len odkazuje). */
export async function loadFeed(url: string): Promise<Feed> {
  const api = window.mailApi!;
  let r = await api.rssFetch(url);
  if (!r.ok) throw new Error(r.error);
  let f = parseFeed(r.data.text, r.data.url);
  if (!f) {
    const alt = discoverFeed(r.data.text, r.data.url);
    if (!alt) throw new Error('Na stránke sa nenašiel RSS kanál. Skús adresu kanála (napr. …/feed alebo …/rss).');
    r = await api.rssFetch(alt);
    if (!r.ok) throw new Error(r.error);
    f = parseFeed(r.data.text, r.data.url);
    if (!f) throw new Error('Kanál sa nedá prečítať.');
  }
  // bezpečnosť: len http(s) odkazy
  f.items = f.items.filter((i) => i.title && /^https?:\/\//i.test(i.link)).map((i) => ({ ...i, image: i.image && /^https:\/\//i.test(i.image) ? i.image : undefined }));
  return f;
}
