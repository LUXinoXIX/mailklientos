// Vložený web ako značka <webview>: vykresľuje sa vnútri stránky (žiadna samostatná vrstva nad oknom).
// Ovládanie (späť, obnoviť, adresa) cez ref; názov stránky a chyby načítania cez udalosti prvku.
import { forwardRef, useEffect, useImperativeHandle, useRef } from 'react';

export interface WebSlotHandle {
  back(): void;
  reload(): void;
  url(): string;
  /** spustí skript v stránke (napr. zistenie počtu neprečítaných) */
  exec<T = unknown>(js: string): Promise<T | null>;
}

interface WebviewElement extends HTMLElement {
  canGoBack(): boolean;
  goBack(): void;
  reload(): void;
  getURL(): string;
  executeJavaScript(js: string): Promise<unknown>;
}

export const WebSlot = forwardRef<
  WebSlotHandle,
  { url: string; partition: string; visible?: boolean; className?: string; onTitle?: (title: string) => void }
>(function WebSlot(props, ref) {
  const el = useRef<WebviewElement | null>(null);
  const ready = useRef(false);
  const onTitle = useRef(props.onTitle);
  onTitle.current = props.onTitle;

  useImperativeHandle(ref, () => ({
    back: () => {
      if (ready.current && el.current?.canGoBack()) el.current.goBack();
    },
    reload: () => {
      if (ready.current) el.current?.reload();
    },
    url: () => (ready.current ? el.current?.getURL() ?? props.url : props.url),
    exec: async <T,>(js: string) => {
      if (!ready.current || !el.current) return null;
      try {
        return (await el.current.executeJavaScript(js)) as T;
      } catch {
        return null;
      }
    },
  }));

  useEffect(() => {
    const w = el.current;
    if (!w) return;
    const onReady = () => {
      ready.current = true;
    };
    const onTitleEv = (e: Event) => onTitle.current?.((e as Event & { title: string }).title);
    w.addEventListener('dom-ready', onReady);
    w.addEventListener('page-title-updated', onTitleEv);
    return () => {
      w.removeEventListener('dom-ready', onReady);
      w.removeEventListener('page-title-updated', onTitleEv);
    };
  }, []);

  // src sa nastaví len raz: ďalej si stránka drží vlastnú navigáciu (zmena src by ju načítala znova)
  const src = useRef(props.url);

  if (!window.mailApi) {
    return (
      <div className={`webslot ${props.className ?? ''}`} style={props.visible === false ? { display: 'none' } : undefined}>
        <div className="webslot-hint">Vložený web sa zobrazí v aplikácii.</div>
      </div>
    );
  }
  return (
    <webview
      ref={el}
      src={src.current}
      partition={props.partition}
      allowpopups="true"
      className={`webslot ${props.className ?? ''} ${props.visible === false ? 'webslot-hidden' : ''}`}
    />
  );
});
