// Ťahadlo na zmenu šírky stĺpca. Počas ťahania hlási nový rozmer, po pustení ho uloží.
import { useRef } from 'react';

export function Splitter(props: {
  /** poloha ťahadla v px od ľavého (side=left) alebo pravého (side=right) okraja rodiča */
  at: number;
  side?: 'left' | 'right';
  value: number;
  min: number;
  max: number;
  onChange: (v: number) => void;
  onDone: (v: number) => void;
  label: string;
  /** predvolená šírka (dvojklik) */
  def?: number;
}) {
  const start = useRef<{ x: number; v: number } | null>(null);
  const last = useRef(props.value);
  const dir = props.side === 'right' ? -1 : 1;
  const clamp = (v: number) => Math.round(Math.min(props.max, Math.max(props.min, v)));
  return (
    <div
      className="splitter"
      role="separator"
      aria-orientation="vertical"
      aria-label={props.label}
      aria-valuenow={props.value}
      tabIndex={0}
      title={`${props.label} – ťahaním zmeníte šírku, dvojklik vráti predvolenú`}
      style={props.side === 'right' ? { right: props.at - 3 } : { left: props.at - 3 }}
      onPointerDown={(e) => {
        e.preventDefault();
        (e.target as HTMLElement).setPointerCapture(e.pointerId);
        start.current = { x: e.clientX, v: props.value };
        document.body.classList.add('resizing');
      }}
      onPointerMove={(e) => {
        if (!start.current) return;
        last.current = clamp(start.current.v + dir * (e.clientX - start.current.x));
        props.onChange(last.current);
      }}
      onPointerUp={() => {
        if (!start.current) return;
        start.current = null;
        document.body.classList.remove('resizing');
        props.onDone(last.current);
      }}
      onDoubleClick={() => {
        if (props.def == null) return;
        props.onChange(props.def);
        props.onDone(props.def);
      }}
      onKeyDown={(e) => {
        const step = e.shiftKey ? 40 : 10;
        if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
          e.preventDefault();
          const v = clamp(props.value + (e.key === 'ArrowRight' ? step : -step) * dir);
          props.onChange(v);
          props.onDone(v);
        }
      }}
    />
  );
}
