// Vlastné menu na pravý klik (správy, priečinky, prílohy) – otvorí sa na mieste kurzora.
import { Menu, MenuDivider, MenuItem, MenuList, MenuPopover, MenuTrigger } from '@fluentui/react-components';
import type { ReactNode } from 'react';

export type CtxItem =
  | { divider: true }
  | { label: string; icon?: ReactNode; onClick?: () => void; disabled?: boolean; sub?: CtxItem[]; hint?: string };

export interface CtxState {
  x: number;
  y: number;
  items: CtxItem[];
}

function Items({ items }: { items: CtxItem[] }) {
  return (
    <>
      {items.map((it, i) => {
        if ('divider' in it) return <MenuDivider key={`d${i}`} />;
        if (it.sub) {
          return (
            <Menu key={it.label}>
              <MenuTrigger disableButtonEnhancement>
                <MenuItem icon={it.icon as never} disabled={it.disabled || !it.sub.length}>{it.label}</MenuItem>
              </MenuTrigger>
              <MenuPopover style={{ maxHeight: 420, overflowY: 'auto' }}>
                <MenuList>
                  <Items items={it.sub} />
                </MenuList>
              </MenuPopover>
            </Menu>
          );
        }
        return (
          <MenuItem key={it.label} icon={it.icon as never} disabled={it.disabled} onClick={it.onClick} secondaryContent={it.hint}>
            {it.label}
          </MenuItem>
        );
      })}
    </>
  );
}

export function ContextMenu(props: { state: CtxState | null; onClose: () => void }) {
  const s = props.state;
  if (!s) return null;
  const target = {
    getBoundingClientRect: () => ({ x: s.x, y: s.y, left: s.x, top: s.y, right: s.x, bottom: s.y, width: 0, height: 0, toJSON: () => ({}) }) as DOMRect,
  };
  return (
    <Menu open positioning={{ target, position: 'below', align: 'start', overflowBoundaryPadding: 8 }} onOpenChange={(_, d) => !d.open && props.onClose()}>
      <MenuPopover style={{ minWidth: 240, maxHeight: 'calc(100vh - 24px)', overflowY: 'auto' }}>
        <MenuList>
          <Items items={s.items} />
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
