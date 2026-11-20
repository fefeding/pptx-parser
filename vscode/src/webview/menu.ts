// One popup implementation for every menu in the editor: canvas / slide-panel context menus
// and the ribbon's dropdowns (New Slide ▾, Layout ▾, shape palette, colour picker …).
//
// Office-like details worth keeping: icon column, right-aligned shortcut hint, a section label
// row, keyboard navigation (↑ ↓ ↵ Esc) and edge clamping.
import { h } from './util';
import { icon } from './icons';

export type MenuItem =
  | { type?: 'item'; label?: string; icon?: string; accel?: string; disabled?: boolean; checked?: boolean; onClick?: () => void }
  | { type: 'sep' }
  | { type: 'label'; label: string };

let popup: HTMLElement | null = null;
let closeHandlers: Array<() => void> = [];

function closePopup(): void {
  if (!popup) return;
  popup.remove();
  popup = null;
  for (const fn of closeHandlers.splice(0)) fn();
}

function ensurePopup(): HTMLElement {
  if (!popup) {
    popup = h('div', { class: 'menu', role: 'menu' });
    document.body.append(popup);
  }
  return popup;
}

function place(x: number, y: number, align: 'left' | 'right' = 'left'): void {
  const el = popup!;
  const rect = el.getBoundingClientRect();
  let left = align === 'right' ? x - rect.width : x;
  let top = y;
  if (left + rect.width > window.innerWidth - 6) left = Math.max(6, window.innerWidth - rect.width - 6);
  if (top + rect.height > window.innerHeight - 6) top = Math.max(6, y - rect.height);
  if (top < 6) top = 6;
  el.style.left = left + 'px';
  el.style.top = top + 'px';
}

/**
 * Show a popup menu. `onClose` runs once the popup goes away, so callers that need a
 * "was it dismissed?" signal (split buttons) can wire it here.
 */
export function openMenu(
  items: MenuItem[],
  x: number,
  y: number,
  opts: { align?: 'left' | 'right'; minWidth?: number; onClose?: () => void } = {}
): void {
  const wasOpen = !!popup;
  closePopup();
  const el = ensurePopup();
  el.style.minWidth = (opts.minWidth || 168) + 'px';

  for (const it of items) {
    if (it.type === 'sep') {
      el.append(h('div', { class: 'menu-sep' }));
      continue;
    }
    if (it.type === 'label') {
      el.append(h('div', { class: 'menu-caption' }, [it.label]));
      continue;
    }
    const row = h(
      'div',
      {
        class: 'menu-item' + (it.disabled ? ' disabled' : ''),
        role: 'menuitem',
        tabindex: '-1'
      },
      [
        h('span', { class: 'menu-ico' }, [it.checked ? '✓' : it.icon ? icon(it.icon) : '']),
        h('span', { class: 'menu-text' }, [it.label || '']),
        it.accel ? h('span', { class: 'menu-accel' }, [it.accel]) : null
      ]
    );
    if (!it.disabled) {
      row.addEventListener('click', () => {
        closePopup();
        it.onClick?.();
      });
      row.addEventListener('mouseenter', () => row.classList.add('active'));
      row.addEventListener('mouseleave', () => row.classList.remove('active'));
    }
    el.append(row);
  }

  document.addEventListener('pointerdown', onDocDown, true);
  window.addEventListener('blur', closePopup);
  window.addEventListener('resize', closePopup);
  if (opts.onClose) closeHandlers.push(opts.onClose);
  place(x, y, opts.align);
  // Re-place once the first paint has real dimensions (needed for right/bottom alignment).
  place(x, y, opts.align);
  if (wasOpen) return;
  const first = el.querySelector('.menu-item:not(.disabled)') as HTMLElement | null;
  first?.focus({ preventScroll: true });
}

function onDocDown(e: PointerEvent): void {
  if (popup && !popup.contains(e.target as Node)) closePopup();
}

/** Open a popup anchored under a ribbon control (dropdown look, aligned with it). */
export function openMenuAt(
  anchor: HTMLElement,
  items: MenuItem[],
  opts: { align?: 'left' | 'right'; minWidth?: number } = {}
): void {
  const r = anchor.getBoundingClientRect();
  // Align right edges when the anchor is close to the window's right edge, like Office does.
  const align = opts.align || (r.right > window.innerWidth - 180 ? 'right' : 'left');
  openMenu(items, align === 'right' ? r.right : r.left, r.bottom + 2, { ...opts, align });
}

/** ↑ ↓ ↵ Esc handling while a popup is open; returns true when the key was consumed. */
export function popupKeydown(e: KeyboardEvent): boolean {
  if (!popup) return false;
  if (e.key === 'Escape') {
    closePopup();
    return true;
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    const rows = Array.from(popup.querySelectorAll('.menu-item:not(.disabled)')) as HTMLElement[];
    if (!rows.length) return true;
    const cur = rows.indexOf(document.activeElement as HTMLElement);
    const next = e.key === 'ArrowDown' ? (cur + 1) % rows.length : cur <= 0 ? rows.length - 1 : cur - 1;
    rows[next].focus();
    return true;
  }
  if (e.key === 'Enter' || e.key === ' ') {
    const active = document.activeElement as HTMLElement | null;
    if (active && popup.contains(active) && active.classList.contains('menu-item')) {
      active.click();
      return true;
    }
  }
  return false;
}