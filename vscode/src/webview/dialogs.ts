// Modal dialogs for the complex element types (tables, charts, shapes, media) plus the
// shortcut reference and a styled replacement for `window.confirm`.
//
// This module is deliberately model-agnostic: it only builds DOM and reports the user's intent
// back through callbacks, so `main.ts` stays the single owner of the store/actions.
import { h } from './util';
import { t, data, kbd } from './i18n';
import type { MenuItem } from './menu';
import { SHAPES, CHART_TYPES } from '@fefeding/ppt-parser';

let activeModal: HTMLElement | null = null;

export function hasOpenModal(): boolean {
  return !!activeModal;
}

/** Close the current dialog, if any. Bound to Escape and to the backdrop. */
export function closeModal(): void {
  activeModal?.remove();
  activeModal = null;
  document.removeEventListener('keydown', onModalKey, true);
}

/** Escape cancels, Enter confirms — Office dialog behaviour. */
function onModalKey(e: KeyboardEvent): void {
  if (!activeModal) return;
  if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    closeModal();
    return;
  }
  if (e.key === 'Enter' && !(e.target as HTMLElement)?.closest?.('.dlg-table, textarea')) {
    const ok = activeModal.querySelector('.dlg-actions .primary') as HTMLElement | null;
    if (ok) {
      e.preventDefault();
      e.stopPropagation();
      ok.click();
    }
  }
}

// ---------- tiny helpers ----------
function input(type: string, value: string, cls = 'dlg-in'): HTMLInputElement {
  const el = document.createElement('input');
  el.type = type;
  el.value = value;
  el.className = cls;
  return el;
}
function select(options: { value: string; name: string }[], value: string, cls = 'dlg-in'): HTMLSelectElement {
  const el = document.createElement('select');
  el.className = cls;
  for (const o of options) {
    const opt = document.createElement('option');
    opt.value = o.value;
    opt.textContent = o.name;
    if (o.value === value) opt.selected = true;
    el.append(opt);
  }
  return el;
}
function intOf(v: string, min: number, max: number, def: number): number {
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? Math.min(max, Math.max(min, n)) : def;
}
function btn(label: string, onclick: () => void, cls = 'btn'): HTMLElement {
  return h('button', { class: cls, type: 'button', onclick }, [label]);
}
function row(nodes: HTMLElement[], cls = 'dlg-row'): HTMLElement {
  return h('div', { class: cls }, nodes);
}
function labelled(label: string, node: HTMLElement): HTMLElement {
  return h('label', { class: 'dlg-line' }, [h('span', {}, [label]), node]);
}
function checkBox(label: string, value: boolean, onChange: (v: boolean) => void): HTMLElement {
  const box = input('checkbox', '', 'dlg-check');
  box.checked = value;
  box.addEventListener('change', () => onChange(box.checked));
  return h('label', { class: 'dlg-check-row' }, [box, h('span', {}, [label])]);
}

// ---------- dialog shell ----------
function openDialog(title: string, build: (close: () => void) => (HTMLElement | null)[]): void {
  closeModal();
  const inner = h('div', { class: 'dlg' }, [
    h('div', { class: 'dlg-head' }, [
      h('span', {}, [title]),
      h('button', { class: 'dlg-x', type: 'button', title: t('common.close'), onclick: closeModal }, ['✕'])
    ])
  ]);
  const modal = h('div', { class: 'modal' }, [inner]);
  modal.addEventListener('pointerdown', (e) => {
    if (e.target === modal) closeModal();
  });
  document.body.append(modal);
  activeModal = modal;
  const body = h('div', { class: 'dlg-body' });
  build(closeModal).forEach((n) => n && body.append(n));
  inner.append(body);
  document.addEventListener('keydown', onModalKey, true);
}

// ---------- confirm (replaces window.confirm, which blocks the webview) ----------
export function openConfirm(opts: {
  message: string;
  detail?: string;
  okLabel?: string;
  danger?: boolean;
}): Promise<boolean> {
  return new Promise((resolve) => {
    let done = false;
    const finish = (v: boolean, close: () => void) => {
      if (done) return;
      done = true;
      closeModal();
      resolve(v);
    };
    openDialog(t('dlg.confirmTitle'), (close) => [
      h('p', { class: 'dlg-msg' }, [opts.message]),
      opts.detail ? h('p', { class: 'dlg-detail' }, [opts.detail]) : null,
      row(
        [
          btn(t('common.cancel'), () => finish(false, close)),
          btn(opts.okLabel || t('common.ok'), () => finish(true, close), 'btn primary' + (opts.danger ? ' danger' : ''))
        ],
        'dlg-actions'
      )
    ]);
  });
}

// ---------- table: insert ----------
export function openTableSizeDialog(api: { create: (rows: number, cols: number) => void }): void {
  openDialog(t('dlg.table.insert'), (close) => {
    const rowsIn = input('number', '3');
    rowsIn.min = '1';
    rowsIn.max = '30';
    const colsIn = input('number', '3');
    colsIn.min = '1';
    colsIn.max = '20';
    return [
      labelled(t('dlg.rows'), rowsIn),
      labelled(t('dlg.cols'), colsIn),
      row([
        btn(t('common.cancel'), close),
        btn(t('common.insert'), () => {
          api.create(intOf(rowsIn.value, 1, 30, 3), intOf(colsIn.value, 1, 20, 3));
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- table: edit ----------
export function openTableDialog(
  el: any,
  api: { update: (patch: any) => void; resize: (rows: number, cols: number) => void; reopen: () => void }
): void {
  openDialog(t('dlg.table.edit'), (close) => {
    const rows: any[] = Array.isArray(el.rows) ? el.rows : [];
    const cols = Math.max(1, ...rows.map((r: any) => (r.cells || []).length));
    const draft: string[][] = rows.map((r: any) =>
      Array.from({ length: cols }, (_, j) => String((((r.cells || [])[j] || {}) as any).text ?? ''))
    );

    const rowsIn = input('number', String(rows.length || 1), 'dlg-cell');
    const colsIn = input('number', String(cols), 'dlg-cell');

    const grid = h('div', { class: 'dlg-table' });
    draft.forEach((r, i) =>
      r.forEach((v, j) => {
        const cell = input('text', v, 'dlg-cell');
        cell.addEventListener('input', () => {
          draft[i][j] = cell.value;
        });
        grid.append(cell);
      })
    );

    return [
      row([
        labelled(t('dlg.rows'), rowsIn),
        labelled(t('dlg.cols'), colsIn),
        btn(t('dlg.resize'), () => {
          api.resize(intOf(rowsIn.value, 1, 30, 1), intOf(colsIn.value, 1, 20, 1));
          // The action rebuilds this table's rows, so reopen against the fresh model
          close();
          api.reopen();
        })
      ]),
      grid,
      row([
        btn(t('common.cancel'), close),
        btn(t('common.save'), () => {
          const next = draft.map((r, i) => ({
            height: (rows[i] || {}).height,
            cells: r.map((text, j) => ({ ...((rows[i] || {}).cells || [])[j], text }))
          }));
          api.update({ rows: next });
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- chart: edit ----------
export function openChartDialog(el: any, api: { update: (patch: any) => void }): void {
  openDialog(t('dlg.chart.edit'), (close) => {
    const draft: any = {
      chartType: el.chartType || 'barChart',
      title: el.title || '',
      legend: el.legend !== false,
      dataLabels: !!el.dataLabels,
      smooth: !!el.smooth,
      marker: !!el.marker,
      categories: Array.isArray(el.categories) ? el.categories.slice() : [],
      series: Array.isArray(el.series) ? JSON.parse(JSON.stringify(el.series)) : []
    };

    const typeSel = select(CHART_TYPES.map((c: any) => ({ value: c.value, name: data('chartType', c.value, c.name) })), draft.chartType);
    typeSel.addEventListener('change', () => {
      draft.chartType = typeSel.value;
    });
    const titleIn = input('text', String(draft.title ?? ''));
    titleIn.addEventListener('input', () => {
      draft.title = titleIn.value;
    });

    const catWrap = h('div', { class: 'dlg-col' });
    const serWrap = h('div', { class: 'dlg-col' });

    const renderSeries = () => {
      serWrap.innerHTML = '';
      draft.series.forEach((s: any, si: number) => {
        const nameIn = input('text', String(s.name ?? ''), 'dlg-cell');
        nameIn.addEventListener('input', () => {
          s.name = nameIn.value;
        });
        const vals = row([]);
        draft.categories.forEach((_: any, ci: number) => {
          const v = input('number', String((s.values || [])[ci] ?? 0), 'dlg-cell');
          v.addEventListener('input', () => {
            s.values = s.values || [];
            s.values[ci] = parseFloat(v.value) || 0;
          });
          vals.append(v);
        });
        serWrap.append(
          h('div', { class: 'dlg-block' }, [
            row([nameIn, btn('✕', () => { draft.series.splice(si, 1); renderSeries(); })]),
            vals
          ])
        );
      });
      serWrap.append(
        btn(t('dlg.addSeries'), () => {
          draft.series.push({ name: t('dlg.newSeries', draft.series.length + 1), values: draft.categories.map(() => 0) });
          renderSeries();
        })
      );
    };

    const renderCats = () => {
      catWrap.innerHTML = '';
      draft.categories.forEach((c: string, i: number) => {
        const inp = input('text', String(c ?? ''), 'dlg-cell');
        inp.addEventListener('input', () => {
          draft.categories[i] = inp.value;
        });
        catWrap.append(row([inp, btn('✕', () => {
          draft.categories.splice(i, 1);
          draft.series.forEach((s: any) => (s.values || []).splice(i, 1));
          renderCats();
          renderSeries();
        })]));
      });
      catWrap.append(
        btn(t('dlg.addCategory'), () => {
          draft.categories.push('');
          draft.series.forEach((s: any) => (s.values = s.values || []).push(0));
          renderCats();
          renderSeries();
        })
      );
    };

    renderCats();
    renderSeries();

    return [
      labelled(t('dlg.type'), typeSel),
      labelled(t('dlg.title'), titleIn),
      row([
        checkBox(t('insp.legend'), draft.legend, (v) => (draft.legend = v)),
        checkBox(t('insp.dataLabels'), draft.dataLabels, (v) => (draft.dataLabels = v)),
        checkBox(t('dlg.smooth'), draft.smooth, (v) => (draft.smooth = v)),
        checkBox(t('dlg.markers'), draft.marker, (v) => (draft.marker = v))
      ]),
      h('div', { class: 'dlg-split' }, [
        h('div', {}, [h('b', {}, [t('dlg.categories')]), catWrap]),
        h('div', {}, [h('b', {}, [t('dlg.series')]), serWrap])
      ]),
      row([
        btn(t('common.cancel'), close),
        btn(t('common.save'), () => {
          api.update(draft);
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- shapes ----------
const SHAPE_GROUPS = ['basic', 'line', 'arrow', 'flowchart'] as const;

/**
 * Localized shape name. Entries are keyed by group because several groups reuse one prstGeom
 * (e.g. `straightConnector1` is both a connector and an arrow), and the arrow group tells its
 * four entries apart only by list index.
 */
export function shapeLabel(shape: any, indexInGroup: number): string {
  const key = shape.group === 'arrow' ? `data.shape.arrow.${indexInGroup}` : `data.shape.${shape.group}.${shape.type}`;
  return t(key, shape.name);
}

function shapeGlyph(shape: any, cls: string): string {
  // 线条/连接符类用描边而非填充，箭头才清晰可见
  const isLine = shape.stroke || (shape.type && /connector/.test(shape.type));
  const attrs = isLine
    ? 'fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"'
    : 'fill="currentColor"';
  return `<svg viewBox="0 0 24 24" width="${cls}" height="${cls}" aria-hidden="true"><path d="${shape.d}" ${attrs}/></svg>`;
}

export function openShapePicker(api: { addShape: (type: string, opt?: any) => void }): void {
  openDialog(t('dlg.shape.insert'), (close) => {
    const nodes: (HTMLElement | null)[] = [];
    for (const g of SHAPE_GROUPS) {
      const list = (SHAPES as any[]).filter((s) => s.group === g);
      if (!list.length) continue;
      nodes.push(h('div', { class: 'dlg-group-title' }, [t(`shapeGroup.${g}`, g)]));
      const wrap = h('div', { class: 'dlg-shapes' });
      list.forEach((s, i) => {
        const item = document.createElement('button');
        item.className = 'shape-item';
        item.title = shapeLabel(s, i);
        item.innerHTML = `${shapeGlyph(s, '26')}<span>${shapeLabel(s, i)}</span>`;
        item.addEventListener('click', () => {
          api.addShape(s.type, s.opt);
          close();
        });
        wrap.append(item);
      });
      nodes.push(wrap);
    }
    return nodes;
  });
}

/** Compact shape list for the ribbon's split-button dropdown. */
export function shapeMenuItems(api: { addShape: (type: string, opt?: any) => void }): MenuItem[] {
  const items: MenuItem[] = [];
  for (const g of SHAPE_GROUPS) {
    const list = (SHAPES as any[]).filter((s) => s.group === g);
    if (!list.length) continue;
    items.push({ type: 'label', label: t(`shapeGroup.${g}`, g) });
    list.forEach((s, i) => {
      items.push({ label: shapeLabel(s, i), onClick: () => api.addShape(s.type, s.opt) });
    });
  }
  return items;
}

// ---------- SmartArt: edit the text of each node ----------
/**
 * SmartArt 文字编辑对话框。
 * 图示的文字存在两处部件里（数据部件 + 缓存绘图部件），保存时由 parser 的
 * updateDiagramTexts 同步改写，这里只负责收集用户输入。
 */
export function openDiagramTextDialog(
  texts: string[],
  api: { save: (texts: string[]) => void }
): void {
  openDialog(t('dlg.diagram.text'), (close) => {
    const inputs: HTMLInputElement[] = texts.map((v) => input('text', v, 'dlg-cell'));
    const list = h('div', { class: 'dlg-col' });
    inputs.forEach((inp, i) => {
      list.append(h('div', { class: 'dlg-line' }, [h('span', {}, [t('dlg.diagram.node', String(i + 1))]), inp]));
    });
    return [
      h('p', { class: 'dlg-msg' }, [t('dlg.diagram.hint')]),
      list,
      row([
        btn(t('common.cancel'), close),
        btn(t('common.save'), () => {
          api.save(inputs.map((x) => x.value));
          close();
        }, 'btn primary')
      ], 'dlg-actions')
    ];
  });
}

// ---------- media ----------
export function openMediaDialog(api: { addMedia: (type: 'video' | 'audio', data: string, ext: string) => void }): void {
  openDialog(t('dlg.media.insert'), (close) => {
    const pick = (type: 'video' | 'audio', accept: string) => {
      const file = input('file', '', '');
      file.style.display = 'none';
      file.accept = accept;
      file.addEventListener('change', () => {
        const f = file.files?.[0];
        if (!f) return;
        const rd = new FileReader();
        rd.onload = () => {
          api.addMedia(type, String(rd.result), (f.name.split('.').pop() || 'mp4').toLowerCase());
          close();
        };
        rd.readAsDataURL(f);
      });
      document.body.append(file);
      file.click();
    };
    return [
      row([
        btn(t('dlg.video'), () => pick('video', 'video/*,.mp4,.m4v,.mov,.webm,.avi')),
        btn(t('dlg.audio'), () => pick('audio', 'audio/*,.mp3,.m4a,.wav,.aac,.ogg,.wma'))
      ]),
      row([btn(t('common.cancel'), close)], 'dlg-actions')
    ];
  });
}

// ---------- keyboard shortcuts help ----------
type ShortcutGroup = { title: string; rows: [string, string][] };

export function openShortcuts(): void {
  const groups: ShortcutGroup[] = [
    {
      title: t('sc.group.general'),
      rows: [
        [kbd('f5'), t('sc.present')],
        [kbd('shift+f5'), t('sc.presentHere')],
        [kbd('mod+z'), t('sc.undo')],
        [kbd('mod+y'), t('sc.redo')],
        [kbd('mod+a'), t('sc.selectAll')],
        [kbd('mod+c'), t('sc.copy')],
        [kbd('mod+x'), t('sc.cut')],
        [kbd('mod+v'), t('sc.paste')],
        [kbd('mod+d'), t('sc.duplicate')],
        [kbd('esc'), t('sc.escape')],
        ['?', t('sc.help')]
      ]
    },
    {
      title: t('sc.group.selection'),
      rows: [
        [t('sc.click'), t('sc.selectElement')],
        [t('sc.shiftClick'), t('sc.addToSelection')],
        [t('sc.drag'), t('sc.moveElement')],
        [t('sc.corner'), t('sc.resizeElement')],
        [t('sc.topHandle'), t('sc.rotateElement')],
        [kbd('tab'), t('sc.nextElement')],
        [kbd('arrowleft') + ' / ' + kbd('arrowright'), t('sc.nudge')],
        [kbd('del'), t('sc.delete')],
        [kbd('mod+g'), t('sc.group')],
        [kbd('mod+shift+g'), t('sc.ungroup')]
      ]
    },
    {
      title: t('sc.group.text'),
      rows: [
        [t('sc.dblText'), t('sc.editTextInline')],
        [kbd('enter'), t('sc.editTextInline')],
        [kbd('mod+b') + ' / ' + kbd('mod+i') + ' / ' + kbd('mod+u'), t('sc.boldItalic')]
      ]
    },
    {
      title: t('sc.group.slides'),
      rows: [
        [t('sc.dblGroup'), t('sc.enterGroupEdit')],
        [t('sc.dblLine'), t('sc.editVertices')]
      ]
    }
  ];
  openDialog(t('dlg.shortcuts'), (close) => {
    const wrap = h('div', { class: 'shortcuts' });
    for (const g of groups) {
      wrap.append(h('h4', {}, [g.title]));
      const table = h('table', {});
      for (const [k, d] of g.rows) {
        table.append(h('tr', {}, [h('td', { class: 'key' }, [k]), h('td', {}, [d])]));
      }
      wrap.append(table);
    }
    return [wrap, h('div', { class: 'dlg-actions' }, [btn(t('common.close'), close, 'btn primary')])];
  });
}