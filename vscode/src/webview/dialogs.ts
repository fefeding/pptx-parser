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
/** Per-dialog "closed" callbacks, so promise-based dialogs resolve on *every* close path. */
const modalCloseHandlers = new WeakMap<HTMLElement, () => void>();

export function hasOpenModal(): boolean {
  return !!activeModal;
}

/** Close the current dialog, if any. Bound to Escape and to the backdrop. */
export function closeModal(): void {
  const modal = activeModal;
  const onClose = modal ? modalCloseHandlers.get(modal) : undefined;
  modal?.remove();
  activeModal = null;
  document.removeEventListener('keydown', onModalKey, true);
  onClose?.();
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
function openDialog(
  title: string,
  build: (close: () => void) => (HTMLElement | null)[],
  onClose?: () => void
): void {
  closeModal();
  const inner = h('div', { class: 'dlg' }, [
    h('div', { class: 'dlg-head' }, [
      h('span', {}, [title]),
      h('button', { class: 'dlg-x', type: 'button', title: t('common.close'), onclick: closeModal }, ['✕'])
    ])
  ]);
  const modal = h('div', { class: 'modal' }, [inner]);
  if (onClose) modalCloseHandlers.set(modal, onClose);
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
    const finish = (v: boolean) => {
      if (done) return;
      done = true;
      closeModal();
      resolve(v);
    };
    openDialog(
      t('dlg.confirmTitle'),
      () => [
        h('p', { class: 'dlg-msg' }, [opts.message]),
        opts.detail ? h('p', { class: 'dlg-detail' }, [opts.detail]) : null,
        row(
          [
            btn(t('common.cancel'), () => finish(false)),
            btn(opts.okLabel || t('common.ok'), () => finish(true), 'btn primary' + (opts.danger ? ' danger' : ''))
          ],
          'dlg-actions'
        )
      ],
      () => finish(false)
    );
  });
}

// ---------- prompt (replaces window.prompt, which is blocked in webviews) ----------
export function openPrompt(opts: {
  message: string;
  value?: string;
  okLabel?: string;
}): Promise<string | null> {
  return new Promise((resolve) => {
    let done = false;
    const finish = (v: string | null) => {
      if (done) return;
      done = true;
      closeModal();
      resolve(v);
    };
    const field = input('text', opts.value ?? '');
    openDialog(
      t('dlg.promptTitle'),
      () => [
        h('p', { class: 'dlg-msg' }, [opts.message]),
        field,
        row(
          [
            btn(t('common.cancel'), () => finish(null)),
            btn(opts.okLabel || t('common.ok'), () => finish(field.value), 'btn primary')
          ],
          'dlg-actions'
        )
      ],
      () => finish(null)
    );
    field.focus();
    field.select();
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
    const rows: any[] = JSON.parse(JSON.stringify(Array.isArray(el.rows) ? el.rows : []));
    const cols = Math.max(1, ...rows.map((r: any) => (r.cells || []).length));
    while (rows.length < 1) rows.push({ cells: [] });
    rows.forEach((r) => {
      if (!Array.isArray(r.cells)) r.cells = [];
      while (r.cells.length < cols) r.cells.push({});
    });
    const colWidths: number[] = Array.isArray(el.colWidths) ? el.colWidths.slice() : [];
    const rowHeights: number[] = Array.isArray(el.rowHeights) ? el.rowHeights.slice() : [];
    while (colWidths.length < cols) colWidths.push(100);
    while (rowHeights.length < rows.length) rowHeights.push(30);

    let anchor: { r: number; c: number } | null = null;
    let focus: { r: number; c: number } | null = null;
    const cells: HTMLElement[] = [];
    const colHeads: HTMLElement[] = [];
    const rowHeads: HTMLElement[] = [];
    const grid = h('div', { class: 'dlg-table' });

    const range = () => {
      if (!anchor || !focus) return null;
      return {
        r0: Math.min(anchor.r, focus.r), r1: Math.max(anchor.r, focus.r),
        c0: Math.min(anchor.c, focus.c), c1: Math.max(anchor.c, focus.c)
      };
    };
    const eachSelected = (fn: (cell: any, r: number, c: number) => void) => {
      const box = range();
      if (!box) return;
      for (let r = box.r0; r <= box.r1; r++)
        for (let c = box.c0; c <= box.c1; c++) fn(rows[r].cells[c], r, c);
    };
    const paintSel = () => {
      const box = range();
      cells.forEach((node) => {
        const r = Number((node as any).dataset.r), c = Number((node as any).dataset.c);
        node.classList.toggle('sel', !!box && r >= box.r0 && r <= box.r1 && c >= box.c0 && c <= box.c1);
      });
    };
    const setSizes = () => {
      colHeads.forEach((node, j) => (node.style.width = `${colWidths[j]}px`));
      rowHeads.forEach((node, i) => (node.style.height = `${rowHeights[i]}px`));
      cells.forEach((node) => {
        const c = Number((node as any).dataset.c);
        const cell = rows[Number((node as any).dataset.r)].cells[c];
        const span = Math.max(1, Number(cell.colSpan) || 1);
        node.style.width = `${colWidths.slice(c, c + span).reduce((a, b) => a + b, 0)}px`;
      });
    };
    const dragSize = (e: PointerEvent, kind: 'col' | 'row', index: number) => {
      e.preventDefault();
      e.stopPropagation();
      const start = kind === 'col' ? e.clientX : e.clientY;
      const base = kind === 'col' ? colWidths[index] : rowHeights[index];
      const move = (ev: PointerEvent) => {
        const delta = (kind === 'col' ? ev.clientX : ev.clientY) - start;
        if (kind === 'col') colWidths[index] = Math.max(48, base + delta);
        else rowHeights[index] = Math.max(24, base + delta);
        setSizes();
      };
      const up = () => {
        document.removeEventListener('pointermove', move);
        document.removeEventListener('pointerup', up);
      };
      document.addEventListener('pointermove', move);
      document.addEventListener('pointerup', up);
    };

    const head = h('div', { class: 'dlg-table-head' }, [h('div', { class: 'dlg-table-corner' })]);
    colWidths.forEach((_, j) => {
      const node = h('div', { class: 'dlg-table-colhead' }, [String(j + 1)]);
      node.addEventListener('pointerdown', (e: PointerEvent) => dragSize(e, 'col', j));
      colHeads.push(node);
      head.append(node);
    });
    grid.append(head);
    rows.forEach((r, i) => {
      const line = h('div', { class: 'dlg-table-line' });
      const rh = h('div', { class: 'dlg-table-rowhead' }, [String(i + 1)]);
      rh.addEventListener('pointerdown', (e: PointerEvent) => dragSize(e, 'row', i));
      rowHeads.push(rh);
      line.append(rh);
      r.cells.forEach((cell: any, j: number) => {
        const wrap = h('div', {
          class: 'dlg-table-cell',
          dataset: { r: String(i), c: String(j) }
        });
        const text = input('text', String(cell.text || ''), 'dlg-cell');
        text.addEventListener('input', () => (cell.text = text.value));
        wrap.append(text);
        wrap.addEventListener('click', (ev: MouseEvent) => {
          const at = { r: i, c: j };
          if (ev.shiftKey && anchor) focus = at;
          else { anchor = at; focus = at; }
          paintSel();
          syncStyle();
        });
        cells.push(wrap);
        line.append(wrap);
      });
      grid.append(line);
    });
    setSizes();
    paintSel();

    const fillIn = input('color', '#ffffff', 'dlg-cell');
    const borderIn = input('color', '#808080', 'dlg-cell');
    const borderW = input('number', '1', 'dlg-cell');
    const fontIn = input('color', '#000000', 'dlg-cell');
    const fontSizeIn = input('number', '14', 'dlg-cell');
    const boldIn = input('checkbox', ''), italicIn = input('checkbox', ''), underlineIn = input('checkbox', '');
    [boldIn, italicIn, underlineIn].forEach((n) => (n.className = 'dlg-check'));
    const alignIn = select([
      { value: 'left', name: t('data.align.left') }, { value: 'center', name: t('data.align.center') }, { value: 'right', name: t('data.align.right') }
    ], 'left', 'dlg-cell');
    const valignIn = select([
      { value: 'top', name: t('insp.valign') + ': top' }, { value: 'middle', name: t('insp.valign') + ': middle' }, { value: 'bottom', name: t('insp.valign') + ': bottom' }
    ], 'top', 'dlg-cell');
    const syncStyle = () => {
      const box = range();
      const cell = box ? rows[box.r0].cells[box.c0] : null;
      if (!cell) return;
      fillIn.value = /^#[0-9a-f]{6}$/i.test(String(cell.fill || '')) ? cell.fill : '#ffffff';
      borderIn.value = /^#[0-9a-f]{6}$/i.test(String((cell.borders || {}).bottom?.color || '')) ? cell.borders.bottom.color : '#808080';
      borderW.value = String((cell.borders || {}).bottom?.width ?? 1);
      fontIn.value = /^#[0-9a-f]{6}$/i.test(String(cell.color || '')) ? cell.color : '#000000';
      fontSizeIn.value = String(cell.fontSize || 14);
      boldIn.checked = !!cell.bold;
      italicIn.checked = !!cell.italic;
      underlineIn.checked = !!cell.underline;
      alignIn.value = cell.align || 'left';
      valignIn.value = cell.valign || 'top';
    };
    const applyStyle = (patch: any) => eachSelected((cell) => Object.assign(cell, patch));
    const applyBorders = (patch: any) => eachSelected((cell) => {
      const borders: any = { ...(cell.borders || {}) };
      for (const side of ['left', 'right', 'top', 'bottom']) {
        borders[side] = { ...(borders[side] || {}), ...patch };
      }
      cell.borders = borders;
    });
    fillIn.addEventListener('input', () => applyStyle({ fill: fillIn.value }));
    borderIn.addEventListener('input', () => applyBorders({ color: borderIn.value }));
    borderW.addEventListener('input', () => applyBorders({ width: Number(borderW.value) }));
    fontIn.addEventListener('input', () => applyStyle({ color: fontIn.value }));
    fontSizeIn.addEventListener('input', () => applyStyle({ fontSize: Number(fontSizeIn.value) }));
    boldIn.addEventListener('change', () => applyStyle({ bold: boldIn.checked }));
    italicIn.addEventListener('change', () => applyStyle({ italic: italicIn.checked }));
    underlineIn.addEventListener('change', () => applyStyle({ underline: underlineIn.checked }));
    alignIn.addEventListener('change', () => applyStyle({ align: alignIn.value }));
    valignIn.addEventListener('change', () => applyStyle({ valign: valignIn.value }));

    const saveAndReopen = () => {
      api.update({ rows, colWidths, rowHeights });
      close();
      api.reopen();
    };
    const merge = () => {
      const box = range();
      if (!box || (box.r0 === box.r1 && box.c0 === box.c1)) return;
      const texts: string[] = [];
      for (let r = box.r0; r <= box.r1; r++)
        for (let c = box.c0; c <= box.c1; c++) {
          const value = String(rows[r].cells[c].text || '').trim();
          if (value) texts.push(value);
        }
      const top = rows[box.r0].cells[box.c0];
      top.text = texts.join(' ');
      top.colSpan = box.c1 - box.c0 + 1;
      top.rowSpan = box.r1 - box.r0 + 1;
      for (let r = box.r0; r <= box.r1; r++)
        for (let c = box.c0; c <= box.c1; c++) {
          if (r === box.r0 && c === box.c0) continue;
          const cell = rows[r].cells[c];
          cell.hMerge = true;
          cell.vMerge = true;
          cell.text = '';
        }
      saveAndReopen();
    };
    const split = () => {
      const box = range();
      if (!box) return;
      const top = rows[box.r0].cells[box.c0];
      const cs = Math.max(1, Number(top.colSpan) || 1), rs = Math.max(1, Number(top.rowSpan) || 1);
      top.colSpan = 1;
      top.rowSpan = 1;
      for (let r = box.r0; r < Math.min(rows.length, box.r0 + rs); r++)
        for (let c = box.c0; c < Math.min(rows[r].cells.length, box.c0 + cs); c++) {
          if (r === box.r0 && c === box.c0) continue;
          const cell = rows[r].cells[c];
          cell.hMerge = false;
          cell.vMerge = false;
        }
      saveAndReopen();
    };

    const rowsIn = input('number', String(rows.length), 'dlg-cell');
    const colsIn = input('number', String(cols), 'dlg-cell');
    return [
      row([
        labelled(t('dlg.rows'), rowsIn),
        labelled(t('dlg.cols'), colsIn),
        btn(t('dlg.resize'), () => {
          api.resize(intOf(rowsIn.value, 1, 30, 1), intOf(colsIn.value, 1, 20, 1));
          close();
          api.reopen();
        })
      ]),
      h('div', { class: 'dlg-hint' }, [t('dlg.table.selectHint')]),
      grid,
      row([
        btn(t('dlg.table.merge'), merge),
        btn(t('dlg.table.split'), split)
      ]),
      row([
        labelled(t('dlg.table.fill'), fillIn),
        labelled(t('dlg.table.border'), borderIn),
        labelled(t('dlg.table.borderWidth'), borderW),
        labelled(t('dlg.table.fontColor'), fontIn),
        labelled(t('dlg.table.fontSize'), fontSizeIn)
      ]),
      row([
        labelled(t('dlg.table.bold'), boldIn),
        labelled(t('dlg.table.italic'), italicIn),
        labelled(t('dlg.table.underline'), underlineIn),
        labelled(t('insp.align'), alignIn),
        labelled(t('insp.valign'), valignIn)
      ]),
      row([
        btn(t('common.cancel'), close),
        btn(t('common.save'), () => {
          api.update({ rows, colWidths, rowHeights });
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
      legendPosition: el.legendPosition || 'right',
      dataLabels: !!el.dataLabels,
      smooth: !!el.smooth,
      marker: !!el.marker,
      axisTitles: {
        category: String(el.axisTitles?.category || ''),
        value: String(el.axisTitles?.value || '')
      },
      gridlines: { major: el.gridlines?.major !== false },
      numberFormat: String(el.numberFormat || ''),
      trendType: String(el.series?.[0]?.trendlines?.[0]?.type || 'none'),
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

    // Excel-style grid: paste tab/comma separated cells (first row = series names, first column = categories).
    const dataArea = h('textarea', { class: 'dlg-data', spellcheck: 'false' }) as HTMLTextAreaElement;
    const serializeData = () => {
      const header = ['', ...draft.series.map((s: any, i: number) => s.name || t('dlg.newSeries', i + 1))].join('\t');
      const body = draft.categories.map((c: string, i: number) =>
        [c, ...draft.series.map((s: any) => (s.values || [])[i] ?? 0)].join('\t')
      );
      return [header, ...body].join('\n');
    };
    dataArea.value = serializeData();
    const initialData = dataArea.value;
    const parseData = () => {
      const lines = dataArea.value.split(/\r?\n/).filter((line) => line.trim());
      if (!lines.length) return;
      const table = lines.map((line) => line.split('\t').map((v) => v.trim()));
      const names = table[0].slice(1);
      const body = table.slice(1);
      const old = draft.series;
      draft.categories = body.map((r) => r[0] || '');
      draft.series = names.map((name: string, i: number) => ({
        ...(old[i] || {}),
        name,
        values: body.map((r) => parseFloat(r[i + 1]) || 0)
      }));
      renderCats();
      renderSeries();
    };

    const legendPosSel = select([
      { value: 'right', name: t('dlg.chart.legendRight') },
      { value: 'left', name: t('dlg.chart.legendLeft') },
      { value: 'top', name: t('dlg.chart.legendTop') },
      { value: 'bottom', name: t('dlg.chart.legendBottom') }
    ], draft.legendPosition, 'dlg-cell');
    legendPosSel.addEventListener('change', () => (draft.legendPosition = legendPosSel.value));
    const catAxisIn = input('text', draft.axisTitles.category, 'dlg-cell');
    catAxisIn.addEventListener('input', () => (draft.axisTitles.category = catAxisIn.value));
    const valAxisIn = input('text', draft.axisTitles.value, 'dlg-cell');
    valAxisIn.addEventListener('input', () => (draft.axisTitles.value = valAxisIn.value));
    const numberIn = input('text', draft.numberFormat, 'dlg-cell');
    numberIn.addEventListener('input', () => (draft.numberFormat = numberIn.value));
    const trendSel = select([
      { value: 'none', name: t('dlg.chart.trendNone') },
      { value: 'linear', name: t('dlg.chart.trendLinear') },
      { value: 'exponential', name: t('dlg.chart.trendExp') },
      { value: 'polynomial', name: t('dlg.chart.trendPoly') }
    ], draft.trendType, 'dlg-cell');
    trendSel.addEventListener('change', () => (draft.trendType = trendSel.value));

    return [
      labelled(t('dlg.type'), typeSel),
      labelled(t('dlg.title'), titleIn),
      row([
        checkBox(t('insp.legend'), draft.legend, (v) => (draft.legend = v)),
        checkBox(t('insp.dataLabels'), draft.dataLabels, (v) => (draft.dataLabels = v)),
        checkBox(t('dlg.smooth'), draft.smooth, (v) => (draft.smooth = v)),
        checkBox(t('dlg.markers'), draft.marker, (v) => (draft.marker = v))
      ]),
      row([
        labelled(t('dlg.chart.legendPos'), legendPosSel),
        labelled(t('dlg.chart.axisCat'), catAxisIn),
        labelled(t('dlg.chart.axisVal'), valAxisIn)
      ]),
      row([
        labelled(t('dlg.chart.numberFormat'), numberIn),
        labelled(t('dlg.chart.trendline'), trendSel),
        checkBox(t('dlg.chart.gridlines'), draft.gridlines.major, (v) => (draft.gridlines.major = v))
      ]),
      h('div', { class: 'dlg-block' }, [
        h('b', {}, [t('dlg.chart.dataPaste')]),
        dataArea,
        btn(t('dlg.chart.applyData'), parseData)
      ]),
      h('div', { class: 'dlg-split' }, [
        h('div', {}, [h('b', {}, [t('dlg.categories')]), catWrap]),
        h('div', {}, [h('b', {}, [t('dlg.series')]), serWrap])
      ]),
      row([
        btn(t('common.cancel'), close),
        btn(t('common.save'), () => {
          if (dataArea.value !== initialData) parseData();
          draft.series.forEach((s: any) => {
            s.trendlines = draft.trendType === 'none'
              ? []
              : [{ type: draft.trendType, name: t('dlg.chart.trendline') }];
          });
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

// ---------- master / layout inspector (read-only + limited background/theme editing) ----------
/**
 * 母版/版式查看器。库端把原始 slideMasterN.xml / slideLayoutN.xml 以 __rawXml 无损保留，
 * 这里只做「只读展示 + 背景色/主题色的有限编辑」，不重建母版骨架（PowerPoint 校验严格）。
 */
export function openMasterViewDialog(data: {
  masters: { name: string; themeName: string; colors: { name: string; value: string }[]; background: string; layoutNames: string[]; slideUsage: number[] }[];
  themeSlots: { name: string; value: string }[];
}, api: {
  setBackground: (masterIndex: number, color: string) => void;
}): void {
  openDialog(t('dlg.master.title'), (close) => {
    const body = h('div', { class: 'dlg-master' });
    if (!data.masters.length) {
      body.append(h('p', { class: 'dlg-msg' }, [t('dlg.master.empty')]));
      return [body, row([btn(t('common.close'), close)], 'dlg-actions')];
    }

    const detail = h('div', { class: 'dlg-master-detail' });
    const swatches = (list: { name: string; value: string }[], onClick?: (v: string) => void) => {
      const wrap = h('div', { class: 'dlg-master-swatches' });
      for (const c of list) {
        const box = h('span', {
          class: 'dlg-master-swatch',
          style: `background:${c.value}`,
          title: `${c.name} ${c.value}`,
          ...(onClick ? { onclick: () => onClick(c.value) } : {})
        }, []);
        wrap.append(box);
      }
      return wrap;
    };

    const render = (mi: number) => {
      detail.innerHTML = '';
      const m = data.masters[mi];
      detail.append(h('h4', {}, [`${t('dlg.master.master')} ${mi + 1} · ${m.name}`]));
      detail.append(h('div', { class: 'dlg-line' }, [h('span', {}, [t('dlg.master.theme')]), h('span', {}, [m.themeName])]));
      detail.append(h('div', { class: 'dlg-line' }, [h('span', {}, [t('dlg.master.background')]), swatches([{ name: 'bg', value: m.background }])]));
      const bgInput = input('color', /^#[0-9a-f]{6}$/i.test(m.background) ? m.background : '#ffffff');
      detail.append(h('div', { class: 'dlg-line' }, [
        h('span', {}, [t('dlg.master.backgroundEdit')]), bgInput,
        btn(t('common.save'), () => { api.setBackground(mi, bgInput.value); close(); }, 'btn')
      ]));
      detail.append(h('h4', {}, [t('dlg.master.layouts')]));
      detail.append(h('div', { class: 'dlg-master-layouts' }, [h('span', {}, [m.layoutNames.join(' · ') || '—'])]));
      detail.append(h('h4', {}, [t('dlg.master.colors')]));
      detail.append(swatches(m.colors));
    };

    const tabs = h('div', { class: 'dlg-master-tabs' });
    data.masters.forEach((m, i) => {
      tabs.append(btn(`${i + 1}: ${m.name}`, () => render(i), 'btn'));
    });
    body.append(tabs, detail);

    const themeRow = h('div', { class: 'dlg-line' }, [
      h('span', {}, [t('dlg.master.themeColors')]),
      swatches(data.themeSlots)
    ]);
    render(0);
    return [body, themeRow, row([btn(t('common.close'), close)], 'dlg-actions')];
  });
}

// ---------- formula (OMML) ----------
/** 把用户输入的线性/Unicode 公式文本包成最小但合法的 OMML（m:oMathPara/m:oMath）。 */
function linearTextToOmml(text: string): string {
  const esc = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>${esc}</m:t></m:r></m:oMath></m:oMathPara>`;
}

/**
 * 公式编辑器：显示 OMML 的纯文本形式供直接编辑，并允许高级用户改原始 OMML XML。
 * 保存时写回 text（画布/fallback 显示）与 omml（PowerPoint 里的真实公式体）。
 */
export function openFormulaDialog(
  formula: { text?: string; omml?: string },
  api: { save: (text: string, omml: string) => void }
): void {
  openDialog(t('dlg.formula.title'), (close) => {
    const textField = h('textarea', { class: 'dlg-formula-text', rows: '3', placeholder: t('dlg.formula.placeholder') }, []) as HTMLTextAreaElement;
    textField.value = formula.text || '';
    const ommlField = h('textarea', { class: 'dlg-formula-omml', rows: '5', spellcheck: 'false' }, []) as HTMLTextAreaElement;
    ommlField.value = formula.omml || '';

    const regenerate = () => {
      ommlField.value = linearTextToOmml(textField.value);
    };
    const save = () => {
      const text = textField.value;
      let omml = ommlField.value.trim();
      if (!omml && text) omml = linearTextToOmml(text);
      api.save(text, omml);
      close();
    };
    return [
      h('p', { class: 'dlg-hint' }, [t('dlg.formula.hint')]),
      h('label', { class: 'dlg-label' }, [t('dlg.formula.text')]),
      textField,
      h('div', { class: 'row' }, [btn(t('dlg.formula.regenerate'), regenerate)]),
      h('label', { class: 'dlg-label' }, [t('dlg.formula.omml')]),
      ommlField,
      row([btn(t('common.cancel'), close), btn(t('common.save'), save, 'btn primary')], 'dlg-actions')
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
      file.addEventListener('cancel', () => file.remove());
      file.addEventListener('change', () => {
        const f = file.files?.[0];
        file.remove();
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