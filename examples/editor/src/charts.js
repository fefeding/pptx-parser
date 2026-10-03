/**
 * 轻量 SVG 图表渲染器（编辑器内预览用）
 * 支持：柱状/条形/堆积、折线、面积、饼图、环形、散点、雷达
 */
import { escapeHtml } from './util.js';

const DEFAULT_PALETTE = ['#1A73E8', '#4285F4', '#34A853', '#FBBC04', '#EA4335', '#8430CE'];

function niceScale(min, max, ticks = 5) {
  if (!isFinite(min) || !isFinite(max)) return { min: 0, max: 1, step: 0.2 };
  if (min === max) { max = min + 1; }
  const span = max - min;
  const rawStep = span / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(rawStep)));
  const norm = rawStep / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
  return {
    min: Math.floor(min / step) * step,
    max: Math.ceil(max / step) * step,
    step
  };
}

function fmtNum(v) {
  if (Math.abs(v) >= 10000) return (v / 1000).toFixed(0) + 'k';
  if (Number.isInteger(v)) return String(v);
  return String(Math.round(v * 100) / 100);
}

export function renderChartSVG(el, opts = {}) {
  const W = opts.width || el.width || 640;
  const H = opts.height || el.height || 380;
  const palette = opts.palette && opts.palette.length ? opts.palette : DEFAULT_PALETTE;
  const [baseType, variant] = String(el.chartType || 'barChart').split('|');
  const cats = (el.categories || []).map(String);
  const series = (el.series || []).filter((s) => s);
  const fontScale = Math.max(0.7, Math.min(1.4, H / 380));
  const fs = Math.round(11 * fontScale);
  const textColor = opts.textColor || '#5F6368';

  const hasTitle = !!el.title;
  const hasLegend = !!el.legend && series.length > 0 && baseType !== 'pieChart' && baseType !== 'doughnutChart';
  const padTop = hasTitle ? Math.round(26 * fontScale) : 8;
  const padBottom = hasLegend ? Math.round(24 * fontScale) : Math.round(20 * fontScale);
  const parts = [];

  if (hasTitle) {
    parts.push(`<text x="${W / 2}" y="${Math.round(17 * fontScale)}" text-anchor="middle" font-size="${Math.round(13 * fontScale)}" font-weight="600" fill="${opts.titleColor || '#202124'}">${escapeHtml(el.title)}</text>`);
  }

  if (baseType === 'pieChart' || baseType === 'doughnutChart') {
    const ser = series[0] || { values: [] };
    const values = (ser.values || []).map((v) => Math.max(0, Number(v) || 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2;
    const cy = padTop + (H - padTop - padBottom) / 2;
    const r = Math.max(10, Math.min(W, H - padTop - padBottom) / 2 - 12);
    const inner = baseType === 'doughnutChart' ? r * ((el.holeSize ?? 50) / 100) : 0;
    let start = -Math.PI / 2;
    values.forEach((v, i) => {
      const ang = (v / total) * Math.PI * 2;
      const end = start + ang;
      if (v <= 0) { start = end; return; }
      const color = (ser.color && i === 0 && !el.varyColors) ? ser.color : palette[i % palette.length];
      const x1 = cx + r * Math.cos(start), y1 = cy + r * Math.sin(start);
      const x2 = cx + r * Math.cos(end), y2 = cy + r * Math.sin(end);
      const large = ang > Math.PI ? 1 : 0;
      if (inner > 0) {
        const ix2 = cx + inner * Math.cos(end), iy2 = cy + inner * Math.sin(end);
        const ix1 = cx + inner * Math.cos(start), iy1 = cy + inner * Math.sin(start);
        parts.push(`<path d="M${x1} ${y1}A${r} ${r} 0 ${large} 1 ${x2} ${y2}L${ix2} ${iy2}A${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1}Z" fill="${color}" stroke="#fff" stroke-width="1"/>`);
      } else {
        parts.push(`<path d="M${cx} ${cy}L${x1} ${y1}A${r} ${r} 0 ${large} 1 ${x2} ${y2}Z" fill="${color}" stroke="#fff" stroke-width="1"/>`);
      }
      if (el.dataLabels) {
        const mid = start + ang / 2;
        const lr = inner > 0 ? (r + inner) / 2 : r * 0.68;
        parts.push(`<text x="${cx + lr * Math.cos(mid)}" y="${cy + lr * Math.sin(mid)}" text-anchor="middle" dominant-baseline="middle" font-size="${fs}" fill="#fff">${escapeHtml(String(Math.round((v / total) * 100)) + '%')}</text>`);
      }
      start = end;
    });
    if (hasLegend) {
      const items = cats.map((c, i) => {
        const x = 6 + (i % Math.max(1, Math.floor(W / 90))) * 90;
        const y = H - 8 - Math.floor(i / Math.max(1, Math.floor(W / 90))) * 16;
        return `<rect x="${x}" y="${y - 8}" width="9" height="9" rx="2" fill="${palette[i % palette.length]}"/><text x="${x + 13}" y="${y}" font-size="${fs}" fill="${textColor}">${escapeHtml(c)}</text>`;
      });
      parts.push(...items);
    }
    return wrap(W, H, parts);
  }

  // ---- 直角坐标类 ----
  const isBar = baseType === 'barChart' && variant === 'bar';
  const leftPad = Math.round(38 * fontScale);
  const rightPad = 10;
  const plotX = leftPad;
  const plotY = padTop + 6;
  const plotW = Math.max(20, W - leftPad - rightPad);
  const plotH = Math.max(20, H - plotY - padBottom);

  // 数值范围
  let min = Infinity, max = -Infinity;
  const stacked = variant === 'stacked';
  if (stacked) {
    for (let i = 0; i < cats.length; i++) {
      let sum = 0;
      for (const s of series) sum += Number(s.values?.[i]) || 0;
      min = Math.min(min, Math.min(0, sum)); max = Math.max(max, sum);
    }
  } else {
    for (const s of series) for (const v of s.values || []) {
      const n = Number(v) || 0;
      min = Math.min(min, n); max = Math.max(max, n);
    }
  }
  if (!isFinite(min)) { min = 0; max = 1; }
  if (min > 0) min = 0;
  const scale = niceScale(min, max, 5);
  const vToY = (v) => plotY + plotH - ((v - scale.min) / (scale.max - scale.min)) * plotH;
  const vToX = (v) => plotX + ((v - scale.min) / (scale.max - scale.min)) * plotW;

  // 网格 + 数值轴刻度
  for (let v = scale.min; v <= scale.max + 1e-9; v += scale.step) {
    const y = vToY(v);
    parts.push(`<line x1="${plotX}" y1="${y}" x2="${plotX + plotW}" y2="${y}" stroke="#E8EAED" stroke-width="1"/>`);
    parts.push(`<text x="${plotX - 6}" y="${y + fs / 3}" text-anchor="end" font-size="${fs}" fill="${textColor}">${fmtNum(Math.round(v * 100) / 100)}</text>`);
  }
  parts.push(`<line x1="${plotX}" y1="${plotY + plotH}" x2="${plotX + plotW}" y2="${plotY + plotH}" stroke="#BDC1C6" stroke-width="1"/>`);

  const n = Math.max(1, cats.length);
  const band = (isBar ? plotH : plotW) / n;
  const colors = series.map((s, i) => s.color || palette[i % palette.length]);

  if (isBar) {
    cats.forEach((c, i) => {
      const y0 = plotY + i * band;
      const per = band * 0.62 / Math.max(1, series.length);
      series.forEach((s, j) => {
        const v = Number(s.values?.[i]) || 0;
        const x = stacked ? plotX : vToX(0);
        const w = Math.max(0, vToX(v) - vToX(0));
        const yy = stacked ? y0 + band * 0.19 : y0 + band * 0.19 + per * j;
        parts.push(`<rect x="${Math.min(x, x + w)}" y="${yy}" width="${Math.abs(w)}" height="${stacked ? band * 0.62 : per}" fill="${colors[j]}" rx="1"/>`);
        if (el.dataLabels && v !== 0) {
          parts.push(`<text x="${x + w + 4}" y="${yy + (stacked ? band * 0.31 : per / 2) + fs / 3}" font-size="${fs}" fill="${textColor}">${fmtNum(v)}</text>`);
        }
      });
      parts.push(`<text x="${plotX - 6}" y="${y0 + band / 2 + fs / 3}" text-anchor="end" font-size="${fs}" fill="${textColor}">${escapeHtml(c)}</text>`);
    });
  } else {
    // 类目轴标签
    cats.forEach((c, i) => {
      const x = plotX + band * (i + 0.5);
      parts.push(`<text x="${x}" y="${plotY + plotH + Math.round(14 * fontScale)}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${escapeHtml(c)}</text>`);
    });

    if (baseType === 'barChart') {
      cats.forEach((c, i) => {
        const groupX = plotX + i * band;
        const per = band * 0.7 / Math.max(1, series.length);
        let accBase = 0;
        series.forEach((s, j) => {
          const v = Number(s.values?.[i]) || 0;
          const y = vToY(stacked ? accBase + v : v);
          const h = Math.max(0, vToY(0) - y);
          const x = stacked ? groupX + band * 0.15 : groupX + band * 0.15 + per * j;
          parts.push(`<rect x="${x}" y="${y}" width="${stacked ? band * 0.7 : per}" height="${h}" fill="${colors[j]}" rx="1"/>`);
          if (el.dataLabels && v !== 0) {
            parts.push(`<text x="${x + (stacked ? band * 0.35 : per / 2)}" y="${y - 4}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${fmtNum(v)}</text>`);
          }
          accBase += v;
        });
      });
    } else if (baseType === 'lineChart' || baseType === 'areaChart') {
      const pts = series.map((s) => cats.map((_, i) => {
        const x = plotX + band * (i + 0.5);
        const y = vToY(Number(s.values?.[i]) || 0);
        return [x, y];
      }));
      if (baseType === 'areaChart') {
        pts.forEach((p, j) => {
          if (!p.length) return;
          const d = `M${p[0][0]} ${vToY(0)}` + p.map(([x, y]) => `L${x} ${y}`).join('') + `L${p[p.length - 1][0]} ${vToY(0)}Z`;
          parts.push(`<path d="${d}" fill="${colors[j]}" fill-opacity="0.28"/>`);
        });
      }
      pts.forEach((p, j) => {
        if (!p.length) return;
        let d = p.length === 1 ? '' : `M${p[0][0]} ${p[0][1]}`;
        if (el.smooth && p.length > 2) {
          for (let i = 1; i < p.length; i++) {
            const [x0, y0] = p[i - 1], [x1, y1] = p[i];
            const cx = (x0 + x1) / 2;
            d += `C${cx} ${y0} ${cx} ${y1} ${x1} ${y1}`;
          }
        } else {
          for (let i = 1; i < p.length; i++) d += `L${p[i][0]} ${p[i][1]}`;
        }
        parts.push(`<path d="${d}" fill="none" stroke="${colors[j]}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`);
        if (el.marker !== false) {
          p.forEach(([x, y]) => parts.push(`<circle cx="${x}" cy="${y}" r="3" fill="#fff" stroke="${colors[j]}" stroke-width="2"/>`));
        }
        if (el.dataLabels) {
          p.forEach(([x, y], i) => parts.push(`<text x="${x}" y="${y - 8}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${fmtNum(Number(series[j].values?.[i]) || 0)}</text>`));
        }
      });
    } else if (baseType === 'scatterChart') {
      series.forEach((s, j) => {
        (s.values || []).forEach((v, i) => {
          const x = plotX + band * (i + 0.5);
          const y = vToY(Number(v) || 0);
          parts.push(`<circle cx="${x}" cy="${y}" r="4" fill="${colors[j]}" fill-opacity="0.8"/>`);
          if (el.dataLabels) parts.push(`<text x="${x}" y="${y - 8}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${fmtNum(Number(v) || 0)}</text>`);
        });
      });
    } else if (baseType === 'radarChart') {
      const cx = plotX + plotW / 2, cy = plotY + plotH / 2;
      const rad = Math.min(plotW, plotH) / 2 - 6;
      const ang = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
      for (let ring = 1; ring <= 4; ring++) {
        const rr = (rad * ring) / 4;
        const p = cats.map((_, i) => `${cx + rr * Math.cos(ang(i))},${cy + rr * Math.sin(ang(i))}`).join(' ');
        parts.push(`<polygon points="${p}" fill="none" stroke="#E8EAED"/>`);
      }
      cats.forEach((c, i) => {
        const x = cx + rad * Math.cos(ang(i)), y = cy + rad * Math.sin(ang(i));
        parts.push(`<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#E8EAED"/>`);
        parts.push(`<text x="${cx + (rad + 12) * Math.cos(ang(i))}" y="${cy + (rad + 12) * Math.sin(ang(i)) + fs / 3}" text-anchor="middle" font-size="${fs}" fill="${textColor}">${escapeHtml(c)}</text>`);
      });
      series.forEach((s, j) => {
        const p = cats.map((_, i) => {
          const v = Number(s.values?.[i]) || 0;
          const rr = rad * Math.max(0, Math.min(1, (v - scale.min) / (scale.max - scale.min || 1)));
          return `${cx + rr * Math.cos(ang(i))},${cy + rr * Math.sin(ang(i))}`;
        }).join(' ');
        parts.push(`<polygon points="${p}" fill="${colors[j]}" fill-opacity="0.22" stroke="${colors[j]}" stroke-width="2"/>`);
      });
    }
  }

  if (hasLegend) {
    let lx = plotX;
    const ly = H - Math.round(8 * fontScale);
    series.forEach((s, j) => {
      const name = s.name || `系列 ${j + 1}`;
      const w = name.length * fs * 0.62 + 26;
      if (lx + w > W - 4) return;
      parts.push(`<rect x="${lx}" y="${ly - fs}" width="9" height="9" rx="2" fill="${colors[j]}"/>`);
      parts.push(`<text x="${lx + 13}" y="${ly - fs / 4}" font-size="${fs}" fill="${textColor}">${escapeHtml(name)}</text>`);
      lx += w;
    });
  }

  return wrap(W, H, parts);
}

function wrap(W, H, parts) {
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">${parts.join('')}</svg>`;
}
