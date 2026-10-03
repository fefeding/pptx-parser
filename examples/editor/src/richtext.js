/**
 * 富文本编辑：contenteditable 与内部 paragraphs 模型的互转
 */
import { pxToPt, ptToPx, normalizeColor, rgbToHex } from './util.js';

/** 从计算样式读取 run 样式 */
function styleOfElement(node, stopAt, defaults) {
  let cur = node;
  const res = {
    bold: undefined, italic: undefined, underline: undefined,
    color: undefined, fontSize: undefined, fontFace: undefined
  };
  while (cur && cur.nodeType === 1 && cur !== stopAt) {
    const cs = getComputedStyle(cur);
    if (res.bold === undefined) {
      const w = cs.fontWeight;
      res.bold = w === 'bold' || w === 'bolder' || parseInt(w, 10) >= 600;
    }
    if (res.italic === undefined) res.italic = cs.fontStyle === 'italic' || cs.fontStyle === 'oblique';
    if (res.underline === undefined) {
      const d = (cs.textDecorationLine || cs.textDecoration || '');
      res.underline = d.includes('underline');
    }
    if (res.color === undefined && cs.color) {
      const m = /rgba?\(([^)]+)\)/.exec(cs.color);
      if (m) {
        const [r, g, b, a] = m[1].split(',').map((v) => parseFloat(v));
        if (!(a === 0)) res.color = rgbToHex(r || 0, g || 0, b || 0);
      }
    }
    if (res.fontSize === undefined && cs.fontSize) res.fontSize = pxToPt(parseFloat(cs.fontSize));
    if (res.fontFace === undefined && cs.fontFamily) {
      res.fontFace = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    }
    cur = cur.parentElement;
  }
  return {
    bold: !!res.bold,
    italic: !!res.italic,
    underline: !!res.underline,
    color: res.color || defaults.color,
    fontSize: Math.round((res.fontSize || defaults.fontSize) * 10) / 10,
    fontFace: res.fontFace || defaults.fontFace
  };
}

function textNodesIn(root, range) {
  const out = [];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = walker.nextNode())) {
    if (range) {
      if (!range.intersectsNode(n)) continue;
      let start = 0, end = n.nodeValue.length;
      if (n === range.startContainer) start = range.startOffset;
      if (n === range.endContainer) end = range.endOffset;
      if (end <= start && !(start === 0 && end === 0)) continue;
      out.push({ node: n, start, end });
    } else {
      out.push({ node: n, start: 0, end: n.nodeValue.length });
    }
  }
  return out;
}

/** 解析 contenteditable 内容 → paragraphs */
export function parseBody(body, defaults) {
  const paragraphs = [];
  let current = null;
  const pushCurrent = () => {
    if (current) paragraphs.push(current);
    current = null;
  };
  for (const child of Array.from(body.childNodes)) {
    if (child.nodeType === 3) {
      if (!child.nodeValue) continue;
      current = { nodes: [child] };
      pushCurrent();
    } else if (child.nodeName === 'BR') {
      pushCurrent();
      paragraphs.push({ nodes: [] });
    } else if (child.nodeType === 1) {
      if (/^(DIV|P|LI)$/.test(child.nodeName)) { pushCurrent(); current = { nodes: [child], el: child }; }
      else { if (!current) current = { nodes: [] }; current.nodes.push(child); }
    }
  }
  pushCurrent();
  if (!paragraphs.length) paragraphs.push({ nodes: [] });

  return paragraphs.map((p) => {
    const runs = [];
    const anchor = p.el || body;
    const collect = (node) => {
      if (node.nodeType === 3) {
        if (!node.nodeValue) return;
        const st = styleOfElement(node.parentElement || anchor, body.parentNode, defaults);
        const last = runs[runs.length - 1];
        if (last && sameStyle(last, st)) last.text += node.nodeValue;
        else runs.push({ text: node.nodeValue, ...st });
      } else if (node.nodeType === 1) {
        for (const c of Array.from(node.childNodes)) collect(c);
      }
    };
    if (p.nodes.length) p.nodes.forEach(collect);
    else runs.push({ text: '', ...defaults });
    return { runs };
  });
}

function sameStyle(a, b) {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline &&
    a.color === b.color && a.fontSize === b.fontSize && a.fontFace === b.fontFace;
}

/** 把样式写入 span */
function paintSpan(span, styles) {
  if (styles.bold !== undefined) span.style.fontWeight = styles.bold ? '700' : '400';
  if (styles.italic !== undefined) span.style.fontStyle = styles.italic ? 'italic' : 'normal';
  if (styles.underline !== undefined) span.style.textDecoration = styles.underline ? 'underline' : 'none';
  if (styles.color !== undefined) span.style.color = normalizeColor(styles.color) || '#000';
  if (styles.fontSize !== undefined) span.style.fontSize = `${ptToPx(styles.fontSize)}px`;
  if (styles.fontFace !== undefined) span.style.fontFamily = styles.fontFace;
  return span;
}

/**
 * 对选区应用内联样式；返回 true 表示已处理（需回写模型）
 */
export function applyInlineToSelection(body, styles) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return false;
  const range = sel.getRangeAt(0);
  if (!body.contains(range.commonAncestorContainer)) return false;
  if (range.collapsed) return false;

  const items = textNodesIn(body, range);
  if (!items.length) return false;
  const spans = [];
  for (const { node, start, end } of items) {
    if (!node.nodeValue) continue;
    let target = node;
    const len = node.nodeValue.length;
    if (end < len) node.splitText(end);
    if (start > 0) target = node.splitText(start);
    const span = paintSpan(document.createElement('span'), styles);
    target.parentNode.insertBefore(span, target);
    span.appendChild(target);
    spans.push(span);
  }
  if (!spans.length) return false;
  // 合并相邻同样式 span
  for (const span of spans) {
    const prev = span.previousElementSibling;
    if (prev && prev.tagName === 'SPAN' && prev.getAttribute('style') === span.getAttribute('style')) {
      while (span.firstChild) prev.appendChild(span.firstChild);
      span.remove();
    }
  }
  const alive = spans.filter((s) => s.parentNode);
  if (alive.length) {
    const r = document.createRange();
    r.setStartBefore(alive[0]);
    r.setEndAfter(alive[alive.length - 1]);
    sel.removeAllRanges();
    sel.addRange(r);
  }
  return true;
}

/** 当前光标/选区所在段落索引（相对 body 的子节点） */
export function selectionParagraphIndexes(body) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return [];
  const range = sel.getRangeAt(0);
  if (!body.contains(range.commonAncestorContainer)) return [];
  const blocks = Array.from(body.children);
  const out = [];
  blocks.forEach((b, i) => {
    if (range.intersectsNode(b)) out.push(i);
  });
  if (!out.length) {
    let n = range.commonAncestorContainer;
    while (n && n !== body) {
      const idx = blocks.indexOf(n);
      if (idx >= 0) { out.push(idx); break; }
      n = n.parentNode;
    }
  }
  return out.length ? out : blocks.map((_, i) => i);
}

/** 读取光标处的样式状态（用于工具栏高亮） */
export function queryState(body, defaults) {
  const sel = window.getSelection();
  if (sel && sel.rangeCount) {
    const range = sel.getRangeAt(0);
    if (body.contains(range.commonAncestorContainer) && !range.collapsed) {
      const items = textNodesIn(body, range);
      if (items.length) {
        const first = styleOfElement(items[0].node.parentElement || body, body.parentNode, defaults);
        return first;
      }
    }
  }
  return { ...defaults };
}

/** 进入编辑：可指定点击位置放置光标 */
export function focusBody(body, point) {
  body.contentEditable = 'true';
  body.focus();
  const sel = window.getSelection();
  const range = document.createRange();
  if (point && document.caretRangeFromPoint) {
    const r = document.caretRangeFromPoint(point.x, point.y);
    if (r && body.contains(r.startContainer)) {
      sel.removeAllRanges();
      sel.addRange(r);
      return;
    }
  }
  range.selectNodeContents(body);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

export function blurBody(body) {
  if (!body) return;
  body.contentEditable = 'false';
  try { body.blur(); } catch { /* ignore */ }
}
