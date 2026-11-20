// Small DOM/utility helpers shared by the webview renderer and the app shell.
// Colour, geometry and unit primitives live in @fefeding/ppt-parser and are re-exported from
// there by the renderer, so nothing numeric is duplicated here.

/** DOM builder: `h('div', { class, onclick }, child, 'text')`. */
export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props?: Record<string, any> | null,
  ...children: any[]
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (props) {
    for (const [k, v] of Object.entries(props)) {
      if (v == null || v === false) continue;
      if (k === 'class') node.className = v;
      // Accept both a style object (renderer) and a style string (app shell)
      else if (k === 'style') {
        if (typeof v === 'object') Object.assign(node.style, v);
        else node.setAttribute('style', String(v));
      } else if (k === 'html') node.innerHTML = v;
      else if (k === 'text') node.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function')
        node.addEventListener(k.slice(2).toLowerCase(), v);
      else if (k === 'dataset') Object.assign(node.dataset, v);
      else node.setAttribute(k, v === true ? '' : String(v));
    }
  }
  // `.flat()` lets callers pass either variadic children or a single array
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

export const clamp = (v: number, a: number, b: number): number => Math.min(b, Math.max(a, v));