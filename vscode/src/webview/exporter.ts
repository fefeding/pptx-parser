// Offscreen slide rasterizer for "export as image / PDF".
//
// Why render offscreen from the editor model instead of rasterizing the visible canvas:
//  - the preview DOM (`pptxToHtml` output) carries an attribute whose name is a bare quote, which
//    makes html-to-image's foreignObject XML unparseable — preview slides cannot be rasterized
//    directly (verified: the generated SVG fails `<img>` decode on every slide);
//  - the visible canvas is zoomed and carries selection overlays, while export wants a fixed
//    natural size.
//
// Library choice (measured, same edit-mode DOM, pixel diff vs a real browser screenshot):
// html-to-image (foreignObject, browser-rendered) mean channel diff 2.91 vs html2canvas 4.30, and
// it serializes <canvas> content (ECharts charts) correctly. html2canvas re-draws the DOM itself
// and drifts on text/gradient rendering.
import { toPng } from 'html-to-image';
import { h } from './util';
import { renderSlideInto, disposeDetachedCharts, chartInstances } from './render';

const nextFrame = () => new Promise<void>((r) => requestAnimationFrame(() => r()));

export type ExportedSlide = { index: number; png: string };

/**
 * Render the given slides at natural size and rasterize each to a PNG data URL.
 * `pixelRatio` scales the output resolution (2 → 2560×1440 for a 16:9 deck).
 */
export async function exportSlidesToPng(opts: {
  doc: any;
  indices: number[];
  pixelRatio?: number;
  onProgress?: (done: number, total: number) => void;
}): Promise<ExportedSlide[]> {
  const { doc, indices } = opts;
  const pixelRatio = opts.pixelRatio ?? 2;
  const W = doc.slideSize.width;
  const H = doc.slideSize.height;
  const host = h('div', {
    style: {
      position: 'fixed', left: '-100000px', top: '0', margin: '0',
      width: `${W}px`, height: `${H}px`, zIndex: '-1', pointerEvents: 'none'
    }
  });
  document.body.appendChild(host);
  const out: ExportedSlide[] = [];
  try {
    for (let n = 0; n < indices.length; n++) {
      const slide = doc.slides[indices[n]];
      if (!slide) continue;
      renderSlideInto(host, slide, doc, {});
      // Charts init on the next frames (enqueueEchart defers to rAF); switch off their entry
      // animation so the raster shows the final frame instead of a mid-flight one.
      await nextFrame();
      await nextFrame();
      for (const inst of chartInstances()) {
        try { inst.setOption({ animation: false }); } catch { /* instance may be mid-dispose */ }
      }
      await nextFrame();
      // Chromium usually draws a cloned <video>'s poster inside foreignObject, but that is not
      // guaranteed behaviour — swap in an <img> of the poster so the frame always shows. The host
      // is rebuilt per slide, so no restore is needed.
      for (const v of Array.from(host.querySelectorAll('video'))) {
        if (!v.poster) continue;
        const img = document.createElement('img');
        img.src = v.poster;
        const cs = getComputedStyle(v);
        img.style.cssText = `position:${cs.position};left:${cs.left};top:${cs.top};width:${cs.width};height:${cs.height};object-fit:contain;display:block;background:#000`;
        v.style.visibility = 'hidden';
        v.parentNode?.insertBefore(img, v);
      }
      // The clone inlines the host's own computed style — including the offscreen `position: fixed;
      // left: -100000px` — which would draw the content outside the SVG viewport and produce a
      // fully transparent PNG. Reset the positioning on the clone only.
      const png = await toPng(host, {
        pixelRatio,
        width: W,
        height: H,
        style: { position: 'static', left: '0', top: '0', margin: '0' }
      });
      out.push({ index: indices[n], png });
      opts.onProgress?.(n + 1, indices.length);
      host.innerHTML = '';
      disposeDetachedCharts();
    }
  } finally {
    host.remove();
    disposeDetachedCharts();
  }
  return out;
}

/** dataURL → base64 payload (the host writes bytes, and the webview protocol carries base64). */
export function base64Of(dataUrl: string): string {
  const i = dataUrl.indexOf(',');
  return i >= 0 ? dataUrl.slice(i + 1) : dataUrl;
}
