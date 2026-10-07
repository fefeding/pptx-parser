import fs from 'node:fs';
import path from 'path';
import { PNG } from '/Users/jiamao/project/github/pptx-parser/node_modules/pngjs/mod.js';
import pixelmatch from '/Users/jiamao/project/github/pptx-parser/node_modules/pixelmatch/index.js';
const dir = '/Users/jiamao/project/github/pptx-parser/test/shot7';
const a = PNG.sync.read(fs.readFileSync(path.join(dir, 'editor7.png')));
const b = PNG.sync.read(fs.readFileSync(path.join(dir, 'preview7.png')));
function diffWithShift(shiftX, shiftY) {
  const w = a.width, h = a.height;
  const imgA = Buffer.alloc(w * h * 4);
  const imgB = Buffer.alloc(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const srcX = Math.max(0, Math.min(w - 1, x - shiftX));
      const srcY = Math.max(0, Math.min(h - 1, y - shiftY));
      const ai = (y * w + x) * 4;
      const bi = (srcY * w + srcX) * 4;
      imgA[ai] = a.data[ai]; imgA[ai+1] = a.data[ai+1]; imgA[ai+2] = a.data[ai+2]; imgA[ai+3] = a.data[ai+3];
      imgB[ai] = b.data[bi]; imgB[ai+1] = b.data[bi+1]; imgB[ai+2] = b.data[bi+2]; imgB[ai+3] = b.data[bi+3];
    }
  }
  return pixelmatch(imgA, imgB, null, w, h, { threshold: 0.1 }) / (w * h);
}
for (let dy = -2; dy <= 2; dy++) {
  for (let dx = -2; dx <= 2; dx++) {
    const d = diffWithShift(dx, dy);
    console.log(`shift(${dx.toString().padStart(2,' ')},${dy.toString().padStart(2,' ')}) diff=${(d*100).toFixed(2)}%`);
  }
}
