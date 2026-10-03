/**
 * 导入 / 导出：JSON → PPTX 文件、PPTX 文件 → 编辑器文档
 */
import { jsonToPptx, pptxToStandard } from '../../../dist/ppt-parser.browser.js';
import { docToPptx, docFromPptx } from './model.js';
import { store } from './store.js';
import { toast } from './util.js';

function safeName(s) {
  return (s || 'presentation').replace(/[\\/:*?"<>|]/g, '_').slice(0, 60);
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ======================= 导出 ======================= */
export async function exportPptx() {
  try {
    const pptxDoc = docToPptx(store.doc);
    const blob = await jsonToPptx(pptxDoc, { outputType: 'blob' });
    downloadBlob(blob, (safeName(store.doc.title) || 'presentation') + '.pptx');
    toast('已导出 PPTX', 'ok');
  } catch (err) {
    console.error(err);
    toast('导出失败：' + (err.message || err), 'err');
  }
}

export function saveJson() {
  const data = JSON.stringify(store.doc, null, 2);
  downloadBlob(new Blob([data], { type: 'application/json' }), (safeName(store.doc.title) || 'presentation') + '.json');
  toast('已保存 JSON', 'ok');
}

/* ======================= 导入 ======================= */
export async function importPptxFile(file) {
  const buf = await file.arrayBuffer();
  try {
    const standard = await pptxToStandard(buf);
    const doc = docFromPptx(standard);
    store.setDoc(doc, { noHistory: true });
    store.fitted = true;
    toast('导入成功', 'ok');
    return doc;
  } catch (err) {
    console.error(err);
    toast('导入失败：' + (err.message || err), 'err');
    throw err;
  }
}

export function loadJsonFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const doc = docFromPptx(JSON.parse(reader.result));
      store.setDoc(doc, { noHistory: true });
      store.fitted = true;
      toast('已载入 JSON', 'ok');
    } catch (err) {
      toast('JSON 解析失败', 'err');
    }
  };
  reader.readAsText(file);
}
