// 端到端校验「版式/母版装饰形状（inherited）」与「主题字体引用」两项两端同步契约。
//
// 背景：编辑器（pptxToStandard → render.js）与预览端（pptxToHtml）曾是两条零共享链路。
// 本测试固化同步后必须成立的三条契约：
//   1. 解析：版式/母版的**非占位符**形状（背景底纹、斜切块、装饰条）被导入并标记 inherited，
//      否则编辑器画布大片空白（实测某样例封面页像素差 97.9%）；
//   2. 生成：inherited 元素必须被跳过，不得写进 slide 的 spTree（否则导出后与版式重复渲染），
//      同时 layout/master 部件要原样保留装饰（否则导出后装饰彻底消失）；
//   3. 字体：主题字体引用 `+mn-ea`/`+mj-lt` 展开为实际字体名，a:ea（中文实际字形）单独保留，
//      否则中文回退系统默认字形（宋体/黑体），与预览端字体不符。
import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { fileURLToPath } from 'url';
import { jsonToPptx, pptxToStandard } from '../src/index.ts';
import { docFromPptx, docToPptx } from '../examples/editor/src/model.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLE = path.resolve(__dirname, '../examples/企业微信应用介绍.pptx');

function toArrayBuffer(u8: Uint8Array): ArrayBuffer {
  return u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength);
}

/** 递归收集元素与其 group 后代 */
function walkEls(els: any[] | undefined, out: any[] = []): any[] {
  for (const el of els || []) {
    out.push(el);
    if (el.type === 'group') walkEls(el.children, out);
  }
  return out;
}

async function shapeNames(zip: JSZip, part: string): Promise<string[]> {
  const f = zip.file(part);
  if (!f) return [];
  const s = await f.async('string');
  return [...s.matchAll(/<p:cNvPr[^>]*name="([^"]*)"/g)].map((m) => m[1]);
}

const DECO_RE = /矩形 9|椭圆 7|圆顶角/;

describe('版式/母版装饰形状与主题字体（两端同步契约）', () => {
  it('解析端导入版式非占位符装饰形状并标记 inherited，且不含占位符', async () => {
    const src = await pptxToStandard(toArrayBuffer(fs.readFileSync(SAMPLE)));

    const deco = walkEls(src.slides[0].elements).filter((e) => e.inherited);
    // 该样例版式含：背景矩形、椭圆装饰、两个圆顶角斜光条
    expect(deco.length).toBeGreaterThanOrEqual(4);
    // 装饰形状位于元素列表最前（底层），不能压住 slide 自身内容
    expect(src.slides[0].elements.slice(0, deco.length).every((e) => e.inherited)).toBe(true);

    // 占位符只提供空壳，绝不能作为装饰导入（否则与 slide 同名占位符重复渲染）
    expect(deco.some((e) => String(e.name || '').includes('占位符'))).toBe(false);
  }, 60000);

  it('版式声明 showMasterSp="0" 时母版装饰不导入；未声明时按 OOXML 语义照常导入', async () => {
    const src = await pptxToStandard(toArrayBuffer(fs.readFileSync(SAMPLE)));

    // 仅 slide1 的版式带 showMasterSp="0"，其余各页未声明（等价于 "1"，应显示母版形状）
    const MASTER_DECO = '任意多边形';   // 母版里的 custGeom 右上斜切块
    const slide1Has = walkEls(src.slides[0].elements)
      .some((e) => e.inherited && String(e.name || '').includes(MASTER_DECO));
    const otherHas = src.slides.slice(1).some((s) => walkEls(s.elements)
      .some((e) => e.inherited && String(e.name || '').includes(MASTER_DECO)));

    expect(slide1Has).toBe(false);   // 隐藏母版形状 → 两端都不绘制
    expect(otherHas).toBe(true);     // 未隐藏 → 与预览端 getBackground 行为一致
  }, 60000);

  it('生成端跳过 inherited：导出后 slide 的 spTree 不含装饰，且版式装饰原样保留', async () => {
    const buf = fs.readFileSync(SAMPLE);
    const src = await pptxToStandard(toArrayBuffer(buf));

    const srcZip = await JSZip.loadAsync(buf);
    const layoutPart = 'ppt/slideLayouts/slideLayout22.xml';
    const srcLayoutNames = await shapeNames(srcZip, layoutPart);
    const decoInLayout = srcLayoutNames.filter((n) => DECO_RE.test(n));
    expect(decoInLayout.length).toBeGreaterThanOrEqual(3);

    // 走完整编辑器链路：导入 → 导出 → 生成 pptx
    const edoc = docFromPptx(JSON.parse(JSON.stringify(src)) as any);
    // 编辑器内部模型必须保留 inherited（否则 jsonToPptx 无从跳过）
    const edocDeco = walkEls(edoc.slides[0].elements).filter((e: any) => e.inherited);
    expect(edocDeco.length).toBeGreaterThanOrEqual(4);
    // 默认锁定，防止误拖/误删破坏与版式的一致性
    expect(edocDeco.every((e: any) => e.locked)).toBe(true);

    const outBuf = await jsonToPptx(docToPptx(edoc as any));
    const outZip = await JSZip.loadAsync(Buffer.from(outBuf));

    // (a) slide 自身不含任何版式装饰形状
    const slide1Names = await shapeNames(outZip, 'ppt/slides/slide1.xml');
    expect(slide1Names.filter((n) => DECO_RE.test(n))).toEqual([]);
    expect(slide1Names.some((n) => n.includes('标题'))).toBe(true);

    // (b) 版式装饰仍由 layout 提供（导出后打开不会丢装饰）
    const outLayoutNames = await shapeNames(outZip, layoutPart);
    for (const n of decoInLayout) expect(outLayoutNames).toContain(n);
  }, 90000);

  it('主题字体引用展开为实际字体名，并单独保留东亚字体', async () => {
    const src = await pptxToStandard(toArrayBuffer(fs.readFileSync(SAMPLE)));

    // 该样例主题 majorFont/minorFont：latin=Arial、ea=微软雅黑；
    // 源 run 写的是 `a:latin="+mn-lt"` / `a:ea="+mn-ea"` 这类引用
    const runs = walkEls(src.slides[0].elements)
      .flatMap((e) => e.paragraphs || [])
      .flatMap((p: any) => p.runs || []);

    expect(runs.length).toBeGreaterThan(0);
    for (const r of runs) {
      // 未解析的主题引用会让渲染端当作不存在的字体名而回退系统默认字形
      if (r.fontFace) expect(String(r.fontFace)).not.toMatch(/^\+(mj|mn)-/);
      if (r.fontFaceEa) expect(String(r.fontFaceEa)).not.toMatch(/^\+(mj|mn)-/);
    }
    // 中文公司名所在 run：西文 Arial + 中文微软雅黑，两者都要有
    const corp = runs.find((r) => String(r.text || '').includes('金腾'));
    expect(corp).toBeTruthy();
    expect(corp.fontFace).toBe('Arial');
    expect(corp.fontFaceEa).toBe('微软雅黑');
  }, 60000);
});
