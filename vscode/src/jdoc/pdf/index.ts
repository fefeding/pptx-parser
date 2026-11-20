/**
 * 精简版 PDF barrel（原样移植自 jdoc 的 src/pdf/index.ts 时被裁剪）。
 *
 * 本插件只需要「生成 PDF」（PPTX -> PDF 直排）；jdoc 的完整 barrel 还 re-export
 * 了 ./parse（pdfjs-dist worker + cMap/wasm 资源）与 ./edit，那会把 PDF 解析链
 * 整条拉进扩展宿主，而这里根本没有「PDF 作为源」的转换。
 *
 * jdoc 里的 jsonToPdf 就是 generatePdf 的一层薄包装，签名完全一致，因此直接
 * 别名导出即可，p_pptx-pdf.ts 的 `import { jsonToPdf } from '../pdf/index'`
 * 无需改动（保持上游文件逐字可比对：convert/pptx-pdf.ts）。
 */
export { generatePdf as jsonToPdf } from './generate';
export * from './types';
