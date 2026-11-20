/**
 * PDF barrel：只暴露「生成 PDF」这一条链（PPTX -> PDF 直排）。
 *
 * 这里有意不包含 PDF 解析（pdfjs-dist worker + cMap/wasm）与编辑能力：
 * 转换引擎只需要生成方向，无需把整条 PDF 解析链拉进依赖图。
 */
export { generatePdf as jsonToPdf } from './generate';
export * from './types';
