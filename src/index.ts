/**
 * PPT解析序列化核心库 - 统一导出入口
 * 极简调用：一行导入，全部能力可用
 * ✅ 修复：去掉.ts后缀，符合TS+Bundler规范，零报错
 */
import { PptParseUtils, parsePptx, serializePptx } from './core';
import { PptParseUtilsExtended } from './core-extended';
import { parsePptxEnhanced, parseSlide as parseSlideEnhanced } from './parser-enhanced';
import { parseSlideWithElements, createPptxDocument } from './parseSlideEnhanced';
import { emu2px, px2emu, getAttrs, parseRels, parseMetadata } from './utils';
import { NS, EMU_PER_INCH, PIXELS_PER_INCH } from './constants';
import type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';
import type { PptxParseResult, SlideParseResult, ParseOptions } from './types-enhanced';

// 元素类导出
export {
  BaseElement,
  ShapeElement,
  ImageElement,
  OleElement,
  ChartElement,
  GroupElement,
  SlideElement,
  PptxDocument,
  createElementFromNode
} from './elements';

// 增强解析导出
export { parseSlideWithElements, createPptxDocument } from './parseSlideEnhanced';

// 默认导出：命名空间调用（推荐，极简友好）
const PptParserCore = {
  utils: PptParseUtils,
  utilsExtended: PptParseUtilsExtended,
  parse: parsePptx,
  parseEnhanced: parsePptxEnhanced,
  parseSlide: parseSlideEnhanced,
  serialize: serializePptx
};

export default PptParserCore;

// 按需导出：解构调用
export { PptParseUtils, PptParseUtilsExtended, parsePptx, serializePptx };

// 增强版导出
export { parsePptxEnhanced, parseSlideEnhanced as parseSlide };

// 工具函数导出
export { emu2px, px2emu, getAttrs, parseRels, parseMetadata };

// 常量导出
export { NS, EMU_PER_INCH, PIXELS_PER_INCH };

// 导出所有类型，TS项目友好
export type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType };

// 导出增强类型
export type { PptxParseResult, SlideParseResult, ParseOptions };