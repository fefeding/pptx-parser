/**
 * PPT解析序列化核心库 - 统一导出入口
 * 极简调用：一行导入，全部能力可用
 */
import { PptParseUtils, parsePptx, serializePptx } from './core';
import { emu2px, px2emu, getAttrs, parseRels, parseMetadata } from './utils';
import { NS, EMU_PER_INCH, PIXELS_PER_INCH } from './constants';
import type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';
import type { PptxParseResult, SlideParseResult, ParseOptions } from './core/types';

// 元素类导出
export {
  BaseElement,
  ShapeElement,
  ImageElement,
  OleElement,
  ChartElement,
  TableElement,
  DiagramElement,
  GroupElement,
  SlideElement,
  PptxDocument,
  createElementFromNode,
  createElementFromData
} from './elements';

// 默认导出：命名空间调用（推荐，极简友好）
const PptParserCore = {
  utils: PptParseUtils,
  parse: parsePptx,
  serialize: serializePptx
};

export default PptParserCore;

// 按需导出：解构调用
export { PptParseUtils, parsePptx, serializePptx };

// 导出增强类型
export type { PptxParseResult, SlideParseResult, ParseOptions } from './core/types';
export { NS, EMU_PER_INCH, PIXELS_PER_INCH } from './constants';
export { emu2px, px2emu, getAttrs, parseRels, parseMetadata } from './utils';
export type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';

