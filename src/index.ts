/**
 * PPT解析序列化核心库 - 统一导出入口
 * 极简调用：一行导入，全部能力可用
 */
import { PptParseUtils, parsePptx, serializePptx } from './core';
import { emu2px, px2emu, getAttrs, parseRels, parseMetadata } from './utils';
import { NS, EMU_PER_INCH, PIXELS_PER_INCH } from './constants';
import type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';
import type { PptxParseResult, SlideParseResult, ParseOptions } from './core/types';
import type { HtmlRenderOptions } from './elements/DocumentElement';

/**
 * @deprecated 请使用 DocumentElement.toHTML() 或 createDocument() 替代
 * 保留这些函数仅用于向后兼容
 */
export const slide2HTML = (slide: any, options?: HtmlRenderOptions) => {
  const { createDocument } = require('./elements');
  const doc = createDocument({ slides: [slide], props: { width: 960, height: 540 } });
  return doc.getSlide(0)?.toHTML() || '';
};

export const ppt2HTML = (result: PptxParseResult, options?: HtmlRenderOptions) => {
  const { createDocument } = require('./elements');
  const doc = createDocument(result);
  return doc.slides.map((s: any) => s.toHTML());
};

export const ppt2HTMLDocument = (result: PptxParseResult, options?: HtmlRenderOptions) => {
  const { createDocument } = require('./elements');
  const doc = createDocument(result);
  return doc.toHTML({ ...options, withNavigation: false });
};

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
  DocumentElement,
  createDocument,
  createElementFromNode,
  createElementFromData,
  // 布局和母版相关
  LayoutElement,
  PlaceholderElement,
  MasterElement,
  // 备注相关
  NotesMasterElement,
  NotesSlideElement,
  // 标签和扩展
  TagsElement
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

// 导出 HTML 渲染类型
export type { HtmlRenderOptions };

// 导出增强类型
export type { PptxParseResult, SlideParseResult, ParseOptions } from './core/types';
export { NS, EMU_PER_INCH, PIXELS_PER_INCH } from './constants';
export { emu2px, px2emu, getAttrs, parseRels, parseMetadata } from './utils';
export type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';

