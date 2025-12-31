/**
 * PPT解析序列化核心库 - 统一导出入口
 * 极简调用：一行导入，全部能力可用
 * ✅ 修复：去掉.ts后缀，符合TS+Bundler规范，零报错
 */
import { PptParseUtils, parsePptx, serializePptx } from './core';
import { PptParseUtilsExtended } from './core-extended';
import type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType } from './types';

// 默认导出：命名空间调用（推荐，极简友好）
const PptParserCore = {
  utils: PptParseUtils,
  utilsExtended: PptParseUtilsExtended,
  parse: parsePptx,
  serialize: serializePptx
};

export default PptParserCore;

// 按需导出：解构调用
export { PptParseUtils, PptParseUtilsExtended, parsePptx, serializePptx };

// 导出所有类型，TS项目友好
export type { PptDocument, PptSlide, PptElement, PptRect, PptStyle, PptNodeType };