/**
 * 元素类导出
 * 对齐 PPTXjs 的完整解析能力
 */

import type { BaseElement } from './BaseElement';
import { ShapeElement } from './ShapeElement';
import { ImageElement } from './ImageElement';
import { OleElement } from './OleElement';
import { ChartElement } from './ChartElement';
import { DiagramElement } from './DiagramElement';
import { TagsElement } from './TagsElement';
import { TableElement } from './TableElement';
import { GroupElement } from './GroupElement';
import { SlideElement, PptxDocument } from './SlideElement';
import { LayoutElement, PlaceholderElement } from './LayoutElement';
import { MasterElement } from './MasterElement';
import { NotesMasterElement, NotesSlideElement } from './NotesElement';
import { DocumentElement, createDocument } from './DocumentElement';
import { getFirstChildByTagNS } from '../utils';
import { NS } from '../constants';

/**
 * 元素类导出
 */
export { BaseElement } from './BaseElement';
export { ShapeElement, type BulletStyle, type TextRun } from './ShapeElement';
export { ImageElement, type MediaType, type VideoInfo, type AudioInfo } from './ImageElement';
export { OleElement } from './OleElement';
export { ChartElement, type ChartData, type ChartSeries, type ChartDataPoint } from './ChartElement';
export { TableElement, type TableCell, type TableRow, type TableStyle } from './TableElement';
export { DiagramElement, type DiagramData, type DiagramShape } from './DiagramElement';
export { GroupElement } from './GroupElement';
export { SlideElement, PptxDocument } from './SlideElement';
export { DocumentElement, createDocument } from './DocumentElement';
// 布局和母版相关
export { LayoutElement, PlaceholderElement } from './LayoutElement';
export { MasterElement } from './MasterElement';
// 备注相关
export { NotesMasterElement, NotesSlideElement } from './NotesElement';
// 标签和扩展
export { TagsElement } from './TagsElement';

/**
 * 从解析后的数据创建元素实例（用于toHTML渲染）
 */
export function createElementFromData(data: any, relsMap: Record<string, any> = {}): BaseElement | null {
  if (!data || !data.type) return null;

  switch (data.type) {
    case 'shape':
    case 'text': {
      const element = new ShapeElement(data.id, data.type, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'image': {
      const element = new ImageElement(data.id, data.rect, data.src || '', data.relId || '', data.props || {}, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'ole': {
      const element = new OleElement(data.id, data.rect, data.progId, data.relId, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'chart': {
      const element = new ChartElement(data.id, 'chart', data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'table': {
      const element = new TableElement(data.id, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'diagram': {
      const element = new DiagramElement(data.id, data.rect, data.content, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    case 'group': {
      const children = (data.children || []).map((child: any) => createElementFromData(child, relsMap)).filter(Boolean) as BaseElement[];
      const element = new GroupElement(data.id, 'group', data.rect, children, data.props, relsMap);
      Object.assign(element, data);
      return element;
    }
    default:
      return null;
  }
}

/**
 * 创建元素的工厂函数（从XML节点）
 */
export function createElementFromNode(node: Element, relsMap: Record<string, any>): BaseElement | null {
  const tagName = node.tagName;

  if (tagName === 'p:sp' || tagName === 'sp') {
    return ShapeElement.fromNode(node, relsMap);
  } else if (tagName === 'p:pic' || tagName === 'pic') {
    return ImageElement.fromNode(node, relsMap);
  } else if (tagName === 'p:graphicFrame' || tagName === 'graphicFrame') {
    // 判断是OLE、图表、表格还是图解
    const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
    const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

    if (!graphicData) return null;

    const uri = graphicData.getAttribute('uri') || '';

    if (uri.includes('oleObject')) {
      return OleElement.fromNode(node, relsMap);
    } else if (uri.includes('chart')) {
      return ChartElement.fromNode(node, relsMap);
    } else if (uri.includes('diagram')) {
      return DiagramElement.fromNode(node, relsMap);
    } else if (uri.includes('table')) {
      return TableElement.fromNode(node, relsMap);
    }
    return null;
  } else if (tagName === 'p:grpSp' || tagName === 'grpSp') {
    return GroupElement.fromNode(node, relsMap);
  }

  return null;
}

// 重新导入BaseElement类型，避免循环依赖
import type { BaseElement as BaseElementType } from './BaseElement';
