/**
 * 导入BaseElement类型
 */
import type { BaseElement } from './BaseElement';
import { ShapeElement } from './ShapeElement';
import { ImageElement } from './ImageElement';
import { OleElement } from './OleElement';
import { ChartElement } from './ChartElement';
import { GroupElement } from './GroupElement';
import { SlideElement, PptxDocument } from './SlideElement';
import { getFirstChildByTagNS, getAttrSafe } from '../utils';
import { NS } from '../constants';

/**
 * 元素类导出
 */
export { BaseElement } from './BaseElement';
export { ShapeElement } from './ShapeElement';
export { ImageElement } from './ImageElement';
export { OleElement } from './OleElement';
export { ChartElement } from './ChartElement';
export { GroupElement } from './GroupElement';
export { SlideElement, PptxDocument } from './SlideElement';

/**
 * 从解析后的数据创建元素实例（用于toHTML渲染）
 */
export function createElementFromData(data: any): BaseElement | null {
  if (!data || !data.type) return null;

  switch (data.type) {
    case 'shape':
    case 'text': {
      const element = new ShapeElement(data.id, data.type, data.rect, data.content, data.props, {});
      Object.assign(element, data);
      return element;
    }
    case 'image': {
      const element = new ImageElement(data.id, data.rect, data.src, data.relId, data.props, {});
      Object.assign(element, data);
      return element;
    }
    case 'ole': {
      const element = new OleElement(data.id, data.rect, data.progId, data.relId, data.props, {});
      Object.assign(element, data);
      return element;
    }
    case 'chart': {
      const element = new ChartElement(data.id, data.rect, data.chartType, data.relId, data.props, {});
      Object.assign(element, data);
      return element;
    }
    case 'group': {
      const children = (data.children || []).map((child: any) => createElementFromData(child)).filter(Boolean);
      const element = new GroupElement(data.id, data.rect, children, data.props, {});
      Object.assign(element, data);
      return element;
    }
    default:
      return null;
  }
}

/**
 * 创建元素的工厂函数
 */
export function createElementFromNode(node: Element, relsMap: Record<string, any>): BaseElement | null {
  const tagName = node.tagName;

  if (tagName === 'p:sp' || tagName === 'sp') {
    return ShapeElement.fromNode(node, relsMap);
  } else if (tagName === 'p:pic' || tagName === 'pic') {
    return ImageElement.fromNode(node, relsMap);
  } else if (tagName === 'p:graphicFrame' || tagName === 'graphicFrame') {
    // 判断是OLE还是图表
    const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
    const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

    if (graphicData) {
      const uri = graphicData.getAttribute('uri') || '';

      if (uri.includes('oleObject')) {
        return OleElement.fromNode(node, relsMap);
      } else if (uri.includes('chart')) {
        return ChartElement.fromNode(node, relsMap);
      }
    }
    return null;
  } else if (tagName === 'p:grpSp' || tagName === 'grpSp') {
    return GroupElement.fromNode(node, relsMap);
  }

  return null;
}

