/**
 * 图表元素类
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr } from '../utils';
import { NS } from '../constants';
import type { ParsedChartElement, RelsMap } from '../types-enhanced';

/**
 * 图表元素类
 */
export class ChartElement extends BaseElement {
  type = 'chart' as const;

  /** 图表类型 */
  chartType?: string;

  /** 关联ID */
  relId: string;

  /**
   * 从XML节点创建图表元素
   */
  static fromNode(node: Element, relsMap: RelsMap): ChartElement | null {
    try {
      // 解析ID和名称
      const nvGraphicFramePr = getFirstChildByTagNS(node, 'nvGraphicFramePr', NS.p);
      const cNvPr = nvGraphicFramePr ? getFirstChildByTagNS(nvGraphicFramePr, 'cNvPr', NS.p) : null;

      const id = getAttrSafe(cNvPr, 'id', `chart_${Date.now()}`);
      const name = getAttrSafe(cNvPr, 'name', '');
      const hidden = getBoolAttr(cNvPr, 'hidden');

      // 解析位置尺寸（使用xfrm而不是spPr）
      const xfrm = getFirstChildByTagNS(node, 'xfrm', NS.p);
      let rect = { x: 0, y: 0, width: 0, height: 0 };

      if (xfrm) {
        const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
        const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

        if (off && ext) {
          rect.x = parseInt(off.getAttribute('x') || '0') / 914400;
          rect.y = parseInt(off.getAttribute('y') || '0') / 914400;
          rect.width = parseInt(ext.getAttribute('cx') || '0') / 914400;
          rect.height = parseInt(ext.getAttribute('cy') || '0') / 914400;
        }
      }

      const element = new ChartElement(id, rect, '', '', {}, relsMap);
      element.name = name;
      element.hidden = hidden;

      // 查找graphicData判断类型
      const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
      const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

      if (!graphicData) {
        return null;
      }

      // 查找chart节点 - 注意这里使用的namespace可能不同
      const chart = getFirstChildByTagNS(graphicData, 'chart', NS.c);
      if (!chart) {
        return null;
      }

      const chartType = chart.getAttribute('type') || 'unknown';
      const relId = chart.getAttributeNS(NS.r, 'id') || chart.getAttribute('r:id') || '';

      element.chartType = chartType;
      element.relId = relId;

      element.content = {
        chartType: 'unknown',
        data: []
      };

      element.props = {
        relId,
        chartType
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse chart element:', error);
      return null;
    }
  }

  /**
   * 构造函数
   */
  constructor(
    id: string,
    rect: { x: number; y: number; width: number; height: number },
    chartType: string,
    relId: string,
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    super(id, 'chart', rect, props, relsMap);
    this.chartType = chartType;
    this.relId = relId;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const innerStyle = [
      `width: 100%`,
      `height: 100%`,
      `display: flex`,
      `align-items: center`,
      `justify-content: center`,
      `background-color: #fafafa`,
      `border: 1px solid #ddd`,
      `color: #333`,
      `font-size: 14px`,
      `text-align: center`
    ].join('; ');

    const label = this.chartType ? `${this.chartType} Chart` : 'Chart';

    return `<div style="${style}">
      <div style="${innerStyle}">
        ${label}
      </div>
    </div>`;
  }

  /**
   * 转换为ParsedChartElement格式
   */
  toParsedElement(): ParsedChartElement {
    return {
      id: this.id,
      type: 'chart',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      relId: this.relId,
      chartType: this.chartType,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
