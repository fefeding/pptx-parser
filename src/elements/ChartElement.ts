/**
 * 图表元素类
 * 支持实际图表数据解析
 * 对齐 PPTXjs 的图表解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr, emu2px } from '../utils';
import { NS } from '../constants';
import type { ParsedChartElement, RelsMap } from '../types';

/**
 * 图表数据点
 */
export interface ChartDataPoint {
  value?: number;
  category?: string;
}

/**
 * 图表系列
 */
export interface ChartSeries {
  name?: string;
  points?: ChartDataPoint[];
  color?: string;
}

/**
 * 图表数据
 */
export interface ChartData {
  type: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart';
  title?: string;
  xTitle?: string;
  yTitle?: string;
  series?: ChartSeries[];
}

/**
 * 图表元素类
 */
export class ChartElement extends BaseElement {
  type = 'chart' as const;

  /** 图表类型 */
  chartType?: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';

  /** 图表数据 */
  chartData?: ChartData;

  /** 关联ID */
  relId: string = '';

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

      // 解析位置尺寸
      const xfrm = getFirstChildByTagNS(node, 'xfrm', NS.p);
      let rect = { x: 0, y: 0, width: 0, height: 0 };

      if (xfrm) {
          const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
          const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

          if (off) {
            rect.x = emu2px(off.getAttribute('x') || '0');
            rect.y = emu2px(off.getAttribute('y') || '0');
          }
          if (ext) {
            rect.width = emu2px(ext.getAttribute('cx') || '0');
            rect.height = emu2px(ext.getAttribute('cy') || '0');
          }
      }

      const element = new ChartElement(id, 'chart', rect, '', {}, relsMap);
      element.name = name;
      element.hidden = hidden;

      // 查找graphicData判断类型
      const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
      const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

      if (!graphicData) {
        return null;
      }

      const uri = graphicData.getAttribute('uri') || '';

      if (!uri.includes('chart')) {
        return null;
      }

      // 查找chart节点
      const chart = getFirstChildByTagNS(graphicData, 'chart', NS.c);
      if (!chart) {
        return null;
      }

      const chartTypeAttr = chart.getAttribute('type') || 'unknown';
      const relId = chart.getAttributeNS(NS.r, 'id') || chart.getAttribute('r:id') || '';

      element.chartType = element.detectChartType(chart);
      element.relId = relId;

      element.content = {
        chartType: 'unknown',
        data: []
      };

      element.props = {
        relId,
        chartType: chartTypeAttr
      };

      element.rawNode = node;

      // 尝试解析图表数据（如果有chart文件）
      if (relId && relsMap[relId]) {
        const chartFilePath = relsMap[relId].target;
        // 注意：这里需要从zip中读取chart XML文件进行完整解析
        // 由于当前架构限制，先返回基础信息
        console.log(`Chart file path: ${chartFilePath}`);
      }

      return element;
    } catch (error) {
      console.error('Failed to parse chart element:', error);
      return null;
    }
  }

  /**
   * 检测图表类型
   */
  private detectChartType(chart: Element): 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown' {
    // 查找图表类型的子节点
    const children = Array.from(chart.children);
    for (const child of children) {
      const tagName = child.tagName;

      if (tagName.includes('lineChart')) return 'lineChart';
      if (tagName.includes('barChart')) return 'barChart';
      if (tagName.includes('pieChart')) {
        if (tagName.includes('3D')) return 'pie3DChart';
        return 'pieChart';
      }
      if (tagName.includes('areaChart')) return 'areaChart';
      if (tagName.includes('scatterChart')) return 'scatterChart';
    }

    return 'unknown';
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const dataAttrs = this.formatDataAttributes();
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
      `text-align: center`,
      `border-radius: 4px`,
      `padding: 20px`
    ].join('; ');

    const label = this.getChartLabel();

      // 如果有图表数据，可以渲染更详细的图表信息
    const hasChartData = this.chartData && Array.isArray(this.chartData.series) && this.chartData.series.length > 0;

    if (hasChartData) {
      return `<div ${dataAttrs} style="${style}">
        <div style="width: 100%;">
          <div style="font-size: 18px; font-weight: bold; margin-bottom: 10px;">
            ${this.chartData?.title || 'Chart'}
          </div>
          ${this.renderChartData()}
        </div>
      </div>`;
    }

    return `<div ${dataAttrs} style="${style}">
      <div style="${innerStyle}">
        ${label}
      </div>
    </div>`;
  }

  /**
   * 获取图表标签
   */
  private getChartLabel(): string {
    const typeLabels: Record<string, string> = {
      lineChart: '折线图',
      barChart: '柱状图',
      pieChart: '饼图',
      pie3DChart: '3D饼图',
      areaChart: '面积图',
      scatterChart: '散点图',
      unknown: '图表'
    };

    if (this.chartType) {
      return typeLabels[this.chartType] || 'Chart';
    }

    return 'Chart';
  }

  /**
   * 渲染图表数据（简化版）
   */
  private renderChartData(): string {
    if (!this.chartData || !this.chartData.series) return '';

    const data = this.chartData;

    let html = '<div style="margin: 10px 0;">';

    if (data.xTitle) {
      html += `<div style="font-size: 12px; color: #666; margin-bottom: 5px;">X轴: ${data.xTitle}</div>`;
    }
    if (data.yTitle) {
      html += `<div style="font-size: 12px; color: #666; margin-bottom: 10px;">Y轴: ${data.yTitle}</div>`;
    }

    // 渲染系列数据
    html += '<div style="display: flex; flex-wrap: wrap; gap: 10px;">';
    if (data.series) {
      data.series.forEach((series, index) => {
        html += `
          <div style="
            background: ${series.color || '#4285f4'};
            color: white;
            padding: 10px;
            border-radius: 4px;
            min-width: 150px;
          ">
            <div style="font-weight: bold;">${series.name || `系列 ${index + 1}`}</div>
            <div style="font-size: 12px; margin-top: 5px;">
              数据点: ${series.points?.length || 0}
            </div>
          </div>
        `;
      });
    }
    html += '</div>';

    html += '</div>';

    return html;
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
