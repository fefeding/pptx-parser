/**
 * Diagram/SmartArt 元素类
 * 支持图解和 SmartArt
 * 对齐 PPTXjs 的 Diagram 解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr, emu2px } from '../utils';
import { NS } from '../constants';
import type { RelsMap, PptRect } from '../types';

/**
 * Diagram 数据模型
 */
export interface DiagramData {
  colors?: Record<string, string>;
  data?: Record<string, any>;
  layout?: string;
  shapes?: DiagramShape[];
}

/**
 * Diagram 形状
 */
export interface DiagramShape {
  id: string;
  type: string;
  position?: { x: number; y: number };
  size?: { width: number; height: number };
  text?: string;
}

/**
 * Diagram/SmartArt 元素类
 */
export class DiagramElement extends BaseElement {
  type = 'diagram' as const;

  /** Diagram 数据 */
  diagramData?: DiagramData;

  /** 关联ID */
  relId: string = '';

  constructor(
    id: string,
    rect: PptRect,
    content: any = {},
    props: any = {},
    relsMap: Record<string, any> = {}
  ) {
    super(id, 'diagram', rect, content, props, relsMap);
  }

  /**
   * 从XML节点创建Diagram元素
   */
  static fromNode(node: Element, relsMap: RelsMap): DiagramElement | null {
    try {
      // 解析ID和名称
      const nvGraphicFramePr = getFirstChildByTagNS(node, 'nvGraphicFramePr', NS.p);
      const cNvPr = nvGraphicFramePr ? getFirstChildByTagNS(nvGraphicFramePr, 'cNvPr', NS.p) : null;

      const id = getAttrSafe(cNvPr, 'id', `diagram_${Date.now()}`);
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

      const element = new DiagramElement(id, rect, '', {}, relsMap);
      element.name = name;
      element.hidden = hidden;

      // 查找graphicData判断类型
      const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
      const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

      if (!graphicData) {
        return null;
      }

      const uri = graphicData.getAttribute('uri') || '';

      if (!uri.includes('diagram')) {
        return null;
      }

      // 查找dgm节点
      const dgm = getFirstChildByTagNS(graphicData, 'dgm', NS.d);
      if (!dgm) {
        return null;
      }

      const relId = dgm.getAttributeNS(NS.r, 'rel') || dgm.getAttribute('r:rel') || '';

      element.relId = relId;
      element.diagramData = element.parseDiagramData(dgm, relsMap);

      element.content = {
        diagramData: element.diagramData
      };

      element.props = {
        relId,
        diagramType: uri
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse diagram element:', error);
      return null;
    }
  }

  /**
   * 解析Diagram数据
   */
  private parseDiagramData(dgm: Element, relsMap: RelsMap): DiagramData {
    const data: DiagramData = {};

    // 需要读取多个XML文件：colors#.xml, data#.xml, layout#.xml, quickStyle#.xml
    // 由于当前架构限制，先解析基础信息

    const relIds = dgm.getAttributeNS(NS.d, 'relIds') || '';
    if (relIds) {
      // 解析关系ID
      const ids = relIds.split(',').map(id => id.trim());
      // TODO: 根据关系ID获取颜色、数据、布局等信息
    }

    // 尝试从dgm节点中提取形状信息
    const spTree = getFirstChildByTagNS(dgm, 'spTree', NS.d);
    if (spTree) {
      data.shapes = this.parseShapes(spTree);
    }

    return data;
  }

  /**
   * 解析形状
   */
  private parseShapes(spTree: Element): DiagramShape[] {
    const shapes: DiagramShape[] = [];
    const shapeElements = Array.from(spTree.children);

    for (const shape of shapeElements) {
      const id = shape.getAttribute('id') || '';
      const type = shape.tagName || 'shape';

      // 解析位置
      const xfrm = getFirstChildByTagNS(shape, 'xfrm', NS.d);
      let position;
      if (xfrm) {
        const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
        const ext = getFirstChildByTagNS(xfrm, 'ext', NS.d);

        if (off) {
          const x = parseInt(off.getAttribute('x') || '0');
          const y = parseInt(off.getAttribute('y') || '0');
          position = { x: emu2px(x), y: emu2px(y) };
        }

        if (ext) {
          const width = parseInt(ext.getAttribute('cx') || '0');
          const height = parseInt(ext.getAttribute('cy') || '0');
          shapes.push({
            id,
            type,
            position,
            size: { width: emu2px(width), height: emu2px(height) }
          });
        }
      } else {
        shapes.push({ id, type });
      }
    }

    return shapes;
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
      `background-color: #f0f0f0`,
      `border: 1px dashed #ccc`,
      `color: #666`,
      `font-size: 14px`,
      `text-align: center`,
      `padding: 20px`,
      `border-radius: 4px`
    ].join('; ');

    const diagramInfo = this.getDiagramInfo();

    return `<div style="${style}">
      <div style="${innerStyle}">
        ${diagramInfo}
      </div>
    </div>`;
  }

  /**
   * 获取Diagram信息
   */
  private getDiagramInfo(): string {
    const shapesCount = this.diagramData?.shapes?.length || 0;

    let html = '<div style="text-align: left;">';
    html += `<div style="font-size: 16px; font-weight: bold; margin-bottom: 10px;">SmartArt / Diagram</div>`;
    html += `<div style="margin-bottom: 10px;">形状数量: ${shapesCount}</div>`;

    if (this.diagramData?.shapes && shapesCount > 0) {
      html += '<div style="margin-top: 10px;">';
      this.diagramData.shapes.slice(0, 10).forEach((shape, index) => {
        html += `<div style="
          display: inline-block;
          width: 60px;
          height: 40px;
          background: #e0e0e0;
          border: 1px solid #ccc;
          border-radius: 4px;
          margin: 5px;
          font-size: 12px;
        ">
          ${shape.type}
        </div>`;
      });
      if (shapesCount > 10) {
        html += `<div style="display: inline-block; padding: 10px;">...还有 ${shapesCount - 10} 个形状</div>`;
      }
      html += '</div>';
    }

    html += '</div>';

    return html;
  }

  /**
   * 转换为ParsedSlideElement格式
   */
  toParsedElement(): any {
    return {
      id: this.id,
      type: 'diagram',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
