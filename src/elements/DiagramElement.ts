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
 * 解析关系ID字符串
 */
interface RelIds {
  colors?: string;
  data?: string;
  layout?: string;
  quickStyle?: string;
}

/**
 * 颜色数据
 */
interface ColorData {
  colorScheme?: Record<string, string>;
}

/**
 * 布局数据
 */
interface LayoutData {
  layoutDef?: string;
}

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

    // 解析关系ID
    const relIdsAttr = dgm.getAttributeNS(NS.d, 'relIds') || '';
    const relIds = this.parseRelIds(relIdsAttr);

    // 根据关系ID获取相关数据
    if (relIds.colors && this.relsMap[relIds.colors]) {
      data.colors = this.fetchColorData(this.relsMap[relIds.colors].target);
    }

    if (relIds.data && this.relsMap[relIds.data]) {
      data.data = this.fetchLayoutData(this.relsMap[relIds.data].target);
    }

    if (relIds.layout && this.relsMap[relIds.layout]) {
      data.layout = this.relsMap[relIds.layout].target;
    }

    // 尝试从dgm节点中提取形状信息
    const spTree = getFirstChildByTagNS(dgm, 'spTree', NS.d);
    if (spTree) {
      data.shapes = this.parseShapes(spTree);
    }

    return data;
  }

  /**
   * 解析关系ID字符串
   */
  private parseRelIds(relIdsAttr: string): RelIds {
    const relIds: RelIds = {};
    
    if (!relIdsAttr) return relIds;
    
    // relIds格式通常为: "rId1,rId2,rId3,rId4"
    // 对应: colors, data, layout, quickStyle
    const ids = relIdsAttr.split(',').map(id => id.trim());
    
    if (ids.length >= 1) relIds.colors = ids[0];
    if (ids.length >= 2) relIds.data = ids[1];
    if (ids.length >= 3) relIds.layout = ids[2];
    if (ids.length >= 4) relIds.quickStyle = ids[3];
    
    return relIds;
  }

  /**
   * 获取颜色数据
   */
  private async fetchColorData(target: string): Promise<Record<string, string> | undefined> {
    try {
      // 在实际实现中，这里需要从ZIP文件中读取对应的XML文件
      // 由于当前是同步方法，我们先返回空对象
      // 在实际使用中，应该通过parser传递ZIP引用或使用异步方法
      
      // 示例：colors1.xml 的路径通常是 ppt/diagrams/colors1.xml
      const colorPath = target.startsWith('../') ? target.substring(3) : target;
      
      // 这里需要访问ZIP文件，但由于架构限制，暂时返回空
      // 在实际项目中，可以通过构造函数传入ZIP引用或使用其他方式
      
      return {};
    } catch (error) {
      console.warn('Failed to fetch color data:', error);
      return undefined;
    }
  }

  /**
   * 获取布局数据
   */
  private async fetchLayoutData(target: string): Promise<Record<string, any> | undefined> {
    try {
      // 类似地，这里需要读取data#.xml文件
      const layoutPath = target.startsWith('../') ? target.substring(3) : target;
      
      // 在实际实现中，需要解析XML并返回数据结构
      // 这里简化处理，返回空对象
      
      return {};
    } catch (error) {
      console.warn('Failed to fetch layout data:', error);
      return undefined;
    }
  }

  /**
   * 同步版本的颜色数据获取（用于当前架构）
   */
  private fetchColorDataSync(target: string): Record<string, string> | undefined {
    // 在当前同步架构下，我们无法直接读取ZIP文件
    // 所以这里提供一个钩子，让调用者可以预先加载这些数据
    // 或者在实际使用时通过其他方式注入
    
    // 返回一个标记，表示需要异步加载
    return { _needsAsyncLoad: true, target };
  }

  /**
   * 同步版本的布局数据获取
   */
  private fetchLayoutDataSync(target: string): Record<string, any> | undefined {
    // 同上，同步版本无法直接读取文件
    return { _needsAsyncLoad: true, target };
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
    const dataAttrs = this.formatDataAttributes();
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

    return `<div ${dataAttrs} style="${style}">
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
