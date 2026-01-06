/**
 * 表格元素类
 * 支持完整表格功能
 * 对齐 PPTXjs 的表格解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getChildrenByTagNS, getAttrSafe, parseTextContent, emu2px } from '../utils';
import { NS } from '../constants';
import type { RelsMap } from '../types';

/**
 * 表格单元格
 */
export interface TableCell {
  text?: string;
  rowSpan?: number;
  colSpan?: number;
  style?: {
    backgroundColor?: string;
    borderColor?: string;
    borderWidth?: string;
    padding?: string;
    verticalAlign?: string;
  };
}

/**
 * 表格行
 */
export interface TableRow {
  cells: TableCell[];
  height?: number;
  style?: {
    backgroundColor?: string;
  borderColor?: string;
  borderWidth?: string;
  isHeader?: boolean;
    isFooter?: boolean;
  };
}

/**
 * 表格样式
 */
export interface TableStyle {
  /** 首行样式 */
  firstRow?: {
    backgroundColor?: string;
    fontWeight?: string;
    fontSize?: number;
  };
  /** 末行样式 */
  lastRow?: {
    backgroundColor?: string;
  };
  /** 首列样式 */
  firstCol?: {
    backgroundColor?: string;
    fontWeight?: string;
  };
  /** 末列样式 */
  lastCol?: {
    backgroundColor?: string;
  };
  /** 交替行样式（斑马纹） */
  bandRow?: {
    odd?: {
      backgroundColor?: string;
    };
    even?: {
      backgroundColor?: string;
    };
  };
  /** 交替列样式 */
  bandCol?: {
    odd?: {
      backgroundColor?: string;
    };
    even?: {
      backgroundColor?: string;
    };
  };
}

/**
 * 表格元素类
 */
export class TableElement extends BaseElement {
  type = 'table' as const;

  /** 表格行 */
  rows: TableRow[] = [];

  /** 表格列宽 */
  colWidths: number[] = [];

  /** 表格样式 */
  tableStyle?: TableStyle;

  /** 是否从右到左（RTL） */
  rtl?: boolean;

  /**
   * 从XML节点创建表格元素
   */
  static fromNode(node: Element, relsMap: RelsMap): TableElement | null {
    try {
      const element = new TableElement('', 'table', { x: 0, y: 0, width: 0, height: 0 }, {}, {}, relsMap);

      // 解析ID和名称
      const nvGraphicFramePr = getFirstChildByTagNS(node, 'nvGraphicFramePr', NS.p);
      const cNvPr = nvGraphicFramePr ? getFirstChildByTagNS(nvGraphicFramePr, 'cNvPr', NS.p) : null;

      element.id = getAttrSafe(cNvPr, 'id', element.generateId());
      element.name = getAttrSafe(cNvPr, 'name', '');
      element.hidden = false;

      // 解析位置尺寸
      const xfrm = getFirstChildByTagNS(node, 'xfrm', NS.p);
      if (xfrm) {
          const off = getFirstChildByTagNS(xfrm, 'off', NS.a);
          const ext = getFirstChildByTagNS(xfrm, 'ext', NS.a);

          if (off) {
            element.rect.x = emu2px(off.getAttribute('x') || '0');
            element.rect.y = emu2px(off.getAttribute('y') || '0');
          }
          if (ext) {
            element.rect.width = emu2px(ext.getAttribute('cx') || '0');
            element.rect.height = emu2px(ext.getAttribute('cy') || '0');
          }
      }

      // 检查是否是表格
      const graphic = getFirstChildByTagNS(node, 'graphic', NS.a);
      const graphicData = graphic ? getFirstChildByTagNS(graphic, 'graphicData', NS.a) : null;

      if (!graphicData) {
        return null;
      }

      const uri = graphicData.getAttribute('uri') || '';
      if (!uri.includes('table')) {
        return null;
      }

      const tbl = getFirstChildByTagNS(graphicData, 'tbl', NS.a);
      if (!tbl) {
        return null;
      }

      // 解析表格
      element.parseTable(tbl);

      // 解析表格样式
      element.tableStyle = element.parseTableStyle(tbl);

      // 检查RTL
      element.rtl = tbl.getAttribute('rtl') === '1';

      element.content = {
        rows: element.rows,
        colWidths: element.colWidths,
        tableStyle: element.tableStyle
      };

      element.props = {
        rowCount: element.rows.length,
        colCount: element.colWidths.length
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse table element:', error);
      return null;
    }
  }

  /**
   * 解析表格
   */
  private parseTable(tbl: Element): void {
    // 解析行
    const trNodes = getChildrenByTagNS(tbl, 'tr', NS.a);

    this.rows = trNodes.map(tr => this.parseTableRow(tr));

    // 解析列宽
    const gridCols = getChildrenByTagNS(tbl, 'gridCol', NS.a);
    this.colWidths = gridCols.map(col => emu2px(col.getAttribute('w') || '0'));
  }

  /**
   * 解析表格行
   */
  private parseTableRow(tr: Element): TableRow {
    const row: TableRow = {
      cells: [],
      style: {}
    };

    // 行高
    const h = tr.getAttribute('h');
    if (h) {
      row.height = emu2px(h);
    }

    // 解析单元格
    const tcNodes = getChildrenByTagNS(tr, 'tc', NS.a);
    row.cells = tcNodes.map(tc => this.parseTableCell(tc));

    return row;
  }

  /**
   * 解析表格单元格
   */
  private parseTableCell(tc: Element): TableCell {
    const cell: TableCell = {
      style: {}
    };

    // 合并
    const rowSpan = tc.getAttribute('rowSpan');
    if (rowSpan) {
      cell.rowSpan = parseInt(rowSpan);
    }

    const gridSpan = tc.getAttribute('gridSpan');
    if (gridSpan) {
      cell.colSpan = parseInt(gridSpan);
    }

    // 解析单元格内容
    const txBody = getFirstChildByTagNS(tc, 'txBody', NS.a);
    if (txBody) {
      cell.text = parseTextContent(txBody);
    }

    // 解析单元格样式
    const tcPr = getFirstChildByTagNS(tc, 'tcPr', NS.a);
    if (tcPr) {
      this.parseCellStyle(tcPr, cell);
    }

    return cell;
  }

  /**
   * 解析单元格样式
   */
  private parseCellStyle(tcPr: Element, cell: TableCell): void {
    // 边框
    const tcBdr = getFirstChildByTagNS(tcPr, 'tcBdr', NS.a);
    if (tcBdr) {
      const left = this.parseBorder(tcBdr, 'left');
      const top = this.parseBorder(tcBdr, 'top');
      const right = this.parseBorder(tcBdr, 'right');
      const bottom = this.parseBorder(tcBdr, 'bottom');

      if (cell.style) {
        cell.style.borderColor = left?.color || top?.color;
        cell.style.borderWidth = top?.width || '1px';
      }
    }

    // 背景填充
    const fill = getFirstChildByTagNS(tcPr, 'fill', NS.a);
    if (fill) {
      const solidFill = getFirstChildByTagNS(fill, 'solidFill', NS.a);
      if (solidFill) {
        const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
        if (srgbClr?.getAttribute('val')) {
          cell.style!.backgroundColor = `#${srgbClr.getAttribute('val')}`;
        }
      }
    }

    // 垂直对齐
    const vAlign = tcPr.getAttribute('vAlign');
    if (vAlign && cell.style) {
      cell.style.verticalAlign = vAlign;
    }

    // 边距
    const marL = tcPr.getAttribute('marL');
    const marR = tcPr.getAttribute('marR');
    const marT = tcPr.getAttribute('marT');
    const marB = tcPr.getAttribute('marB');

    if (cell.style) {
      const marLNum = parseInt(marL || '0');
      const marRNum = parseInt(marR || '0');
      const marTNum = parseInt(marT || '0');
      const marBNum = parseInt(marB || '0');

      const padding = [];
      if (marTNum) padding.push(`${emu2px(marTNum)}px`);
      if (marRNum) padding.push(`${emu2px(marRNum)}px`);
      if (marBNum) padding.push(`${emu2px(marBNum)}px`);
      if (marLNum) padding.push(`${emu2px(marLNum)}px`);

      if (padding.length > 0) {
        cell.style.padding = padding.join(' ');
      }
    }
  }

  /**
   * 解析边框
   */
  private parseBorder(tcBdr: Element, side: string): { color?: string; width?: string } | undefined {
    const border = getFirstChildByTagNS(tcBdr, side, NS.a);
    if (!border) return undefined;

    const color = border.getAttribute('color');
    const sz = parseInt(border.getAttribute('sz') || '0');

    return {
      color: color ? `#${color}` : undefined,
      width: `${sz / 100}px`
    };
  }

  /**
   * 解析表格样式
   */
  private parseTableStyle(tbl: Element): TableStyle {
    const style: TableStyle = {};

    // 表格边框
    const wholeTbl = getFirstChildByTagNS(tbl, 'wholeTbl', NS.a);
    if (wholeTbl) {
      // TODO: 解析表格整体样式
    }

    // 表格背景
    const tblBg = getFirstChildByTagNS(tbl, 'tblBg', NS.a);
    if (tblBg) {
      const fill = getFirstChildByTagNS(tblBg, 'fill', NS.a);
      const solidFill = getFirstChildByTagNS(fill, 'solidFill', NS.a);
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        style.firstRow = {
          backgroundColor: `#${srgbClr.getAttribute('val')}`
        };
      }
    }

    // 首行样式
    const firstRow = getFirstChildByTagNS(tbl, 'firstRow', NS.a);
    if (firstRow) {
      style.firstRow = this.parseRowStyle(firstRow);
    }

    // 末行样式
    const lastRow = getFirstChildByTagNS(tbl, 'lastRow', NS.a);
    if (lastRow) {
      style.lastRow = {
        backgroundColor: this.parseRowBackgroundColor(lastRow)
      };
    }

    // 首列样式
    const firstCol = getFirstChildByTagNS(tbl, 'firstCol', NS.a);
    if (firstCol) {
      style.firstCol = {
        backgroundColor: this.parseRowBackgroundColor(firstCol)
      };
    }

    // 末列样式
    const lastCol = getFirstChildByTagNS(tbl, 'lastCol', NS.a);
    if (lastCol) {
      style.lastCol = {
        backgroundColor: this.parseRowBackgroundColor(lastCol)
      };
    }

    // 交替行样式
    const band1H = getFirstChildByTagNS(tbl, 'band1H', NS.a);
    if (band1H) {
      style.bandRow = {
        odd: {
          backgroundColor: this.parseRowBackgroundColor(band1H)
        }
      };
    }

    const band2H = getFirstChildByTagNS(tbl, 'band2H', NS.a);
    if (band2H) {
      if (!style.bandRow) style.bandRow = {};
      style.bandRow!.even = {
        backgroundColor: this.parseRowBackgroundColor(band2H)
      };
    }

    return style;
  }

  /**
   * 解析行样式
   */
  private parseRowStyle(rowStyle: Element): TableStyle['firstRow'] {
    const fill = getFirstChildByTagNS(rowStyle, 'fill', NS.a);
    if (fill) {
      const solidFill = getFirstChildByTagNS(fill, 'solidFill', NS.a);
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        const bgColor = `#${srgbClr.getAttribute('val')}`;
        return {
          backgroundColor: bgColor
        };
      }
    }

    return {};
  }

  /**
   * 解析行背景色
   */
  private parseRowBackgroundColor(rowStyle: Element): string | undefined {
    const fill = getFirstChildByTagNS(rowStyle, 'fill', NS.a);
    if (fill) {
      const solidFill = getFirstChildByTagNS(fill, 'solidFill', NS.a);
      const srgbClr = getFirstChildByTagNS(solidFill, 'srgbClr', NS.a);
      if (srgbClr?.getAttribute('val')) {
        return `#${srgbClr.getAttribute('val')}`;
      }
    }
    return undefined;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();
    const tableStyle = this.getTableStyle();

    const rowsHTML = this.rows.map(row => this.rowToHTML(row)).join('\n');

    return `<div style="${style}">
      <table style="${tableStyle}">
        <tbody>
${rowsHTML}
        </tbody>
      </table>
    </div>`;
  }

  /**
   * 获取表格样式
   */
  private getTableStyle(): string {
    const styles = [
      `width: 100%`,
      `height: 100%`,
      `border-collapse: collapse`,
      `border-spacing: 0`
    ];

    return styles.join('; ');
  }

  /**
   * 行转HTML
   */
  private rowToHTML(row: TableRow): string {
    const rowStyle = this.getRowStyle(row);
    const cellsHTML = row.cells.map(cell => this.cellToHTML(cell)).join('\n        ');

    return `      <tr style="${rowStyle}">\n${cellsHTML}\n      </tr>`;
  }

  /**
   * 获取行样式
   */
  private getRowStyle(row: TableRow): string {
    const styles = [];

    if (row.height) {
      styles.push(`height: ${row.height}px`);
    }

    // 应用行样式
    if (row.style?.isHeader && this.tableStyle?.firstRow?.backgroundColor) {
      styles.push(`background-color: ${this.tableStyle.firstRow.backgroundColor}`);
    } else if (row.style?.isFooter && this.tableStyle?.lastRow?.backgroundColor) {
      styles.push(`background-color: ${this.tableStyle.lastRow.backgroundColor}`);
    }

    return styles.join('; ');
  }

  /**
   * 单元格转HTML
   */
  private cellToHTML(cell: TableCell): string {
    const cellStyle = this.getCellStyle(cell);

    const attrs = [];
    if (cell.rowSpan && cell.rowSpan > 1) attrs.push(`rowspan="${cell.rowSpan}"`);
    if (cell.colSpan && cell.colSpan > 1) attrs.push(`colspan="${cell.colSpan}"`);

    const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';

    return `<td${attrStr} style="${cellStyle}">${cell.text || ''}</td>`;
  }

  /**
   * 获取单元格样式
   */
  private getCellStyle(cell: TableCell): string {
    const styles = [];

    if (cell.style?.backgroundColor) {
      styles.push(`background-color: ${cell.style.backgroundColor}`);
    }

    if (cell.style?.borderColor && cell.style?.borderWidth) {
      styles.push(`border: ${cell.style.borderWidth} solid ${cell.style.borderColor}`);
    }

    if (cell.style?.padding) {
      styles.push(`padding: ${cell.style.padding}`);
    }

    if (cell.style?.verticalAlign) {
      styles.push(`vertical-align: ${cell.style.verticalAlign}`);
    }

    return styles.join('; ');
  }

  /**
   * 转换为ParsedSlideElement格式
   */
  toParsedElement(): any {
    return {
      id: this.id,
      type: 'table',
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
