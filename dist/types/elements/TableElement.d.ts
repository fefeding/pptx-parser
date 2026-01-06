import { BaseElement } from './BaseElement';
import type { RelsMap } from '../types';
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
export interface TableStyle {
    firstRow?: {
        backgroundColor?: string;
        fontWeight?: string;
        fontSize?: number;
    };
    lastRow?: {
        backgroundColor?: string;
    };
    firstCol?: {
        backgroundColor?: string;
        fontWeight?: string;
    };
    lastCol?: {
        backgroundColor?: string;
    };
    bandRow?: {
        odd?: {
            backgroundColor?: string;
        };
        even?: {
            backgroundColor?: string;
        };
    };
    bandCol?: {
        odd?: {
            backgroundColor?: string;
        };
        even?: {
            backgroundColor?: string;
        };
    };
}
export declare class TableElement extends BaseElement {
    type: "table";
    rows: TableRow[];
    colWidths: number[];
    tableStyle?: TableStyle;
    rtl?: boolean;
    static fromNode(node: Element, relsMap: RelsMap): TableElement | null;
    private parseTable;
    private parseTableRow;
    private parseTableCell;
    private parseCellStyle;
    private parseBorder;
    private parseTableStyle;
    private parseRowStyle;
    private parseRowBackgroundColor;
    toHTML(): string;
    private getTableStyle;
    private rowToHTML;
    private getRowStyle;
    private cellToHTML;
    private getCellStyle;
    toParsedElement(): any;
}
