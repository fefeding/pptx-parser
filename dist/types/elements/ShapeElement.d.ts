import { BaseElement } from './BaseElement';
import type { ParsedShapeElement, RelsMap } from '../types';
export interface TextRun {
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    strike?: boolean;
    color?: string;
    backgroundColor?: string;
    highlight?: string;
}
export interface BulletStyle {
    type?: 'none' | 'char' | 'blip' | 'autoNum';
    char?: string;
    imageSrc?: string;
    autoNumType?: string;
    level?: number;
    color?: string;
    size?: number;
    font?: string;
}
export declare class ShapeElement extends BaseElement {
    type: 'shape' | 'text';
    shapeType?: string;
    text?: string;
    textStyle?: TextRun[];
    paragraphStyle?: {
        align?: 'left' | 'center' | 'right' | 'justify';
        indent?: number;
        lineSpacing?: number;
        spaceBefore?: number;
        spaceAfter?: number;
        rtl?: boolean;
    };
    bulletStyle?: BulletStyle;
    isPlaceholder?: boolean;
    placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    hyperlink?: {
        id?: string;
        url?: string;
        tooltip?: string;
    };
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    static fromNode(node: Element, relsMap: RelsMap): ShapeElement | null;
    private parseShapeProperties;
    private parseFill;
    private parseTextBody;
    private parseParagraph;
    private parseBulletStyle;
    private parseTextRun;
    private parseColor;
    private detectShapeType;
    toHTML(): string;
    private getTextStyle;
    private getRotationStyle;
    private getShapeStyle;
    private textStyleFromAlign;
    private textStyleFromFontSize;
    private escapeHtml;
    private parseStyleString;
    toParsedElement(): ParsedShapeElement;
}
