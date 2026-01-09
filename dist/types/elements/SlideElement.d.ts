import type { SlideParseResult } from '../core/types';
import type { SlideLayoutResult, MasterSlideResult } from '../core/types';
import type { LayoutElement } from './LayoutElement';
import type { MasterElement } from './MasterElement';
import { BaseElement } from './BaseElement';
export declare class SlideElement {
    id: string;
    title: string;
    background: string;
    elements: BaseElement[];
    rawResult: SlideParseResult;
    layoutElement?: LayoutElement;
    masterElement?: MasterElement;
    layout?: SlideLayoutResult;
    master?: MasterSlideResult;
    mediaMap?: Map<string, string>;
    constructor(result: SlideParseResult, elements: BaseElement[], layoutElement?: LayoutElement, masterElement?: MasterElement, layoutResult?: SlideLayoutResult, masterResult?: MasterSlideResult, mediaMap?: Map<string, string>);
    toHTML(): string;
    private renderMasterElements;
    private renderLayoutElements;
    private renderSlideElementsWithLayout;
    private getTextContentStyle;
    private getSlideBackground;
    toHTMLString(): string;
    private escapeHtml;
}
export declare class PptxDocument {
    id: string;
    title: string;
    author?: string;
    slides: SlideElement[];
    width: number;
    height: number;
    ratio: number;
    constructor(id: string, title: string, slides: SlideElement[], width?: number, height?: number, author?: string);
    static fromParseResult(result: any): PptxDocument;
    toHTML(): string;
    toHTMLWithNavigation(): string;
    private escapeHtml;
}
