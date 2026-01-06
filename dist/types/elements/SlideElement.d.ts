import type { SlideParseResult } from '../core/types';
import { BaseElement } from './BaseElement';
import type { SlideLayoutResult, MasterSlideResult } from '../core/types';
export declare class SlideElement {
    id: string;
    title: string;
    background: string;
    elements: BaseElement[];
    rawResult: SlideParseResult;
    layout?: SlideLayoutResult;
    master?: MasterSlideResult;
    mediaMap?: Map<string, string>;
    constructor(result: SlideParseResult, elements: BaseElement[], layout?: SlideLayoutResult, master?: MasterSlideResult, mediaMap?: Map<string, string>);
    toHTML(): string;
    private renderLayoutElements;
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
