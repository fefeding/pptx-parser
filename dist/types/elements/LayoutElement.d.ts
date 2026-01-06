import { BaseElement } from './BaseElement';
import type { SlideLayoutResult } from '../core/types';
import type { TextRun } from './ShapeElement';
export declare class PlaceholderElement extends BaseElement {
    type: 'placeholder';
    placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    idx?: number;
    name?: string;
    text?: string;
    textStyle?: TextRun[];
    constructor(id: string, placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other', rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, props?: any);
    toHTML(): string;
    private getPlaceholderContent;
    private renderTextContent;
    private get isContentSet();
    private getTextRunStyle;
}
export declare class LayoutElement extends BaseElement {
    type: 'layout';
    name?: string;
    placeholders: PlaceholderElement[];
    textStyles?: any;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    constructor(id: string, name?: string, placeholders?: PlaceholderElement[], props?: any);
    static fromResult(result: SlideLayoutResult): LayoutElement;
    toHTML(): string;
    private getBackgroundStyle;
}
