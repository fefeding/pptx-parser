import { BaseElement } from './BaseElement';
import { PlaceholderElement } from './LayoutElement';
import type { MasterSlideResult } from '../core/types';
export declare class MasterElement extends BaseElement {
    type: 'master';
    masterId?: string;
    elements: BaseElement[];
    placeholders: PlaceholderElement[];
    textStyles?: any;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    colorMap: Record<string, string>;
    constructor(id: string, elements?: BaseElement[], placeholders?: PlaceholderElement[], props?: any);
    static fromResult(result: MasterSlideResult): MasterElement;
    toHTML(): string;
    private getBackgroundStyle;
    getPlaceholderStyle(placeholderType: 'title' | 'body' | 'other'): any;
    private parseParagraphProperties;
}
