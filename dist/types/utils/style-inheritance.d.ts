import type { PptRect, PptStyle } from '../types';
import type { Placeholder } from '../core/types';
import type { BaseElement } from '../elements/BaseElement';
export declare function mergePlaceholderStyles(element: BaseElement | null, layoutPlaceholder?: Placeholder, masterPlaceholder?: Placeholder): {
    rect: PptRect;
    style: PptStyle;
    alignmentClass?: string;
};
export declare function mergeBackgroundStyles(slideBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
}, layoutBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
}, masterBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
}): {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
};
export declare function mergeTextStyles(elementStyle?: Partial<PptStyle>, layoutStyle?: Partial<PptStyle>, masterStyle?: Partial<PptStyle>): PptStyle;
export declare function findPlaceholder(placeholders: Placeholder[] | undefined, type: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other', idx?: number): Placeholder | undefined;
