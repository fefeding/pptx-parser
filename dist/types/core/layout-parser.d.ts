import JSZip from 'jszip';
import type { RelsMap, SlideLayoutResult } from './types';
export declare function parseSlideLayout(layoutXml: string, relsMap?: RelsMap): SlideLayoutResult | null;
export declare function parseAllSlideLayouts(zip: JSZip): Promise<Record<string, SlideLayoutResult>>;
export declare function mergeBackgrounds(slideBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
}, layoutBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
}, masterBackground?: {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
}): {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
};
