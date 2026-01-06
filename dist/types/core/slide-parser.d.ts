import JSZip from 'jszip';
import { type BaseElement } from '../elements';
import type { ParseOptions, RelsMap, SlideParseResult } from './types';
export declare function parseSlide(slideXml: string, relsMap?: RelsMap, slideIndex?: number): SlideParseResult;
export declare function parseSlideElements(root: Element, relsMap: RelsMap): BaseElement[];
export declare function parseAllSlides(zip: JSZip, options: ParseOptions): Promise<SlideParseResult[]>;
export declare function parseSlideBackground(root: Element, relsMap?: RelsMap): {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
};
