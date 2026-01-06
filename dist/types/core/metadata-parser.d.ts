import JSZip from 'jszip';
import type { Metadata, SlideSize } from './types';
export declare function parseCoreProperties(zip: JSZip): Promise<Metadata>;
export declare function parseSlideLayoutSize(zip: JSZip): Promise<SlideSize>;
export declare function inferPageSize(ratio: number): '4:3' | '16:9' | '16:10' | 'custom';
