import JSZip from 'jszip';
import type { RelsMap } from './types';
export declare function parseGlobalRels(zip: JSZip): Promise<RelsMap>;
export declare function parseSlideRels(zip: JSZip, slideId: string, relsBasePath?: string): Promise<RelsMap>;
export declare function getSlideLayoutRef(relsMap: RelsMap): string | undefined;
