import JSZip from 'jszip';
import type { PptxParseResult } from './types';
export declare function parseImages(zip: JSZip, result: PptxParseResult): Promise<void>;
