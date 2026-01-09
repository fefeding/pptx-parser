import JSZip from 'jszip';
import { generateId, emu2px, px2emu } from '../utils/index';
import type { PptxParseResult, ParseOptions } from './types';
import type { PptDocument } from '../types';
export declare const PptParseUtils: {
    generateId: typeof generateId;
    parseXmlText: (text: string) => string;
    px2emu: typeof px2emu;
    emu2px: typeof emu2px;
    parseXmlToTree: (xmlStr: string) => any;
    parseXmlAttrs: (attrs: NamedNodeMap) => Record<string, string>;
    parseXmlRect: (attrs: Record<string, string>) => {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    parseXmlStyle: (attrs: Record<string, string>) => any;
    hexToRgb: (hex: string) => {
        r: number;
        g: number;
        b: number;
    };
    rgbToHex: (r: number, g: number, b: number) => string;
    parseColor: (color: string) => string;
};
export declare function parsePptx(file: File | Blob | ArrayBuffer, options?: ParseOptions & {
    returnFormat?: 'enhanced' | 'simple';
}): Promise<PptxParseResult | PptDocument>;
export declare function serializePptx(pptDoc: PptDocument): Promise<Blob>;
export declare function writeSlides(zip: JSZip, slides: any[]): Promise<void>;
