export * from './pptxjs-core-parser';
export * from './pptxjs-utils';
export * from './pptxjs-color-utils';
export { getColorValue, applyColorMap, parseColorMapOverride, generateCssColor, parseColorFill, } from './pptxjs-color-utils';
export { parseTextBoxContent, generateTextBoxHtml, mergeTextStyles, getDefaultTextStyle, } from './pptxjs-text-utils';
export { PptxjsParser } from './pptxjs-parser';
export interface SlideData {
    id: number;
    fileName: string;
    width: number;
    height: number;
    bgColor?: string;
    bgFill?: any;
    shapes: any[];
    images: any[];
    tables: any[];
    charts: any[];
    layout?: {
        fileName: string;
        content: any;
        tables: any;
        colorMapOvr?: any;
    };
    master?: {
        fileName: string;
        content: any;
        tables: any;
        colorMapOvr?: any;
    };
    theme?: {
        fileName: string;
        content: any;
    };
    warpObj: any;
}
export interface PptxjsParserOptions {
    processFullTheme?: boolean;
    incSlideWidth?: number;
    incSlideHeight?: number;
    slideMode?: boolean;
    slideType?: 'div' | 'section' | 'divs2slidesjs' | 'revealjs';
    slidesScale?: string;
}
export declare function parsePptx(file: ArrayBuffer | Blob | Uint8Array, options?: PptxjsParserOptions): Promise<{
    slides: import("./pptxjs-parser").SlideData[];
    size: import("./pptxjs-core-parser").SlideSize;
    thumb?: string;
    globalCSS: string;
}>;
export declare class Pptxjs {
    private parser;
    private parsedData;
    constructor(file: ArrayBuffer | Blob | Uint8Array, options?: PptxjsParserOptions);
    static create(file: ArrayBuffer | Blob | Uint8Array, options?: PptxjsParserOptions): Promise<Pptxjs>;
    parse(): Promise<void>;
    getResult(): {
        slides: import("./pptxjs-parser").SlideData[];
        size: import("./pptxjs-core-parser").SlideSize;
        thumb?: string;
        globalCSS: string;
    } | null;
    getSlides(): import("./pptxjs-parser").SlideData[];
    getSize(): import("./pptxjs-core-parser").SlideSize;
    getThumb(): string | undefined;
    getGlobalCSS(): string;
    generateHtml(): string;
    private generateSlideHtml;
    private generateShapeHtml;
    private generateImageHtml;
    private generateTableHtml;
    private generateChartHtml;
}
export default Pptxjs;
