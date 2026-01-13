import JSZip from 'jszip';
import { WarpObj, SlideSize, IndexTable } from './pptxjs-core-parser';
export interface PptxjsParserOptions {
    processFullTheme?: boolean;
    incSlideWidth?: number;
    incSlideHeight?: number;
    slideMode?: boolean;
    slideType?: 'div' | 'section' | 'divs2slidesjs' | 'revealjs';
    slidesScale?: string;
}
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
        tables: IndexTable;
        colorMapOvr?: any;
    };
    master?: {
        fileName: string;
        content: any;
        tables: IndexTable;
        colorMapOvr?: any;
    };
    theme?: {
        fileName: string;
        content: any;
    };
    warpObj: WarpObj;
}
export declare class PptxjsParser {
    private zip;
    private coreParser;
    private options;
    private tableStyles;
    constructor(zip: JSZip, options?: PptxjsParserOptions);
    parse(): Promise<{
        slides: SlideData[];
        size: SlideSize;
        thumb?: string;
        globalCSS: string;
    }>;
    private processSingleSlide;
    private processNodesInSlide;
    private processSpNode;
    private processCxnSpNode;
    private processPicNode;
    private processGraphicFrameNode;
    private processTableNode;
    private processChartNode;
    private processGroupSpNode;
    private getBackground;
    private getSlideBackgroundFill;
    private generateGlobalCSS;
}
