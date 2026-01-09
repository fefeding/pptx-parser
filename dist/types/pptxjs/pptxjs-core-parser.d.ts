import JSZip from 'jszip';
export declare const PPTXJS_CONSTANTS: {
    readonly slideFactor: number;
    readonly fontSizeFactor: number;
    readonly rtlLangs: readonly ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR", "ur-PK"];
    readonly standardHeight: 6858000;
    readonly standardWidth: 9144000;
};
export interface WarpObj {
    zip: JSZip;
    slideLayoutContent: any;
    slideLayoutTables: any;
    slideMasterContent: any;
    slideMasterTables: any;
    slideContent: any;
    slideResObj: any;
    slideMasterTextStyles: any;
    layoutResObj: any;
    masterResObj: any;
    themeContent: any;
    themeResObj: any;
    digramFileContent?: any;
    diagramResObj?: any;
    defaultTextStyle: any;
}
export interface Relationship {
    id: string;
    type: string;
    target: string;
}
export interface ContentTypes {
    slides: string[];
    slideLayouts: string[];
}
export interface SlideSize {
    width: number;
    height: number;
}
export interface IndexTable {
    idTable: Record<string, any>;
    idxTable: Record<string, any>;
    typeTable: Record<string, any>;
}
export declare class PptxjsCoreParser {
    private zip;
    private slideWidth;
    private slideHeight;
    private slideFactor;
    private fontSizeFactor;
    private defaultTextStyle;
    private appVersion;
    private processFullTheme;
    private incSlide;
    constructor(zip: JSZip, options?: {
        processFullTheme?: boolean;
        incSlideWidth?: number;
        incSlideHeight?: number;
    });
    readXmlFile(filename: string, isSlideContent?: boolean): any;
    private parseXml;
    private domToJson;
    getContentTypes(): ContentTypes;
    getSlideSizeAndSetDefaultTextStyle(): SlideSize;
    indexNodes(content: any): IndexTable;
    getTextByPathList(obj: any, pathList: string[]): any;
    getSlideWidth(): number;
    getSlideHeight(): number;
    getSlideFactor(): number;
    getFontSizeFactor(): number;
    getDefaultTextStyle(): any;
    getAppVersion(): number;
    getProcessFullTheme(): boolean;
}
export declare function angleToDegrees(angle: number | undefined): number;
export declare function degreesToRadians(degrees: number): number;
export * from './pptxjs-utils';
