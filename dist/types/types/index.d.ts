export interface PptxToHtmlOptions {
    pptxFileUrl: string | Buffer;
    fileInputId?: string;
    slidesScale?: string;
    slideMode: boolean;
    keyBoardShortCut: boolean;
    mediaProcess: boolean;
    jsZipV2: string | false;
    themeProcess: boolean | "colorsAndImageOnly";
    incSlide?: {
        height: number;
        width: number;
    };
    slideType: "divs2slidesjs" | "revealjs";
    slideModeConfig: SlideModeConfig;
    revealjsConfig?: RevealJsConfig;
}
export interface SlideModeConfig {
    first: number;
    nav: boolean;
    navTxtColor: string;
    navNextTxt?: string;
    navPrevTxt?: string;
    showPlayPauseBtn: boolean;
    showSlideNum: boolean;
    showTotalSlideNum: boolean;
    autoSlide: false | number;
    randomAutoSlide: boolean;
    loop: boolean;
    background: false | string;
    transition: "slid" | "fade" | "default" | "random";
    transitionTime: number;
}
export interface RevealJsConfig {
    transition?: string;
    backgroundTransition?: string;
    autoSlide?: number;
    loop?: boolean;
    slideNumber?: boolean;
}
export declare enum PptxElementType {
    TEXT = "text",
    TEXT_BLOCK = "text_block",
    SHAPE = "shape",
    MEDIA = "media",
    GRAPH = "graph",
    TABLE = "table",
    SMART_ART = "smart_art",
    EQUATION = "equation"
}
export interface MediaInfo {
    type: "image" | "audio" | "video";
    src: string;
    format: string;
    browserSupport: Record<string, string[]>;
}
export type EnvType = "browser" | "node";
