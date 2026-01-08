import { Buffer } from "buffer";
export declare class PptxParser {
    private buffer;
    private zip;
    private slides;
    private totalSlides;
    constructor(buffer: Buffer);
    init(): Promise<void>;
    private parseSlides;
    private parseSlideContent;
    private parseTextElements;
    private parseMediaElements;
    getSlides(): Record<number, any>;
    getTotalSlides(): number;
    private parseTheme;
    private parseMedia;
    private parseTextBlocks;
    private parseShapes;
    private parseGraphs;
    private parseTables;
    private parseSmartArt;
    private parseEquations;
}
