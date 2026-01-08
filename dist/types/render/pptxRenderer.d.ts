import { PptxToHtmlOptions } from "../types/index";
export declare class PptxRenderer {
    private options;
    private buffer;
    private env;
    private domAdapter;
    private parser;
    constructor(options: PptxToHtmlOptions, buffer: Buffer);
    render(targetId: string): Promise<void | string>;
    private renderSlide;
    private renderDivs2SlidesJsSlide;
    private renderRevealJsSlide;
    private initInteractions;
}
