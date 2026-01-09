import type { PptDocument } from '../types';
export interface HtmlGenerationOptions {
    slideType?: 'div' | 'section';
    includeGlobalCSS?: boolean;
    containerClass?: string;
}
export declare class HtmlGenerator {
    private options;
    private styleTable;
    private styleCounter;
    constructor(options?: HtmlGenerationOptions);
    generate(document: PptDocument): string;
    private generateSlide;
    private generateElement;
    private generateTextElement;
    private generateTextStyle;
    private generateImageElement;
    private generateBackground;
    private generateShapeElement;
    private generateTableElement;
    private generateGlobalCSS;
    private getStyleClass;
    private escapeHtml;
}
export declare function generateHtml(document: PptDocument, options?: HtmlGenerationOptions): string;
