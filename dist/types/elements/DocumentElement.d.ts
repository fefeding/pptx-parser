import type { PptxParseResult } from '../core/types';
import { BaseElement } from './BaseElement';
import { SlideElement } from './SlideElement';
import { LayoutElement } from './LayoutElement';
import { MasterElement } from './MasterElement';
import { TagsElement } from './TagsElement';
import { NotesMasterElement, NotesSlideElement } from './NotesElement';
export interface HtmlRenderOptions {
    includeStyles?: boolean;
    includeScripts?: boolean;
    includeLayoutElements?: boolean;
    withNavigation?: boolean;
    customCss?: string;
}
export declare class DocumentElement extends BaseElement {
    type: 'document';
    title: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    created?: string;
    modified?: string;
    slides: SlideElement[];
    layouts: Record<string, LayoutElement>;
    masters: MasterElement[];
    tags: TagsElement[];
    notesMasters: NotesMasterElement[];
    notesSlides: NotesSlideElement[];
    width: number;
    height: number;
    ratio: number;
    pageSize: '4:3' | '16:9' | '16:10' | 'custom';
    globalRelsMap: Record<string, any>;
    mediaMap?: Map<string, string>;
    constructor(id: string, title: string, width?: number, height?: number, props?: any);
    static fromParseResult(result: PptxParseResult): DocumentElement;
    toHTML(options?: HtmlRenderOptions): string;
    toHTMLDocument(options?: HtmlRenderOptions): string;
    toHTMLWithNavigation(options?: HtmlRenderOptions): string;
    private generateStyles;
    private generateNavigationStyles;
    getSlide(index: number): SlideElement | undefined;
    getLayout(layoutId: string): LayoutElement | undefined;
    getMaster(masterId: string): MasterElement | undefined;
    private escapeHtml;
}
export declare function createDocument(result: PptxParseResult): DocumentElement;
