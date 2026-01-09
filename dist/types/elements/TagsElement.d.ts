import { BaseElement } from './BaseElement';
import type { TagsResult, SlideTag, CustomProperty, ExtensionData } from '../core/types';
export declare class TagsElement extends BaseElement {
    type: 'tags';
    tags: SlideTag[];
    extensions: ExtensionData[];
    customProperties: CustomProperty[];
    slideId?: string;
    constructor(id: string, tags?: SlideTag[], customProperties?: CustomProperty[], extensions?: ExtensionData[], props?: any);
    static fromResult(result: TagsResult): TagsElement;
    toHTML(): string;
    toHTMLDebug(): string;
    getTag(name: string): string | undefined;
    getProperty(name: string): any | undefined;
    setTag(name: string, value: string): void;
    setProperty(name: string, value: any, type?: 'string' | 'number' | 'boolean' | 'date'): void;
}
