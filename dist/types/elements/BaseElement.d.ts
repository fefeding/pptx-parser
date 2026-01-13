import type { PptRect, PptStyle, Position } from '../types';
export declare abstract class BaseElement {
    id: string;
    abstract type: string;
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    rawNode?: Element;
    isPlaceholder?: boolean;
    protected relsMap: Record<string, any>;
    zIndex?: number;
    idx?: number;
    constructor(id: string, type: string, rect: PptRect, content?: any, props?: any, relsMap?: Record<string, any>);
    abstract toHTML(): string;
    protected getContainerStyle(): string;
    protected getDataAttributes(): Record<string, string>;
    protected formatDataAttributes(): string;
    protected parsePosition(node: Element, tag?: string, namespace?: "http://schemas.openxmlformats.org/presentationml/2006/main"): Position;
    protected parseIdAndName(node: Element, nonVisualTag: string, namespace?: "http://schemas.openxmlformats.org/presentationml/2006/main"): {
        id: string;
        name: string;
        hidden: boolean;
    };
    protected generateId(): string;
    protected getAttributes(node: Element): Record<string, string>;
}
