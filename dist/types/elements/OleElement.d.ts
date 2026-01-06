import { BaseElement } from './BaseElement';
import type { ParsedOleElement, RelsMap } from '../types';
export declare class OleElement extends BaseElement {
    type: "ole";
    progId?: string;
    relId: string;
    oleName?: string;
    hasFallback?: boolean;
    static fromNode(node: Element, relsMap: RelsMap): OleElement | null;
    constructor(id: string, rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, progId: string, relId: string, props?: any, relsMap?: Record<string, any>);
    toHTML(): string;
    toParsedElement(): ParsedOleElement;
}
