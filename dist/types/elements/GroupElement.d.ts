import { BaseElement } from './BaseElement';
import type { BaseElement as BaseElementType } from './BaseElement';
import type { ParsedGroupElement, RelsMap } from '../types';
export declare class GroupElement extends BaseElement {
    type: "group";
    children: BaseElementType[];
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    childOffset?: {
        x: number;
        y: number;
    };
    static fromNode(node: Element, relsMap: RelsMap): GroupElement | null;
    private parseGraphicFrame;
    private parseGroupProperties;
    toHTML(): string;
    private getGroupStyle;
    toParsedElement(): ParsedGroupElement;
}
