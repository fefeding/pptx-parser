import { BaseElement } from './BaseElement';
import type { RelsMap, PptRect } from '../types';
export interface DiagramData {
    colors?: Record<string, string>;
    data?: Record<string, any>;
    layout?: string;
    shapes?: DiagramShape[];
}
export interface DiagramShape {
    id: string;
    type: string;
    position?: {
        x: number;
        y: number;
    };
    size?: {
        width: number;
        height: number;
    };
    text?: string;
}
export declare class DiagramElement extends BaseElement {
    type: "diagram";
    diagramData?: DiagramData;
    relId: string;
    constructor(id: string, rect: PptRect, content?: any, props?: any, relsMap?: Record<string, any>);
    static fromNode(node: Element, relsMap: RelsMap): DiagramElement | null;
    private parseDiagramData;
    private parseShapes;
    toHTML(): string;
    private getDiagramInfo;
    toParsedElement(): any;
}
