export declare function parseTextContent(txBody: Element): string;
export declare function parseTextWithStyle(txBody: Element): Array<{
    text: string;
    style?: Record<string, unknown>;
}>;
export declare function parseTextStyle(rPr: Element): Record<string, unknown>;
