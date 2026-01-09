export declare function getAttrs(node: Element): Record<string, string>;
export declare function getChildrenByTagNS(parent: Element | null, tagName: string, namespaceURI: string): Element[];
export declare function getFirstChildByTagNS(parent: Element | null, tagName: string, namespaceURI: string): Element | null;
export declare function getAttrSafe(element: Element | null, attrName: string, defaultValue?: string): string;
export declare function getBoolAttr(element: Element | null, attrName: string): boolean;
