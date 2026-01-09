import type { RelsMap } from '../core/types';
export declare function parseRels(relsXml: string): RelsMap;
export declare function normalizeTargetPath(relsFilePath: string, target: string): string;
export declare function parseRelsWithBase(relsXml: string, relsFilePath: string): RelsMap;
export declare function parseMetadata(coreXml: string): {
    title?: string;
    author?: string;
    created?: string;
    modified?: string;
    subject?: string;
    keywords?: string;
};
export declare function parseBackgroundColor(slideXml: string): string;
export declare function parseSlideSize(slideLayoutXml: string): {
    width: number;
    height: number;
};
