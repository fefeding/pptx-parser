import JSZip from 'jszip';
import type { TagsResult } from './types';
export declare function parseAllSlideTags(zip: JSZip, slideCount: number): Promise<TagsResult[]>;
export declare function parseSlideTags(zip: JSZip, slideId: string): Promise<TagsResult>;
export declare function findSlidesByTag(tagsResults: TagsResult[], tagName: string, tagValue?: string): string[];
export declare function findSlidesByProperty(tagsResults: TagsResult[], propName: string, propValue?: any): string[];
