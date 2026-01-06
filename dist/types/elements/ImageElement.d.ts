import { BaseElement } from './BaseElement';
import type { ParsedImageElement, RelsMap } from '../types';
export type MediaType = 'image' | 'video' | 'audio';
export interface VideoInfo {
    src: string;
    type: 'blob' | 'link';
    format?: string;
    autoplay?: boolean;
    muted?: boolean;
    controls?: boolean;
    loop?: boolean;
}
export interface AudioInfo {
    src: string;
    type: 'blob' | 'link';
    format?: string;
    autoplay?: boolean;
    muted?: boolean;
    controls?: boolean;
    loop?: boolean;
}
export declare class ImageElement extends BaseElement {
    type: "image";
    mediaType: MediaType;
    src: string;
    relId: string;
    mimeType?: string;
    altText?: string;
    videoInfo?: VideoInfo;
    audioInfo?: AudioInfo;
    static fromNode(node: Element, relsMap: RelsMap): ImageElement | null;
    private parseVideoInfo;
    private parseAudioInfo;
    private parseImageSrc;
    private detectVideoFormat;
    private detectAudioFormat;
    toHTML(): string;
    private toImageHTML;
    private toVideoHTML;
    private getVideoSourceTag;
    private toAudioHTML;
    private getAudioSourceTag;
    toParsedElement(): ParsedImageElement;
}
