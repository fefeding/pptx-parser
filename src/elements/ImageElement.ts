/**
 * 图片元素类
 * 支持图片、视频、音频
 * 对齐 PPTXjs 的媒体解析能力
 */

import { BaseElement } from './BaseElement';
import { getFirstChildByTagNS, getAttrSafe, getBoolAttr, emu2px } from '../utils';
import { NS } from '../constants';
import type { ParsedImageElement, RelsMap } from '../types';

/**
 * 媒体类型
 */
export type MediaType = 'image' | 'video' | 'audio';

/**
 * 视频信息
 */
export interface VideoInfo {
  src: string;
  type: 'blob' | 'link';
  format?: string;
  autoplay?: boolean;
  muted?: boolean;
  controls?: boolean;
  loop?: boolean;
}

/**
 * 音频信息
 */
export interface AudioInfo {
  src: string;
  type: 'blob' | 'link';
  format?: string;
  autoplay?: boolean;
  muted?: boolean;
  controls?: boolean;
  loop?: boolean;
}

/**
 * 图片元素类
 */
export class ImageElement extends BaseElement {
  type = 'image' as const;

  /** 媒体类型 */
  mediaType: MediaType = 'image';

  /** 图片URL或Base64 */
  src: string = '';

  /** 图片关联ID */
  relId: string = '';

  /** MIME类型 */
  mimeType?: string;

  /** 替代文本 */
  altText?: string;

  /** 视频信息（如果是视频） */
  videoInfo?: VideoInfo;

  /** 音频信息（如果是音频） */
  audioInfo?: AudioInfo;

  /**
   * 从XML节点创建图片元素
   */
  static fromNode(node: Element, relsMap: RelsMap): ImageElement | null {
    try {
      const element = new ImageElement('', { x: 0, y: 0, width: 0, height: 0 }, '', '', {}, {}, relsMap);

      // 解析ID和名称
      const nvPicPr = getFirstChildByTagNS(node, 'nvPicPr', NS.p);
      const cNvPr = nvPicPr ? getFirstChildByTagNS(nvPicPr, 'cNvPr', NS.p) : null;

      element.id = getAttrSafe(cNvPr, 'id', element.generateId());
      element.name = getAttrSafe(cNvPr, 'name', '');
      element.hidden = getBoolAttr(cNvPr, 'hidden');
      element.altText = getAttrSafe(cNvPr, 'descr', '');

      // 解析位置尺寸
      const spPr = getFirstChildByTagNS(node, 'spPr', NS.p);
      element.rect = element.parsePosition(node);

      // 解析图片关联ID
      const blipFill = getFirstChildByTagNS(node, 'blipFill', NS.p);
      let relId = '';

      if (blipFill) {
        const blip = getFirstChildByTagNS(blipFill, 'blip', NS.a);
        relId = blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed') || '';
      }

      // 检查是否是视频或音频
      const nvPr = nvPicPr ? getFirstChildByTagNS(nvPicPr, 'nvPr', NS.p) : null;
      const videoFile = nvPr ? getFirstChildByTagNS(nvPr, 'videoFile', NS.a) : null;
      const audioFile = nvPr ? getFirstChildByTagNS(nvPr, 'audioFile', NS.a) : null;

      if (videoFile) {
        // 处理视频
        element.mediaType = 'video';
        element.videoInfo = element.parseVideoInfo(videoFile, relId, relsMap);
        element.src = element.videoInfo.src;
        element.relId = relId;
      } else if (audioFile) {
        // 处理音频
        element.mediaType = 'audio';
        element.audioInfo = element.parseAudioInfo(audioFile, relId, relsMap);
        element.src = element.audioInfo.src;
        element.relId = relId;
      } else {
        // 处理图片
        element.mediaType = 'image';
        element.src = element.parseImageSrc(relId, relsMap, node);
        element.relId = relId;

        // 从路径推断MIME类型
        if (element.src.endsWith('.png')) element.mimeType = 'image/png';
        else if (element.src.endsWith('.jpg') || element.src.endsWith('.jpeg')) element.mimeType = 'image/jpeg';
        else if (element.src.endsWith('.gif')) element.mimeType = 'image/gif';
        else if (element.src.endsWith('.svg')) element.mimeType = 'image/svg+xml';
      }

      element.content = {
        src: element.src,
        alt: element.altText,
        mediaType: element.mediaType,
        videoInfo: element.videoInfo,
        audioInfo: element.audioInfo
      };

      element.props = {
        imgId: relId,
        mimeType: element.mimeType,
        mediaType: element.mediaType
      };

      element.rawNode = node;

      return element;
    } catch (error) {
      console.error('Failed to parse image element:', error);
      return null;
    }
  }

  /**
   * 解析视频信息
   */
  private parseVideoInfo(videoFile: Element, relId: string, relsMap: RelsMap): VideoInfo {
    const link = videoFile.getAttribute('link');
    const videoRid = videoFile.getAttributeNS(NS.r, 'link') || '';

    if (link) {
      // 外部链接视频
      return {
        src: link,
        type: 'link',
        autoplay: true,
        muted: false,
        controls: true,
        loop: false
      };
    } else {
      // 本地视频文件
      const videoPath = relsMap[videoRid]?.target || '';
      return {
        src: videoPath,
        type: 'blob',
        format: this.detectVideoFormat(videoPath),
        autoplay: false,
        muted: false,
        controls: true,
        loop: false
      };
    }
  }

  /**
   * 解析音频信息
   */
  private parseAudioInfo(audioFile: Element, relId: string, relsMap: RelsMap): AudioInfo {
    const link = audioFile.getAttribute('link');
    const audioRid = audioFile.getAttributeNS(NS.r, 'link') || '';

    if (link) {
      // 外部链接音频
      return {
        src: link,
        type: 'link',
        autoplay: true,
        muted: false,
        controls: true,
        loop: false
      };
    } else {
      // 本地音频文件
      const audioPath = relsMap[audioRid]?.target || '';
      return {
        src: audioPath,
        type: 'blob',
        format: this.detectAudioFormat(audioPath),
        autoplay: false,
        muted: false,
        controls: true,
        loop: false
      };
    }
  }

  /**
   * 解析图片源
   */
  private parseImageSrc(relId: string, relsMap: RelsMap, node: Element): string {
    if (!relId || !relsMap[relId]) return '';

    const target = relsMap[relId].target;
    return target;
  }

  /**
   * 检测视频格式
   */
  private detectVideoFormat(path: string): string | undefined {
    if (path.endsWith('.mp4')) return 'mp4';
    if (path.endsWith('.webm')) return 'webm';
    if (path.endsWith('.ogg')) return 'ogg';
    return undefined;
  }

  /**
   * 检测音频格式
   */
  private detectAudioFormat(path: string): string | undefined {
    if (path.endsWith('.mp3')) return 'mp3';
    if (path.endsWith('.wav')) return 'wav';
    if (path.endsWith('.ogg')) return 'ogg';
    return undefined;
  }

  /**
   * 转换为HTML
   */
  toHTML(): string {
    const style = this.getContainerStyle();

    if (this.mediaType === 'video' && this.videoInfo) {
      return this.toVideoHTML(style);
    } else if (this.mediaType === 'audio' && this.audioInfo) {
      return this.toAudioHTML(style);
    } else {
      return this.toImageHTML(style);
    }
  }

  /**
   * 转换为图片HTML
   */
  private toImageHTML(containerStyle: string): string {
    const imgStyle = [
      `width: 100%`,
      `height: 100%`,
      `object-fit: contain`
    ].join('; ');

    return `<div style="${containerStyle}">
      <img src="${this.src}" alt="${this.altText || ''}" style="${imgStyle}" />
    </div>`;
  }

  /**
   * 转换为视频HTML
   */
  private toVideoHTML(containerStyle: string): string {
    const video = this.videoInfo!;
    const videoStyle = [
      `width: 100%`,
      `height: 100%`,
      `object-fit: contain`
    ].join('; ');

    if (video.type === 'link') {
      // YouTube, Vimeo 等外部链接
      return `<div style="${containerStyle}">
        <iframe
          src="${video.src}"
          style="${videoStyle}"
          frameborder="0"
          allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture"
          allowfullscreen>
        </iframe>
      </div>`;
    } else {
      // 本地视频
      return `<div style="${containerStyle}">
        <video
          src="${video.src}"
          style="${videoStyle}"
          ${video.autoplay ? 'autoplay' : ''}
          ${video.muted ? 'muted' : ''}
          ${video.loop ? 'loop' : ''}
          ${video.controls ? 'controls' : ''}>
          ${this.getVideoSourceTag(video.format)}
        </video>
      </div>`;
    }
  }

  /**
   * 获取 video source 标签
   */
  private getVideoSourceTag(format?: string): string {
    if (!format) return '';

    const formats: Record<string, string> = {
      mp4: 'video/mp4',
      webm: 'video/webm',
      ogg: 'video/ogg'
    };

    const mimeType = formats[format];
    if (mimeType) {
      return `<source src="${this.videoInfo?.src}" type="${mimeType}" />`;
    }

    return '';
  }

  /**
   * 转换为音频HTML
   */
  private toAudioHTML(containerStyle: string): string {
    const audio = this.audioInfo!;
    const audioStyle = [
      `width: 100%`,
      `display: block`
    ].join('; ');

    if (audio.type === 'link') {
      // 外部链接音频
      return `<div style="${containerStyle}">
        <audio
          src="${audio.src}"
          style="${audioStyle}"
          ${audio.autoplay ? 'autoplay' : ''}
          ${audio.muted ? 'muted' : ''}
          ${audio.loop ? 'loop' : ''}
          ${audio.controls ? 'controls' : ''}>
        </audio>
      </div>`;
    } else {
      // 本地音频
      return `<div style="${containerStyle}">
        <audio id="audio_${this.id}" style="${audioStyle}" ${audio.controls ? 'controls' : ''}>
          ${this.getAudioSourceTag(audio.format)}
        </audio>
      </div>`;
    }
  }

  /**
   * 获取 audio source 标签
   */
  private getAudioSourceTag(format?: string): string {
    if (!format) return '';

    const formats: Record<string, string> = {
      mp3: 'audio/mpeg',
      wav: 'audio/wav',
      ogg: 'audio/ogg'
    };

    const mimeType = formats[format];
    if (mimeType) {
      return `<source src="${this.audioInfo?.src}" type="${mimeType}" />`;
    }

    return '';
  }

  /**
   * 转换为ParsedImageElement格式
   */
  toParsedElement(): ParsedImageElement {
    return {
      id: this.id,
      type: 'image',
      rect: this.rect,
      style: this.style,
      content: this.content,
      props: this.props,
      name: this.name,
      hidden: this.hidden,
      relId: this.relId,
      src: this.src,
      altText: this.altText,
      mimeType: this.mimeType,
      attrs: this.getAttributes(this.rawNode!),
      rawNode: this.rawNode
    };
  }
}
