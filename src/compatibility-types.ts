/**
 * 向后兼容类型（兼容旧版手写 index.d.ts 暴露的公开类型名）
 *
 * 这些类型由构建时的 rollup-plugin-dts 自动生成进 index.d.ts，
 * 保证下游消费者 import 的旧类型名仍然可用。
 *
 * - SlideSize / StyleTable / ComposerPresentation 直接复用源码定义（避免重名冲突）
 * - 其余为旧版策展类型的等价声明
 * - 结果类（PptxHtmlResult 等）与源码真实返回结构一致
 */

import type { SlideSize, StyleTable } from './index';
import type { ComposerPresentation } from './serializer/composer';

/** 关系对象 */
export interface RelationshipObject {
    type: string;
    target: string;
}

/** 样式表项 */
export interface StyleTableItem {
    name: string;
    text: string;
    suffix?: string;
}

/** 回调函数接口 */
export interface Callbacks {
    onFileStart?: () => void;
    onError?: (error: { type: string; message: string }) => void;
    onSlide?: (data: any, info: { slideNum: number; fileName: string }) => void;
    onThumbnail?: (thumbnail: string | null) => void;
    onSlideSize?: (slideSize: SlideSize) => void;
    onGlobalCSS?: (css: string) => void;
    onComplete?: (info: {
        executionTime: number;
        slideWidth: number;
        slideHeight: number;
        styleTable: StyleTable;
        settings: PptxParserOptions;
    }) => void;
}

/** PPTX 解析选项 */
export interface PptxParserOptions {
    mediaProcess?: boolean;
    themeProcess?: boolean | 'colorsAndImageOnly';
    incSlide?: { width: number; height: number };
    styleTable?: StyleTable;
    callbacks?: Callbacks;
}

/** 幻灯片 HTML 结果 */
export interface SlideHtml {
    html: string;
    data: any;
    slideNum: number;
    fileName: string;
}

/** 幻灯片 JSON 结果 */
export interface SlideJson {
    data: any;
    slideNum: number;
    fileName: string;
}

/** PPTX 转 HTML 结果 */
export interface PptxHtmlResult {
    slides: SlideHtml[];
    slideSize: SlideSize;
    thumbnail: string | null;
    styles: { global: string };
    metadata: {
        title?: string;
        subject?: string;
        author?: string;
        keywords?: string;
        description?: string;
        lastModifiedBy?: string;
        created?: string;
        modified?: string;
        category?: string;
        status?: string;
        contentType?: string;
        language?: string;
        version?: string;
    };
    charts: ChartData[];
}

/** PPTX 转 JSON 结果 */
export interface PptxJsonResult {
    slides: SlideJson[];
    slideSize: SlideSize;
    thumbnail: string | null;
    styles: { global: string };
    metadata: {
        title?: string;
        subject?: string;
        author?: string;
        keywords?: string;
        description?: string;
        lastModifiedBy?: string;
        created?: string;
        modified?: string;
        category?: string;
        status?: string;
        contentType?: string;
        language?: string;
        version?: string;
    };
    charts: ChartData[];
}

/** 文件信息 */
export interface FileInfo {
    name: string;
    dir: boolean;
    size: number;
}

/** 文本内容 */
export interface TextContent {
    type: 'text';
    content: string;
}

/** 图片内容 */
export interface ImageContent {
    type: 'image';
    format: string;
    base64: string;
    dataUrl: string;
}

/** 二进制内容 */
export interface BinaryContent {
    type: 'binary';
    base64: string;
}

/** 错误内容 */
export interface ErrorContent {
    type: 'error';
    error: string;
}

/** 文件内容（联合类型） */
export type FileContent = TextContent | ImageContent | BinaryContent | ErrorContent;

/** PPTX 转文件索引和内容结果 */
export interface PptxFilesResult {
    files: FileInfo[];
    content: { [key: string]: FileContent };
}

/** 图表数据点 */
export interface ChartDataPoint {
    x: string;
    y: number;
}

/** 图表系列 */
export interface ChartSeries {
    key: string;
    values: ChartDataPoint[];
    xlabels: { [key: string]: string };
}

/** 图表数据 */
export interface ChartData {
    chartId: string;
    type: string;
    data: ChartSeries[];
}

/** 处理后的幻灯片数据 */
export interface ProcessedSlideData {
    slideLayoutContent: any;
    slideLayoutTables: any;
    slideMasterContent: any;
    slideMasterTables: any;
    slideContent: any;
    slideResObj: { [key: string]: RelationshipObject };
    slideMasterTextStyles: any;
    layoutResObj: { [key: string]: RelationshipObject };
    masterResObj: { [key: string]: RelationshipObject };
    themeContent: any;
    themeResObj: { [key: string]: RelationshipObject };
    diagramContent: any;
    diagramResObj: { [key: string]: RelationshipObject };
    defaultTextStyle: any;
    tableStyles: any;
    styleTable: StyleTable;
    chartId: { value: number };
    msgQueue: any[];
    bulletCounter: { [key: string]: number };
    slideSize: SlideSize;
    index: number;
}

/** 文本运行 */
export interface TextRun {
    text: string;
    options?: {
        fontSize?: number;
        color?: string;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        fontFace?: string;
        href?: string;
    };
}

/** 段落 */
export interface TextParagraph {
    text?: string;
    runs?: TextRun[];
    align?: 'left' | 'center' | 'right' | 'justify';
    bullet?: boolean;
}

/** 幻灯片元素（联合类型） */
export type SlideElement =
    | {
          type: 'text';
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          text?: string;
          runs?: TextRun[];
          paragraphs?: TextParagraph[];
          align?: 'left' | 'center' | 'right' | 'justify';
          valign?: 'top' | 'middle' | 'bottom';
          fontSize?: number;
          color?: string;
          bold?: boolean;
          italic?: boolean;
          underline?: boolean;
          fontFace?: string;
          href?: string;
          name?: string;
      }
    | {
          type: 'shape';
          shapeType?: string;
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          fill?: { color: string } | 'none';
          line?: { color: string; width?: number } | 'none';
          rotation?: number;
          name?: string;
      }
    | {
          type: 'image';
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          data?: string;
          src?: string;
          extension?: string;
          href?: string;
          name?: string;
      };

/** 序列化用幻灯片 */
export interface ComposerSlide {
    background?: string;
    elements?: SlideElement[];
}

/** 序列化选项 */
export interface PptxSerializeOptions {
    outputType?: 'uint8array' | 'arraybuffer' | 'blob' | 'nodebuffer' | 'base64';
}

/** PPTX 编辑器（editPptx 返回值） */
export interface PptxEditor {
    zip: any;
    getSlideCount(): Promise<number>;
    getSlide(slideNum: number): Promise<any>;
    deleteSlide(slideNum: number): Promise<void>;
    moveSlide(from: number, to: number): Promise<void>;
    setMetadata(metadata: Record<string, string>): Promise<void>;
    addSlide(slideJson: ComposerSlide): Promise<void>;
    save(options?: PptxSerializeOptions): Promise<Uint8Array>;
}

/** PPTX 解析器命名空间 */
export declare namespace pptxParser {
    const pptxToHtml: typeof import('./index').pptxToHtml;
    const pptxToJson: typeof import('./index').pptxToJson;
    const pptxToFiles: typeof import('./index').pptxToFiles;
    const jsonToPptx: typeof import('./index').jsonToPptx;
    const editPptx: typeof import('./index').editPptx;
    const PPTXComposer: typeof import('./index').PPTXComposer;
}
