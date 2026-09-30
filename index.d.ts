import JSZip from 'jszip';

interface RunStyle {
    align?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    href?: string;
    lang?: string;
}
interface TextRunSpec extends RunStyle {
    text?: string;
    options?: RunStyle;
}
interface ParagraphSpec {
    text?: string;
    runs?: TextRunSpec[];
    align?: string;
    bullet?: boolean;
}
interface ChartSeriesSpec {
    name?: string;
    values?: number[];
    x?: number[];
    y?: number[];
    color?: string;
}
interface SerializerElement {
    type?: string;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rotation?: number;
    text?: string;
    runs?: TextRunSpec[];
    paragraphs?: ParagraphSpec[];
    align?: string;
    valign?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    lang?: string;
    href?: string;
    fill?: string | {
        color?: string;
    } | null;
    line?: {
        color?: string;
        width?: number;
    } | 'none' | null;
    shapeType?: string;
    data?: string;
    src?: string;
    extension?: string;
    chartType?: string;
    categories?: string[];
    series?: ChartSeriesSpec[];
    varyColors?: boolean;
    barDir?: string;
    title?: string;
    legend?: boolean;
}
interface SerializerSlide {
    background?: string | null;
    elements?: SerializerElement[];
}

type ZipOutputType = 'base64' | 'string' | 'text' | 'binarystring' | 'array' | 'uint8array' | 'arraybuffer' | 'blob' | 'nodebuffer';
declare function jsonToPptx(presentation: unknown, options?: {
    outputType?: ZipOutputType;
}): Promise<string | ArrayBuffer | number[] | Uint8Array<ArrayBufferLike> | Blob | Buffer<ArrayBufferLike>>;
declare function editPptx(fileData: ArrayBuffer | Uint8Array | string): Promise<{
    zip: JSZip;
    save: (options?: {
        outputType?: ZipOutputType;
    }) => Promise<string | ArrayBuffer | number[] | Uint8Array<ArrayBufferLike> | Blob | Buffer<ArrayBufferLike>>;
    getSlideCount(): Promise<number>;
    getSlide(slideNum: number): Promise<any>;
    deleteSlide(slideNum: number): Promise<void>;
    moveSlide(from: number, to: number): Promise<void>;
    setMetadata(metadata: Record<string, unknown>): Promise<void>;
    addSlide(slideJson: SerializerSlide): Promise<void>;
}>;

type FluentBuilder = {
    [key: string]: (value?: unknown) => FluentBuilder;
};
type ElementConfig = ((builder: FluentBuilder) => void) | Record<string, unknown>;
type ChartConfig = ((el: SerializerElement) => void) | Record<string, unknown>;
interface ComposerPresentation {
    metadata: Record<string, unknown>;
    slideSize: {
        width: number;
        height: number;
    };
    slides: SerializerSlide[];
}
declare class SlideComposer {
    slide: SerializerSlide;
    constructor();
    background(color: string): this;
    addText(config: ElementConfig): this;
    addShape(config: ElementConfig): this;
    addImage(config: ElementConfig): this;
    addChart(config: ChartConfig): this;
}
declare class PPTXComposer {
    presentation: ComposerPresentation;
    constructor();
    slideSize(width: number | {
        width: number;
        height: number;
    }, height?: number): this;
    metadata(metadata: Record<string, unknown>): this;
    title(value: string): this;
    author(value: string): this;
    subject(value: string): this;
    keywords(value: string): this;
    description(value: string): this;
    addSlide(config: ((slide: SlideComposer) => void) | SerializerSlide): this;
    toJSON(): any;
    save(options?: {
        outputType?: ZipOutputType;
    }): Promise<string | ArrayBuffer | number[] | Uint8Array<ArrayBufferLike> | Blob | Buffer<ArrayBufferLike>>;
}

interface XmlAttrs {
    order?: number;
    [name: string]: any;
}
interface XmlNode {
    attrs?: XmlAttrs;
    [tagName: string]: any;
}
interface ParseCallbacks {
    onFileStart?: () => void;
    onSlide?: (html: string | Record<string, unknown>, info: {
        slideNum: number;
        fileName: string;
    }) => void;
    onThumbnail?: (thumbnail: string) => void;
    onSlideSize?: (slideSize: {
        width: number;
        height: number;
    }) => void;
    onGlobalCSS?: (css: string) => void;
    onComplete?: (info: {
        executionTime: number;
        slideWidth: number;
        slideHeight: number;
        styleTable?: unknown;
        settings?: unknown;
    }) => void;
    onError?: (err: {
        type: string;
        message: string;
    }) => void;
}
interface ParseSettings {
    themeProcess?: boolean | string;
    mediaProcess?: boolean;
    incSlide?: {
        width: number;
        height: number;
    };
    styleTable?: Record<string, unknown>;
    callbacks?: ParseCallbacks;
}

interface ChartQueueItem {
    type: string;
    data: Record<string, unknown>;
}
interface StyleTableEntry {
    name: string;
    suffix?: string;
    text: string;
}
type StyleTable = Record<string, StyleTableEntry>;
type ResourceMap = Record<string, Record<string, string>>;
interface PptxMetadata {
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
}
interface SlideDataRecord {
    index: number;
    slideContent: XmlNode;
    slideLayoutContent?: XmlNode;
    slideMasterContent?: XmlNode;
    themeContent?: XmlNode;
    diagramContent?: XmlNode | string | null;
    slideLayoutTables?: XmlNode;
    slideMasterTables?: XmlNode;
    slideMasterTextStyles?: unknown;
    tableStyles?: XmlNode;
    slideResObj?: ResourceMap;
    layoutResObj?: ResourceMap;
    masterResObj?: ResourceMap;
    themeResObj?: ResourceMap;
    diagramResObj?: ResourceMap;
    styleTable?: StyleTable;
    chartId?: {
        value: number;
    };
    msgQueue?: ChartQueueItem[];
    bulletCounter?: unknown;
    defaultTextStyle?: XmlNode | null;
    [key: string]: unknown;
}
interface HtmlSlideResult {
    html: string;
    data: SlideDataRecord | undefined;
    slideNum: number | undefined;
    fileName: string | undefined;
}
interface JsonSlideResult {
    data: SlideDataRecord | undefined;
    slideNum: number | undefined;
    fileName: string | undefined;
}
interface FileIndexEntry {
    name: string;
    dir: boolean;
    size: number;
}
type PptxFileData = ArrayBuffer | Uint8Array | string;
declare function pptxToHtml(fileData: PptxFileData, options: Partial<ParseSettings>): Promise<{
    slides: HtmlSlideResult[];
    slideSize: {
        width: number;
        height: number;
        defaultTextStyle: XmlNode;
    };
    thumbnail: string | null;
    styles: {
        global: string;
    };
    metadata: PptxMetadata;
    charts: Array<Record<string, unknown>>;
} | null>;
declare function pptxToJson(fileData: PptxFileData, options: Partial<ParseSettings>): Promise<{
    slides: JsonSlideResult[];
    slideSize: {
        width: number;
        height: number;
        defaultTextStyle: XmlNode;
    };
    thumbnail: string | null;
    styles: {
        global: string;
    };
    metadata: PptxMetadata;
    charts: Array<Record<string, unknown>>;
} | null>;
declare function pptxToFiles(fileData: PptxFileData): Promise<{
    files: FileIndexEntry[];
    content: Record<string, unknown>;
}>;

export { PPTXComposer, pptxToHtml as default, editPptx, jsonToPptx, pptxToFiles, pptxToHtml, pptxToJson };
