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

interface RelationshipObject {
    type: string;
    target: string;
}
interface StyleTableItem {
    name: string;
    text: string;
    suffix?: string;
}
interface Callbacks {
    onFileStart?: () => void;
    onError?: (error: {
        type: string;
        message: string;
    }) => void;
    onSlide?: (data: any, info: {
        slideNum: number;
        fileName: string;
    }) => void;
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
interface PptxParserOptions {
    mediaProcess?: boolean;
    themeProcess?: boolean | 'colorsAndImageOnly';
    incSlide?: {
        width: number;
        height: number;
    };
    styleTable?: StyleTable;
    callbacks?: Callbacks;
}
interface SlideHtml {
    html: string;
    data: any;
    slideNum: number;
    fileName: string;
}
interface SlideJson {
    data: any;
    slideNum: number;
    fileName: string;
}
interface PptxHtmlResult {
    slides: SlideHtml[];
    slideSize: SlideSize;
    thumbnail: string | null;
    styles: {
        global: string;
    };
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
interface PptxJsonResult {
    slides: SlideJson[];
    slideSize: SlideSize;
    thumbnail: string | null;
    styles: {
        global: string;
    };
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
interface FileInfo {
    name: string;
    dir: boolean;
    size: number;
}
interface TextContent {
    type: 'text';
    content: string;
}
interface ImageContent {
    type: 'image';
    format: string;
    base64: string;
    dataUrl: string;
}
interface BinaryContent {
    type: 'binary';
    base64: string;
}
interface ErrorContent {
    type: 'error';
    error: string;
}
type FileContent = TextContent | ImageContent | BinaryContent | ErrorContent;
interface PptxFilesResult {
    files: FileInfo[];
    content: {
        [key: string]: FileContent;
    };
}
interface ChartDataPoint {
    x: string;
    y: number;
}
interface ChartSeries {
    key: string;
    values: ChartDataPoint[];
    xlabels: {
        [key: string]: string;
    };
}
interface ChartData {
    chartId: string;
    type: string;
    data: ChartSeries[];
}
interface ProcessedSlideData {
    slideLayoutContent: any;
    slideLayoutTables: any;
    slideMasterContent: any;
    slideMasterTables: any;
    slideContent: any;
    slideResObj: {
        [key: string]: RelationshipObject;
    };
    slideMasterTextStyles: any;
    layoutResObj: {
        [key: string]: RelationshipObject;
    };
    masterResObj: {
        [key: string]: RelationshipObject;
    };
    themeContent: any;
    themeResObj: {
        [key: string]: RelationshipObject;
    };
    diagramContent: any;
    diagramResObj: {
        [key: string]: RelationshipObject;
    };
    defaultTextStyle: any;
    tableStyles: any;
    styleTable: StyleTable;
    chartId: {
        value: number;
    };
    msgQueue: any[];
    bulletCounter: {
        [key: string]: number;
    };
    slideSize: SlideSize;
    index: number;
}
interface TextRun {
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
interface TextParagraph {
    text?: string;
    runs?: TextRun[];
    align?: 'left' | 'center' | 'right' | 'justify';
    bullet?: boolean;
}
type SlideElement = {
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
} | {
    type: 'shape';
    shapeType?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    fill?: {
        color: string;
    } | 'none';
    line?: {
        color: string;
        width?: number;
    } | 'none';
    rotation?: number;
    name?: string;
} | {
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
interface ComposerSlide {
    background?: string;
    elements?: SlideElement[];
}
interface PptxSerializeOptions {
    outputType?: 'uint8array' | 'arraybuffer' | 'blob' | 'nodebuffer' | 'base64';
}
interface PptxEditor {
    zip: any;
    getSlideCount(): Promise<number>;
    getSlide(slideNum: number): Promise<any>;
    deleteSlide(slideNum: number): Promise<void>;
    moveSlide(from: number, to: number): Promise<void>;
    setMetadata(metadata: Record<string, string>): Promise<void>;
    addSlide(slideJson: ComposerSlide): Promise<void>;
    save(options?: PptxSerializeOptions): Promise<Uint8Array>;
}
declare namespace pptxParser {
    const pptxToHtml: typeof pptxToHtml;
    const pptxToJson: typeof pptxToJson;
    const pptxToFiles: typeof pptxToFiles;
    const jsonToPptx: typeof jsonToPptx;
    const editPptx: typeof editPptx;
    const PPTXComposer: typeof PPTXComposer;
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
interface SlideSize {
    width: number;
    height: number;
    defaultTextStyle?: XmlNode;
}
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

export { PPTXComposer, pptxToHtml as default, editPptx, jsonToPptx, pptxParser, pptxToFiles, pptxToHtml, pptxToJson };
export type { BinaryContent, Callbacks, ChartData, ChartDataPoint, ChartSeries, ComposerSlide, ErrorContent, FileContent, FileInfo, ImageContent, PptxEditor, PptxFilesResult, PptxHtmlResult, PptxJsonResult, PptxParserOptions, PptxSerializeOptions, ProcessedSlideData, RelationshipObject, SlideElement, SlideHtml, SlideJson, SlideSize, StyleTable, StyleTableItem, TextContent, TextParagraph, TextRun };
