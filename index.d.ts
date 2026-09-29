
/**
 * 幻灯片大小信息
 */
export interface SlideSize {
    width: number;
    height: number;
    defaultTextStyle?: any;
}

/**
 * 关系对象
 */
export interface RelationshipObject {
    type: string;
    target: string;
}

/**
 * 样式表项
 */
export interface StyleTableItem {
    name: string;
    text: string;
    suffix?: string;
}

/**
 * 样式表
 */
export interface StyleTable {
    [key: string]: StyleTableItem;
}

/**
 * 回调函数接口
 */
export interface Callbacks {
    /**
     * 文件开始处理时的回调
     */
    onFileStart?: () => void;
    
    /**
     * 错误发生时的回调
     */
    onError?: (error: { type: string; message: string }) => void;
    
    /**
     * 处理完单个幻灯片时的回调
     */
    onSlide?: (data: any, info: { slideNum: number; fileName: string }) => void;
    
    /**
     * 获取缩略图时的回调
     */
    onThumbnail?: (thumbnail: string | null) => void;
    
    /**
     * 获取幻灯片大小时的回调
     */
    onSlideSize?: (slideSize: SlideSize) => void;
    
    /**
     * 获取全局CSS时的回调
     */
    onGlobalCSS?: (css: string) => void;
    
    /**
     * 处理完成时的回调
     */
    onComplete?: (info: {
        executionTime: number;
        slideWidth: number;
        slideHeight: number;
        styleTable: StyleTable;
        settings: PptxParserOptions;
    }) => void;
}

/**
 * PPTX解析选项
 */
export interface PptxParserOptions {
    /**
     * 是否处理媒体文件
     */
    mediaProcess?: boolean;
    
    /**
     * 主题处理方式
     */
    themeProcess?: boolean | 'colorsAndImageOnly';
    
    /**
     * 幻灯片尺寸调整
     */
    incSlide?: {
        width: number;
        height: number;
    };
    
    /**
     * 样式表
     */
    styleTable?: StyleTable;
    
    /**
     * 回调函数
     */
    callbacks?: Callbacks;
}

/**
 * 幻灯片HTML结果
 */
export interface SlideHtml {
    /**
     * 幻灯片HTML
     */
    html: string;
    
    /**
     * 幻灯片结构化数据（可用于后续处理）
     */
    data: any;
    
    /**
     * 幻灯片编号
     */
    slideNum: number;
    
    /**
     * 幻灯片文件名
     */
    fileName: string;
}

/**
 * 幻灯片JSON结果
 */
export interface SlideJson {
    /**
     * 幻灯片数据
     */
    data: any;
    
    /**
     * 幻灯片编号
     */
    slideNum: number;
    
    /**
     * 幻灯片文件名
     */
    fileName: string;
}

/**
 * PPTX转HTML结果
 */
export interface PptxHtmlResult {
    /**
     * 幻灯片HTML结果数组
     */
    slides: SlideHtml[];
    
    /**
     * 幻灯片大小信息
     */
    slideSize: SlideSize;
    
    /**
     * 缩略图
     */
    thumbnail: string | null;
    
    /**
     * 样式信息
     */
    styles: {
        /**
         * 全局CSS
         */
        global: string;
    };
    
    /**
     * 元数据
     */
    metadata: {
        /**
         * 标题
         */
        title?: string;
        /**
         * 主题
         */
        subject?: string;
        /**
         * 作者
         */
        author?: string;
        /**
         * 关键词
         */
        keywords?: string;
        /**
         * 描述
         */
        description?: string;
        /**
         * 最后修改者
         */
        lastModifiedBy?: string;
        /**
         * 创建日期
         */
        created?: string;
        /**
         * 修改日期
         */
        modified?: string;
        /**
         * 类别
         */
        category?: string;
        /**
         * 状态
         */
        status?: string;
        /**
         * 内容类型
         */
        contentType?: string;
        /**
         * 语言
         */
        language?: string;
        /**
         * 版本
         */
        version?: string;
    };
    
    /**
     * 图表数据
     */
    charts: ChartData[];
}

/**
 * PPTX转JSON结果
 */
export interface PptxJsonResult {
    /**
     * 幻灯片JSON结果数组
     */
    slides: SlideJson[];

    /**
     * 幻灯片大小信息
     */
    slideSize: SlideSize;

    /**
     * 缩略图
     */
    thumbnail: string | null;

    /**
     * 样式信息
     */
    styles: {
        /**
         * 全局CSS
         */
        global: string;
    };

    /**
     * 元数据
     */
    metadata: {
        /**
         * 标题
         */
        title?: string;
        /**
         * 主题
         */
        subject?: string;
        /**
         * 作者
         */
        author?: string;
        /**
         * 关键词
         */
        keywords?: string;
        /**
         * 描述
         */
        description?: string;
        /**
         * 最后修改者
         */
        lastModifiedBy?: string;
        /**
         * 创建日期
         */
        created?: string;
        /**
         * 修改日期
         */
        modified?: string;
        /**
         * 类别
         */
        category?: string;
        /**
         * 状态
         */
        status?: string;
        /**
         * 内容类型
         */
        contentType?: string;
        /**
         * 语言
         */
        language?: string;
        /**
         * 版本
         */
        version?: string;
    };

    /**
     * 图表数据
     */
    charts: ChartData[];
}

/**
 * 文件信息
 */
export interface FileInfo {
    /**
     * 文件路径
     */
    name: string;
    /**
     * 是否为目录
     */
    dir: boolean;
    /**
     * 解压后大小
     */
    size: number;
}

/**
 * 文本内容
 */
export interface TextContent {
    /**
     * 类型为 text
     */
    type: 'text';
    /**
     * 文本内容
     */
    content: string;
}

/**
 * 图片内容
 */
export interface ImageContent {
    /**
     * 类型为 image
     */
    type: 'image';
    /**
     * 图片格式
     */
    format: string;
    /**
     * Base64 编码
     */
    base64: string;
    /**
     * Data URL
     */
    dataUrl: string;
}

/**
 * 二进制内容
 */
export interface BinaryContent {
    /**
     * 类型为 binary
     */
    type: 'binary';
    /**
     * Base64 编码
     */
    base64: string;
}

/**
 * 错误内容
 */
export interface ErrorContent {
    /**
     * 类型为 error
     */
    type: 'error';
    /**
     * 错误信息
     */
    error: string;
}

/**
 * 文件内容（联合类型）
 */
export type FileContent = TextContent | ImageContent | BinaryContent | ErrorContent;

/**
 * PPTX转文件索引和内容结果
 */
export interface PptxFilesResult {
    /**
     * 文件索引列表
     */
    files: FileInfo[];
    /**
     * 文件内容映射
     */
    content: {
        [key: string]: FileContent;
    };
}

/**
 * 图表数据点
 */
export interface ChartDataPoint {
    /**
     * X坐标
     */
    x: string;
    /**
     * Y坐标
     */
    y: number;
}

/**
 * 图表系列
 */
export interface ChartSeries {
    /**
     * 系列名称
     */
    key: string;
    /**
     * 系列数据点
     */
    values: ChartDataPoint[];
    /**
     * X轴标签
     */
    xlabels: {
        [key: string]: string;
    };
}

/**
 * 图表数据
 */
export interface ChartData {
    /**
     * 图表ID
     */
    chartId: string;
    /**
     * 图表类型
     */
    type: string;
    /**
     * 图表数据
     */
    data: ChartSeries[];
}

/**
 * 处理后的幻灯片数据
 */
export interface ProcessedSlideData {
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
    chartId: { value: number };
    msgQueue: any[];
    bulletCounter: {
        [key: string]: number;
    };
    slideSize: SlideSize;
    index: number;
}

/**
 * PPTX转HTML转换器
 * @param fileData - PPTX文件数据
 * @param options - 转换选项
 * @returns 转换结果
 */
export declare function pptxToHtml(
    fileData: ArrayBuffer,
    options?: PptxParserOptions
): Promise<PptxHtmlResult | null>;

/**
 * PPTX转JSON转换器
 * @param fileData - PPTX文件数据
 * @param options - 转换选项
 * @returns 转换结果
 */
export declare function pptxToJson(
    fileData: ArrayBuffer,
    options?: PptxParserOptions
): Promise<PptxJsonResult | null>;

/**
 * PPTX转文件索引和内容转换器
 * @param fileData - PPTX文件数据
 * @returns 文件索引和内容结果
 */
export declare function pptxToFiles(
    fileData: ArrayBuffer
): Promise<PptxFilesResult>;

// ===========================================================================
// JSON → PPTX 序列化（Composer / jsonToPptx / editPptx）
// ===========================================================================

/**
 * 文本运行
 */
export interface TextRun {
    /** 运行文本 */
    text: string;
    /** 运行样式（覆盖元素级默认值） */
    options?: {
        fontSize?: number;
        color?: string;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        fontFace?: string;
        /** 超链接：http(s) 外部链接；'#N' 跳转到第 N 页 */
        href?: string;
    };
}

/**
 * 段落
 */
export interface TextParagraph {
    /** 段落文本（与 runs 二选一） */
    text?: string;
    /** 段落内运行列表（与 text 二选一） */
    runs?: TextRun[];
    /** 对齐：left / center / right / justify */
    align?: 'left' | 'center' | 'right' | 'justify';
    /** 是否使用项目符号 */
    bullet?: boolean;
}

/**
 * 幻灯片元素（联合类型）
 */
export type SlideElement =
    | {
          type: 'text';
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          /** 文本（\n 分段，与 runs/paragraphs 三选一） */
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
          /** 预设几何类型：rect / roundRect / ellipse / triangle 等 */
          shapeType?: string;
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          /** 填充：{color} 或 'none' */
          fill?: { color: string } | 'none';
          /** 边框：{color, width(pt)} 或 'none' */
          line?: { color: string; width?: number } | 'none';
          /** 旋转角度（度） */
          rotation?: number;
          name?: string;
      }
    | {
          type: 'image';
          x?: number;
          y?: number;
          width?: number;
          height?: number;
          /** 图片数据：dataURL 或 base64 字符串 */
          data?: string;
          /** 远程图片 URL（运行时通过 fetch 下载） */
          src?: string;
          /** 图片扩展名（data 为裸 base64 时用于推断格式） */
          extension?: string;
          href?: string;
          name?: string;
      };

/**
 * 序列化用幻灯片
 */
export interface ComposerSlide {
    /** 背景色（如 '#ffffff'） */
    background?: string;
    /** 元素列表，按数组顺序即图层顺序（后添加的在上层） */
    elements?: SlideElement[];
}

/**
 * 序列化用演示文稿 JSON 树
 */
export interface ComposerPresentation {
    /** 元数据（字段与 pptxToJson 返回的 metadata 互通） */
    metadata?: Record<string, string>;
    /** 幻灯片尺寸（px），默认 1280x720（16:9） */
    slideSize?: { width: number; height: number };
    /** 幻灯片列表 */
    slides: ComposerSlide[];
}

/**
 * 序列化选项
 */
export interface PptxSerializeOptions {
    /** JSZip 输出类型，默认 'uint8array'（可选 blob/nodebuffer/arraybuffer/base64） */
    outputType?: 'uint8array' | 'arraybuffer' | 'blob' | 'nodebuffer' | 'base64';
}

/**
 * JSON 转 PPTX 序列化器
 * @param presentation - 演示文稿 JSON 或 PPTXComposer 实例
 * @param options - 序列化选项
 * @returns PPTX 文件二进制数据
 */
export declare function jsonToPptx(
    presentation: ComposerPresentation | PPTXComposer,
    options?: PptxSerializeOptions
): Promise<Uint8Array>;

/**
 * 幻灯片构建器（addSlide 回调参数）
 */
export declare class SlideComposer {
    background(color: string): SlideComposer;
    addText(config: (t: any) => void | Record<string, any>): SlideComposer;
    addShape(config: (s: any) => void | Record<string, any>): SlideComposer;
    addImage(config: (i: any) => void | Record<string, any>): SlideComposer;
}

/**
 * 演示文稿流式构建器
 */
export declare class PPTXComposer {
    slideSize(width: number, height: number): PPTXComposer;
    metadata(metadata: Record<string, string>): PPTXComposer;
    title(value: string): PPTXComposer;
    author(value: string): PPTXComposer;
    subject(value: string): PPTXComposer;
    keywords(value: string): PPTXComposer;
    description(value: string): PPTXComposer;
    addSlide(config: (slide: SlideComposer) => void | ComposerSlide): PPTXComposer;
    toJSON(): ComposerPresentation;
    save(options?: PptxSerializeOptions): Promise<Uint8Array>;
}

/**
 * PPTX 编辑器（editPptx 返回值）
 */
export interface PptxEditor {
    /** 底层 JSZip 实例 */
    zip: any;
    /** 获取幻灯片数量 */
    getSlideCount(): Promise<number>;
    /** 获取指定页的简化 XML 树（与 pptxToJson 的 slideContent 同构） */
    getSlide(slideNum: number): Promise<any>;
    /** 删除指定页 */
    deleteSlide(slideNum: number): Promise<void>;
    /** 重排幻灯片 */
    moveSlide(from: number, to: number): Promise<void>;
    /** 写回元数据 */
    setMetadata(metadata: Record<string, string>): Promise<void>;
    /** 追加一页 */
    addSlide(slideJson: ComposerSlide): Promise<void>;
    /** 保存编辑结果 */
    save(options?: PptxSerializeOptions): Promise<Uint8Array>;
}

/**
 * 加载已有 PPTX 并返回编辑器
 * @param fileData - PPTX 文件数据
 * @returns 编辑器实例
 */
export declare function editPptx(fileData: ArrayBuffer | Uint8Array): Promise<PptxEditor>;

/**
 * PPTX解析器命名空间
 */
export declare namespace pptxParser {
    /**
     * PPTX转HTML转换器
     */
    const pptxToHtml: typeof import('./src/js/index').pptxToHtml;

    /**
     * PPTX转JSON转换器
     */
    const pptxToJson: typeof import('./src/js/index').pptxToJson;

    /**
     * PPTX转文件索引和内容转换器
     */
    const pptxToFiles: typeof import('./src/js/index').pptxToFiles;

    /**
     * JSON转PPTX序列化器
     */
    const jsonToPptx: typeof import('./src/js/index').jsonToPptx;

    /**
     * PPTX编辑器
     */
    const editPptx: typeof import('./src/js/index').editPptx;

    /**
     * 演示文稿流式构建器
     */
    const PPTXComposer: typeof import('./src/js/index').PPTXComposer;
}

/**
 * 全局PPTX解析器对象
 */
declare global {
    interface Window {
        pptxParser: {
            pptxToHtml: typeof import('./src/js/index').pptxToHtml;
            pptxToJson: typeof import('./src/js/index').pptxToJson;
            pptxToFiles: typeof import('./src/js/index').pptxToFiles;
            jsonToPptx: typeof import('./src/js/index').jsonToPptx;
            editPptx: typeof import('./src/js/index').editPptx;
            PPTXComposer: typeof import('./src/js/index').PPTXComposer;
        };
    }
}
