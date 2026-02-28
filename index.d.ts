
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
        };
    }
}
