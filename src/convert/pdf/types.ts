import type { AssetSource, DocAsset, DocAssetMode } from '../core/asset';

/**
 * PDF 文档结构类型定义
 *
 * 解析输出为 PdfDocument（含逐页文本与坐标），生成输入为 PdfDocumentSpec（含元素序列）。
 *
 * @module pdf/types
 */

/** 文档元数据 */
export interface PdfMetadata {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    creator?: string;
    producer?: string;
    /** 创建时间（ISO 字符串） */
    created?: string;
    /** 修改时间（ISO 字符串） */
    modified?: string;
    /** PDF 版本，如 '1.7' */
    version?: string;
}

/** 文本片段（带坐标，PDF 坐标系：原点左下、单位 pt） */
export interface PdfTextItem {
    /** 文本内容 */
    text: string;
    /** 左下角 X */
    x: number;
    /** 左下角 Y */
    y: number;
    /** 宽度 */
    width: number;
    /** 高度 */
    height: number;
    /** 字号 */
    fontSize: number;
    /** 字体名 */
    fontName?: string;
    /** 该片段后是否紧跟换行 */
    hasEOL?: boolean;
}

/** 矢量线段（PDF 坐标系：原点左下、单位 pt），由 operator list 解析而来 */
export interface PdfVector {
    /** 线段起点 X */
    x1: number;
    /** 线段起点 Y */
    y1: number;
    /** 线段终点 X */
    x2: number;
    /** 线段终点 Y */
    y2: number;
    /** 是否为水平线（|y1-y2| 极小） */
    horizontal: boolean;
    /** 是否为垂直线（|x1-x2| 极小） */
    vertical: boolean;
}

/**
 * 页面图像（含版面位置），仅当解析选项 `images: true` 时返回
 *
 * PDF 的图文都是绝对定位的，只有带上坐标才能在转 md / html 时把图插回阅读顺序。
 * `data` 是可用字节：原始像素已编码成 PNG，原本就是 JPEG/PNG 的直通不重新编码。
 */
export interface PdfImage {
    /** 页内对象 id（同一张图被多次绘制时用于去重） */
    id: string;
    /** 左下角 X（pt，PDF 坐标系：原点左下、y 向上） */
    x: number;
    /** 左下角 Y（pt） */
    y: number;
    /** 页面上的显示宽度（pt） */
    width: number;
    /** 页面上的显示高度（pt） */
    height: number;
    /** 原始像素宽度 */
    pixelWidth: number;
    /** 原始像素高度 */
    pixelHeight: number;
    /** 图片字节 */
    data: Uint8Array;
    /** MIME，如 image/png / image/jpeg */
    mime: string;
}

/** 页面 */
export interface PdfPage {
    /** 页号（1 起） */
    index: number;
    /** 宽度（pt） */
    width: number;
    /** 高度（pt） */
    height: number;
    /** 旋转角度（0/90/180/270） */
    rotation: number;
    /** 整页纯文本 */
    text?: string;
    /** 文本片段（含坐标） */
    items?: PdfTextItem[];
    /** 矢量线段（表格线等），仅当解析选项 `vectors: true` 时返回 */
    vectors?: PdfVector[];
    /** 页面图像（含位置与字节），仅当解析选项 `images: true` 时返回 */
    images?: PdfImage[];
}

/** 解析结果文档 */
export interface PdfDocument {
    metadata?: PdfMetadata;
    /** 总页数 */
    pageCount: number;
    /** 页面详情（仅当请求 detailed 时返回） */
    pages?: PdfPage[];
    /**
     * 内嵌资源（附件 / 页面图像），仅当 `assets` 选项非 `ignore` 时返回
     *
     * 附件锚点为 `kind: 'other'`，页面图像为 `kind: 'page'` 并记录页码。
     */
    assets?: DocAsset[];
}

/** 解析选项 */
export interface PdfParseOptions {
    /**
     * 内嵌资源提取方式，默认 `ignore`（不提取）
     *
     * 页面图像需要走渲染指令，代价较高（上百页文档会明显变慢），
     * 因此只有显式开启才提取。附件的提取很廉价，开了就一起给。
     */
    assets?: DocAssetMode;
    /** 是否返回逐页文本，默认 true */
    text?: boolean;
    /** 是否返回带坐标的文本片段，默认 false（体积较大） */
    items?: boolean;
    /** 是否额外返回矢量线段（表格线/分隔线），用于版面分析与表格还原，默认 false */
    vectors?: boolean;
    /**
     * 是否额外返回页面图像（含位置与字节），默认 false
     *
     * 需要等 pdfjs 把图像对象异步解析完，比矢量略慢；图文混排的文档转 md / html
     * 必须要它，否则整篇图都丢。
     */
    images?: boolean;
    /**
     * 单张图的像素上限，超出后降采样（默认 2e6，约 1600x1250）；0 表示不降采样
     *
     * 扫描件整页图常达 3000x4000，原样编码会让产物体积爆炸。
     */
    maxImagePixels?: number;
    /** 只解析指定页（1 起），缺省全部 */
    pages?: number[];
    /** 最大返回页数（配合分页预览场景） */
    maxPages?: number;
    /** 中文字体资源目录 URL/路径；缺省自动探测 pdfjs-dist 内的标准字体 */
    standardFontDataUrl?: string;
    /** cMap 资源目录（Node 端文件路径，须以 / 结尾）；缺省自动探测 pdfjs-dist */
    cMapUrl?: string;
    /** wasm 资源目录（JPEG2000 / JBIG2 图像解码）；缺省自动探测 pdfjs-dist */
    wasmUrl?: string;
    /**
     * 内嵌资源的 pdf worker 文件路径（Node 端）。
     *
     * 构建脚本可把 cMap/标准字体/wasm 以 base64 内嵌进 pdf worker（如 VS Code
     * 扩展的 media/pdf.worker.min.mjs）。打包后的扩展没有 node_modules，自动
     * 探测必然失败；传入该文件后，解析前会把资源解包到系统临时目录（按内容
     * 哈希缓存，只做一次），宿主侧的解析/转换才能渲染中文等 CID 字体。
     */
    embeddedWorkerPath?: string;
    /**
     * pdfjs worker 地址（Node 端）。
     *
     * pdfjs v5+ 的 fake worker 必须有 workerSrc 才会加载，否则直接抛
     *「No "GlobalWorkerOptions.workerSrc" specified」。它在源码里的内置默认值是
     * 相对路径 "./pdf.worker.mjs"——当 pdfjs 被打包器内联进 bundle 时，这个相对路径
     * 会指向 bundle 所在目录而不是 pdfjs-dist 目录，worker 必然加载失败。
     *
     * 浏览器端不需要传（由打包器把 worker 变成资源文件）；Node 端若被内联打包，
     * 请显式传入 worker 文件的绝对地址（file:// URL 或路径均可）。
     */
    workerSrc?: string;
}

/** 解析结果 */
export interface PdfResult {
    /** 文档 */
    document: PdfDocument;
    /** 解析提示 */
    messages: string[];
}

/** 纸张尺寸（pt） */
export type PdfPageSize = 'A4' | 'A3' | 'A5' | 'Letter' | 'Legal' | [number, number] | { width: number; height: number };

/** 文本对齐 */
export type PdfAlign = 'left' | 'center' | 'right';

/** 文本元素 */
export interface PdfTextElement {
    type: 'text';
    text: string;
    /** 距左边距（pt，默认 0） */
    x?: number;
    /** 距顶边距（pt，默认 0） */
    y?: number;
    /** 字号（pt，默认 12） */
    size?: number;
    /** 字体名（需已注册，默认随 bold/italic 变化） */
    font?: string;
    bold?: boolean;
    italic?: boolean;
    /** 颜色，#RRGGBB 或 'black'/'red' 等关键字 */
    color?: string;
    align?: PdfAlign;
    /** 文本框宽度（pt），给定后支持换行与对齐 */
    width?: number;
    /** 行高倍数，默认 1.35 */
    lineHeight?: number;
    /** 段前间距（pt），自动分页时生效 */
    spacingBefore?: number;
    /** 段后间距（pt），自动分页时生效 */
    spacingAfter?: number;
}

/** 富文本片段（同一行内可混排多段不同样式的文本） */
export interface PdfRun {
    /** 文本片段内容（不含换行） */
    text: string;
    /** 字号（pt），缺省跟随元素 size */
    size?: number;
    bold?: boolean;
    italic?: boolean;
    /** 颜色，#RRGGBB 或关键字 */
    color?: string;
    /** 等宽（行内代码）样式，渲染时加灰色底纹 */
    code?: boolean;
    /** 超链接地址，渲染时加下划线并着色 */
    link?: string;
    /** 删除线 */
    strike?: boolean;
}

/** 制表位 */
export interface PdfTabStop {
    /** 相对段落左边缘的位置（pt）；缺省为段落右端（目录右对齐页码的典型用法） */
    pos?: number;
    /** 对齐方式，默认 left */
    align?: 'left' | 'center' | 'right';
    /** 前导符：目录条目与标题的「点线导引」用 dot */
    leader?: 'dot' | 'hyphen' | 'underscore' | 'middleDot';
}

/**
 * 富文本元素：一段由多个 run 组成的段落，支持内联样式混排。
 *
 * 与 `PdfTextElement` 的区别：text 只能整段单一字号/颜色，runs 可逐片段混排。
 * 生成端会自行按宽度断行（中文按字、西文按空格），因此调用方无需预先换行。
 * run 文本中的 `\t` 按 tabStops 跳到制表位（未给 tabStops 时按普通空格处理）。
 */
export interface PdfRunsElement {
    type: 'runs';
    runs: PdfRun[];
    /** 距左边距（pt，默认 0） */
    x?: number;
    /** 距顶边距（pt，默认 0） */
    y?: number;
    /** 缺省字号（pt，默认跟随文档 defaultSize） */
    size?: number;
    align?: PdfAlign;
    /** 文本框宽度（pt），缺省用内容区宽度 */
    width?: number;
    /** 文本框高度（pt）：仅在需要按框中心旋转时使用 */
    height?: number;
    /** 行高倍数，默认 1.35 */
    lineHeight?: number;
    /** 段后间距（pt），默认 4 */
    spacingAfter?: number;
    /** 段前间距（pt），自动分页时生效 */
    spacingBefore?: number;
    /** 左缩进（pt）：绘制起点右移并相应收窄行宽 */
    indentLeft?: number;
    /** 右缩进（pt）：行宽相应收窄 */
    indentRight?: number;
    /** 首行缩进（pt）：仅段落第一行右移 */
    indentFirstLine?: number;
    /** 制表位（pos 相对段落左边缘），用于目录「文本 …… 页码」排版 */
    tabStops?: PdfTabStop[];
    /** 绕文本框中心旋转（度，逆时针为正） */
    rotation?: number;
}

/** 标题元素 */
export interface PdfHeadingElement extends Omit<PdfTextElement, 'type'> {
    type: 'heading';
    /** 标题层级 1-6 */
    level: 1 | 2 | 3 | 4 | 5 | 6;
}

/** 图片元素 */
export interface PdfImageElement {
    type: 'image';
    /**
     * 图片数据：data URL / base64 / 二进制，或资源引用 `{ assetId }`
     *
     * 引用写法需配合 `jsonToPdf(doc, { assets })`，
     * 适合「解析 → 改 → 生成」时直接复用解析出的字节，不必转成 data URL。
     */
    src: string | ArrayBuffer | Uint8Array | AssetSource;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    /**
     * 源图裁剪比例（0~1，表示从各边裁掉的比例）
     *
     * PPTX 的「图片裁剪」（a:srcRect）会让同一张图在幻灯片上只显示一部分。
     * 生成端按裁剪窗口放大绘制并用裁剪路径限制在框内。
     */
    crop?: { l?: number; r?: number; t?: number; b?: number };
}

/** 线段元素 */
export interface PdfLineElement {
    type: 'line';
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    /** 线宽（pt），默认 1 */
    lineWidth?: number;
    color?: string;
    /** 虚线间隔，如 [3, 2] */
    dash?: number[];
    /** 起点箭头（PPT 连接线的 startArrow） */
    startArrow?: boolean;
    /** 终点箭头（PPT 连接线的 endArrow） */
    endArrow?: boolean;
}

/**
 * 水平分隔线（流式）：在当前内容位置画一条横贯内容区的细线。
 *
 * 与 `line` 元素的区别：`line` 的 y 坐标是距页面顶部的绝对偏移（不跟随流式光标），
 * 而 `hr` 始终绘制在文档流的当前位置，适合 Markdown 的 `---` 等分隔语义。
 */
export interface PdfHrElement {
    type: 'hr';
    /** 线宽（pt），默认 0.5 */
    lineWidth?: number;
    /** 上方留白（pt），默认 4 */
    spacingBefore?: number;
    /** 下方留白（pt），默认 4 */
    spacingAfter?: number;
    color?: string;
}

/** 矩形元素 */
export interface PdfRectElement {
    type: 'rect';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 填充色 */
    fill?: string;
    /** 描边色 */
    stroke?: string;
    /** 描边宽度（pt） */
    lineWidth?: number;
    /** 圆角半径（pt） */
    radius?: number;
    /** 绘制椭圆（而非矩形）；宽高不等时为真正的椭圆 */
    ellipse?: boolean;
    /** 绕自身中心旋转（度，顺时针）；PDF 生成端用变换矩阵实现 */
    rotation?: number;
    /** 填充不透明度（0~1，缺省 1）：用于渐变停点的透明度 */
    opacity?: number;
}

/**
 * 路径段（贝塞尔路径）
 *
 * 坐标为「元素坐标系」：原点在内容区左上角、y 轴向下（与 rect 元素的
 * x/y/width/height 同一坐标系），生成端绘制时再翻转到 PDF 的 y 向上坐标系。
 */
export interface PdfPathSegment {
    /** m=移动到 / l=直线 / c=三次贝塞尔（6 个参数）/ h=闭合 */
    op: 'm' | 'l' | 'c' | 'h';
    args: number[];
}

/**
 * 组合元素：先按裁剪路径裁剪，再绘制子元素
 *
 * PPTX 的渐变填充在圆角 / 椭圆形状上需要「渐变 + 轮廓」同时表达：
 * 渐变只能用纯色条带近似（PDF 无原生渐变），条带必须被裁剪在形状轮廓内，
 * 否则会溢出成矩形。裁剪路径为元素坐标系下的贝塞尔路径。
 */
export interface PdfGroupElement {
    type: 'group';
    /** 裁剪路径（元素坐标系）；缺省则不裁剪 */
    clip?: PdfPathSegment[];
    /** 子元素（坐标为页面绝对坐标，与 rect 元素一致） */
    children: PdfElement[];
}

/**
 * 任意路径元素（多边形 / 圆柱等预设形状）
 *
 * 坐标为内容区左上原点、y 向下、pt（与 rect / group 的裁剪路径一致），
 * 生成端绘制时翻转到 PDF 的 y 向上坐标系。用于 PPTX 的块箭头
 * （rightArrow / downArrow）等无法用 rect / line 表达的预设形状。
 */
export interface PdfPathElement {
    type: 'path';
    segments: PdfPathSegment[];
    fill?: string;
    stroke?: string;
    /** 线宽（pt） */
    lineWidth?: number;
    /** 填充 / 描边不透明度（0~1） */
    opacity?: number;
    /** 虚线间隔（pt），如 [3, 2] */
    dash?: number[];
    /** 起点箭头（沿路径起点方向） */
    startArrow?: boolean;
    /** 终点箭头（沿路径终点方向） */
    endArrow?: boolean;
}

/**
 * 带合并信息的表格单元格
 *
 * Word 的合并单元格（w:gridSpan 横跨列、w:vMerge 纵跨行）与列宽只有在这一层
 * 保留下来，生成端才能按 tblGrid 网格对齐；否则合并表头会被按「第 N 个
 * cell」分配宽度，退化成窄列并把文字逐字折行。
 */
export interface PdfTableCellObject {
    /** 单元格内容：元素序列或纯文本 */
    content: PdfElement[] | string;
    /** 跨列数（对应 w:gridSpan），默认 1 */
    colspan?: number;
    /** 跨行数（对应 w:vMerge 向下合并的行数），默认 1 */
    rowspan?: number;
    /** 显式宽度（pt，对应 w:tcW）；优先于 columnWidths 推导 */
    width?: number;
    /** 底纹填充色（对应 w:shd） */
    fill?: string;
    /** 垂直对齐（对应 w:vAlign），默认垂直居中 */
    verticalAlign?: 'top' | 'middle' | 'bottom';
}

/** 表格单元格：纯文本简写 / 元素序列 / 带合并信息的对象 */
export type PdfTableCell = string | PdfElement[] | PdfTableCellObject;

/** 表格元素 */
export interface PdfTableElement {
    type: 'table';
    x?: number;
    y?: number;
    /** 各列宽度（pt），缺省按内容区等分 */
    columnWidths?: number[];
    /** 行数据（行优先） */
    rows: PdfTableCell[][];
    /** 首行是否作为表头（默认 true） */
    header?: boolean;
    /** 单元格字号（pt），默认 10 */
    fontSize?: number;
    /** 单元格内边距（pt），默认 4 */
    padding?: number;
    /** 行高（pt），缺省按字号自适应 */
    rowHeight?: number;
    /** 边框色，设为 'none' 则无边框 */
    borderColor?: string;
    /** 表头背景色 */
    headerFill?: string;
}

/** 分页符 */
export interface PdfPageBreakElement {
    type: 'pageBreak';
}

/** PDF 元素联合类型 */
export type PdfElement =
    | PdfTextElement
    | PdfHeadingElement
    | PdfRunsElement
    | PdfImageElement
    | PdfLineElement
    | PdfHrElement
    | PdfRectElement
    | PdfPathElement
    | PdfGroupElement
    | PdfTableElement
    | PdfPageBreakElement;

/** 页面定义 */
export interface PdfPageSpec {
    /** 纸张尺寸，默认 A4 */
    size?: PdfPageSize;
    /** 方向，默认纵向 */
    orientation?: 'portrait' | 'landscape';
    /** 页边距（pt），默认 56（约 2cm） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /** 本页元素；缺省时使用文档级 elements */
    elements?: PdfElement[];
}

/** 生成文档规格 */
export interface PdfDocumentSpec {
    /**
     * 可用资源表：用于解析图片元素的 `{ assetId }` 引用
     *
     * 与 `jsonToPdf` 的 `assets` 选项同效，规格内自带时无需再传。
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
    metadata?: PdfMetadata;
    /** 页面；缺省时按 elements 自动分页 */
    pages?: PdfPageSpec[];
    /** 文档级元素（自动分页时使用） */
    elements?: PdfElement[];
    /** 内嵌字体：{ regular, bold, italic, boldItalic }，用于中文或自定义字体 */
    fonts?: PdfFontSet;
    /** 默认字号（pt），默认 12 */
    defaultSize?: number;
    /** 默认正文颜色 */
    defaultColor?: string;
    /** 纸张尺寸（自动分页模式），默认 A4 */
    pageSize?: PdfPageSize;
    /** 页边距（pt，自动分页模式），默认 56（约 2cm） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /**
     * 自动分页模式下每个元素绘制前的回调（pageNumber 从 1 开始）。
     *
     * 文档互转链路用它做两遍排版：第一遍收集标题实际页码，回填目录页码后
     * 再生成第二遍（页码占位不影响行数，故两遍分页结果一致）。
     */
    onElementLayout?: (element: PdfElement, pageNumber: number) => void;
}

/** 内嵌字体集合 */
export interface PdfFontSet {
    regular?: string | ArrayBuffer | Uint8Array;
    bold?: string | ArrayBuffer | Uint8Array;
    italic?: string | ArrayBuffer | Uint8Array;
    boldItalic?: string | ArrayBuffer | Uint8Array;
}

/** 生成选项 */
export interface PdfGenerateOptions {
    /**
     * 可用资源表：用于解析图片元素上的 `{ assetId }` 引用
     *
     * @example
     * ```js
     * const parsed = await pdfToJson(file, { assets: 'extract' });
     * const bytes  = await jsonToPdf(parsed, { assets: parsed.assets });
     * ```
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
    /** 是否压缩内容流，默认 true */
    compress?: boolean;
    /** 中文/自定义字体，与文档级 fonts 合并 */
    fonts?: PdfFontSet;
}
