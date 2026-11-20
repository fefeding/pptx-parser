/**
 * DOCX 文档结构类型定义
 *
 * 解析（docxToJson）与生成（jsonToDocx）共用这一套结构，
 * 使 DOCX 也能像 PPTX 一样做「解析 → JSON → 生成」的往返。
 *
 * @module docx/types
 */

import type { AssetSource, DocAsset, DocAssetMode } from '../core/asset';
import type { ChartSpec } from '../chart/svg';

/** 文本级样式（对应一次 run） */
export interface DocxRun {
    /** 文本内容 */
    text: string;
    /** 加粗 */
    bold?: boolean;
    /** 斜体 */
    italic?: boolean;
    /** 下划线 */
    underline?: boolean;
    /** 删除线 */
    strike?: boolean;
    /** 下标 */
    subscript?: boolean;
    /** 上标 */
    superscript?: boolean;
    /** 文字颜色，#RRGGBB */
    color?: string;
    /** 字号（磅），如 12 表示 12pt */
    size?: number;
    /** 字体名 */
    font?: string;
    /** 背景高亮色，#RRGGBB */
    highlight?: string;
    /** 超链接地址 */
    link?: string;
    /** 行内代码（等宽样式） */
    code?: boolean;
}

/** 段落对齐方式 */
export type DocxAlign = 'left' | 'center' | 'right' | 'justify';

/** 标题块 */
export interface DocxHeading {
    type: 'heading';
    /** 标题层级 1-6 */
    level: 1 | 2 | 3 | 4 | 5 | 6;
    /** 纯文本简写 */
    text?: string;
    /** 富文本（与 text 二选一，text 优先） */
    runs?: DocxRun[];
    /**
     * 自动编号文本（如 "1." / "3.1."，来自 w:numPr + numbering.xml）
     *
     * 仅供展示：渲染时前置到标题前，还原 Word/WPS 里由列表编号生成的章节号。
     * 生成端忽略该字段（重新写出编号定义超出结构化契约的范围）。
     */
    number?: string;
    /** 段落级属性（对齐、缩进、间距） */
    align?: DocxAlign;
    indentLeft?: number;
    indentRight?: number;
    indentFirstLine?: number;
    spacingBefore?: number;
    spacingAfter?: number;
    lineSpacing?: number;
}

/** 目录条目 */
export interface DocxTocEntry {
    /** 层级 1-3（对应 toc 1 / toc 2 / toc 3 样式） */
    level: number;
    /** 条目文本（不含页码，可能自带 "1." 这样的章节号） */
    text: string;
    /** 页码（TOC 域缓存的 PAGEREF 结果，缺失时留空） */
    page?: number;
}

/**
 * 目录块（TOC 域的缓存结果）
 *
 * Word/WPS 把目录存成一条 `TOC \o "1-3"` 域：域结果是一串 toc N 样式段落，
 * 文本形如 `1.\t背景\t2`（章节号 + Tab + 标题 + Tab + 页码）。解析时把这一串
 * 折叠为单个 toc 块，渲染时再按层级缩进 + 点线导引 + 右对齐页码还原。
 */
export interface DocxToc {
    type: 'toc';
    entries: DocxTocEntry[];
}

/** 普通段落块 */
export interface DocxParagraph {
    type: 'paragraph';
    text?: string;
    runs?: DocxRun[];
    align?: DocxAlign;
    /** 段落缩进（磅） */
    indentLeft?: number;
    indentRight?: number;
    /** 首行缩进（磅） */
    indentFirstLine?: number;
    /** 段前间距（磅） */
    spacingBefore?: number;
    /** 段后间距（磅） */
    spacingAfter?: number;
    /** 行距倍数，1 为单倍行距 */
    lineSpacing?: number;
}

/** 列表项：字符串简写或带子项的富文本 */
export type DocxListItem = string | (DocxRun[] & { children?: DocxListItem[] });

/** 列表块 */
export interface DocxList {
    type: 'list';
    /** true 为有序（编号）列表，false 为无序（项目符号）列表 */
    ordered?: boolean;
    /** 列表项 */
    items: DocxListItem[];
    /**
     * 编号格式：decimal / lowerLetter / lowerRoman / bullet …
     *
     * Word 的有序列表未必是「1. 2. 3.」，中文文档里 a. b. c. 与 i. ii. iii. 极常见。
     * 只保留 ordered: boolean 会一律渲染成阿拉伯数字，与原文不符。
     */
    numFmt?: string;
    /** 文本左缩进（磅），来自 numbering 的 w:ind/@w:left */
    indentLeft?: number;
    /** 悬挂缩进（磅），编号自身占位宽度，来自 w:ind/@w:hanging */
    hanging?: number;
}

/** 单元格边框（w:tcBorders 单边） */
export interface DocxCellBorder {
    /** 线型，如 single / double / nil（无边框） */
    style?: string;
    /** 颜色 #RRGGBB */
    color?: string;
    /** 宽度（磅） */
    size?: number;
}

/** 表格单元格对象形态：携带合并、边框、底纹等属性（对应 w:tcPr） */
export interface DocxCellDetail {
    type: 'cell';
    /** 单元格内容块（多个段落） */
    blocks?: DocxBlock[];
    /** 富文本简写（单段落时） */
    runs?: DocxRun[];
    /** 纯文本简写 */
    text?: string;
    /** 跨列数（w:gridSpan） */
    colspan?: number;
    /** 跨行数（w:vMerge 向下合并的行数） */
    rowspan?: number;
    /** 单元格宽度（磅，w:tcW） */
    width?: number;
    /** 底纹填充色（w:shd fill） */
    fill?: string;
    /** 垂直对齐（w:vAlign） */
    verticalAlign?: 'top' | 'middle' | 'bottom';
    /** 四边边框；缺省沿用表格默认边框 */
    borders?: {
        top?: DocxCellBorder;
        right?: DocxCellBorder;
        bottom?: DocxCellBorder;
        left?: DocxCellBorder;
    };
}

/** 表格单元格：字符串简写 / 富文本 / 带属性对象 */
export type DocxCell = string | DocxRun[] | DocxCellDetail;

/** 表格块 */
export interface DocxTable {
    type: 'table';
    /** 二维单元格内容，行优先（被合并覆盖的位置不产单元格） */
    rows: DocxCell[][];
    /** 首行是否为表头（生成时加粗并重复表头） */
    header?: boolean;
    /** 列宽（磅），长度需与列数一致 */
    columnWidths?: number[];
    /** 表格边框宽度（磅），0 表示无边框 */
    borderSize?: number;
}

/** 图片块 */
export interface DocxImage {
    type: 'image';
    /**
     * 图片数据：data URL / base64 / 二进制，或资源引用 `{ assetId }`
     *
     * 解析时若 `assets` 选项为 `extract` / `both`，这里会是资源引用，
     * 字节在 `document.assets` 中按 id 取；`inline` 模式则为内联 data URL。
     */
    src: string | ArrayBuffer | Uint8Array | AssetSource;
    /**
     * 资源 ID：当图片存放在文档资源表（`document.assets`）中时给出，
     * 配合 `jsonToDocx` 的 `assets` 选项即可按引用写回，字节不必重复传递
     */
    assetId?: string;
    /** 宽度（磅） */
    width?: number;
    /** 高度（磅） */
    height?: number;
    /** 图注文字 */
    caption?: string;
    /** 图片说明（alt 文本） */
    alt?: string;
}

/** 分页符 */
export interface DocxPageBreak {
    type: 'pageBreak';
}

/** 分隔线（水平线） */
export interface DocxHorizontalRule {
    type: 'hr';
}

/** 引用块 */
export interface DocxQuote {
    type: 'quote';
    text?: string;
    runs?: DocxRun[];
}

/** 代码块（等宽字体 + 底纹） */
export interface DocxCode {
    type: 'code';
    text: string;
    /** 代码语言名（仅作标注） */
    language?: string;
}

/** 文本框块（wps 文本框 / VML 文本框，常用于代码片段、标注框） */
export interface DocxTextBox {
    type: 'textBox';
    /** 文本框内的段落块 */
    paragraphs: DocxBlock[];
    /** 填充色 #RRGGBB */
    fill?: string;
    /** 边框色 #RRGGBB */
    stroke?: string;
    /** 宽度（磅，wp:extent） */
    width?: number;
    /** 高度（磅） */
    height?: number;
}

/**
 * 图表块（来自 XLSX / PPTX 的原生图表）
 *
 * 只存格式无关的绘图描述（ChartSpec），矢量图由 `chart/svg` 统一渲染：
 * HTML 目标内嵌 <svg>（可无损缩放、可打印），其余目标退化成 `fallback` 数据表。
 */
export interface DocxChart {
    type: 'chart';
    /** 绘图描述（类型/系列/配色等） */
    spec: ChartSpec;
    /** 渲染宽度（px） */
    width?: number;
    /** 渲染高度（px） */
    height?: number;
    /** 图注 */
    caption?: string;
    /** 无法画矢量图的目标（md / docx / pdf）退化成这张数据表 */
    fallback?: DocxTable;
}

/** 所有支持的块类型 */
export type DocxBlock =
    | DocxHeading
    | DocxParagraph
    | DocxList
    | DocxTable
    | DocxImage
    | DocxPageBreak
    | DocxHorizontalRule
    | DocxQuote
    | DocxCode
    | DocxTextBox
    | DocxChart
    | DocxToc;

/** 文档元数据 */
export interface DocxMetadata {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    creator?: string;
    lastModifiedBy?: string;
    company?: string;
}

/** 页面设置（w:sectPr，单位磅） */
export interface DocxSection {
    /** 页宽（磅） */
    width: number;
    /** 页高（磅） */
    height: number;
    orientation?: 'portrait' | 'landscape';
    /** 上边距（磅） */
    marginTop: number;
    /** 下边距（磅） */
    marginBottom: number;
    /** 左边距（磅） */
    marginLeft: number;
    /** 右边距（磅） */
    marginRight: number;
    /**
     * 文档网格行距（磅），来自 w:docGrid/@w:linePitch（twips/20）
     *
     * 中文 Word/WPS 常用「行网格」（w:docGrid w:type="lines"）：每行行高对齐到
     * linePitch 的整数倍。不还原它，行距会明显偏小，每页能塞下更多内容，
     * 分页落点比 Word/WPS 整体前移。
     */
    gridLinePitch?: number;
}

/** 完整 DOCX 文档（JSON 契约） */
export interface DocxDocument {
    /** 文档元数据 */
    metadata?: DocxMetadata;
    /** 页面设置（纸张与页边距） */
    section?: DocxSection;
    /** 内容块序列 */
    blocks: DocxBlock[];
    /**
     * 文档内二进制资源（图片等）
     *
     * 仅在解析时 `assets` 选项为 `extract` / `both` 时填充。
     * 块上通过 `assetId` 引用这里，生成时把同一份 `assets` 传给 `jsonToDocx` 即可还原，
     * 避免把 base64 在 JSON 里搬来搬去。
     */
    assets?: DocAsset[];
}

/** 解析选项 */
export interface DocxParseOptions {
    /**
     * 自定义 mammoth styleMap，把 Word 内置样式映射到语义标签。
     * 传入的映射会与内置映射合并。
     */
    styleMap?: string[] | Record<string, string>;
    /** 是否使用 mammoth 内置样式映射（默认 true） */
    includeDefaultStyleMap?: boolean;
    /** 图片转换策略：'data-url'（内联 base64）或 'ignore'（丢弃） */
    convertImage?: 'data-url' | 'ignore';
    /**
     * 二进制资源（图片等）处理方式，默认 `inline`（保持既有行为，块里是 data URL）
     *
     * - `inline`：块上是 data URL，浏览器直接能显示，但 JSON 体积膨胀
     * - `extract`：块上只留 `assetId`，字节进 `document.assets`，
     *              适合后台/服务端存储与另存为（可用 `assetToNodeStream` 直接落盘）
     * - `both`：两者都给——前端直接显示 `src`，后台按 `assetId` 拿原始字节
     * - `ignore`：完全丢弃
     */
    assets?: DocAssetMode;
}

/** 解析结果（含警告信息） */
export interface DocxParseResult<T> {
    /** 解析产物 */
    value: T;
    /** mammoth 转换过程中的警告 */
    messages: string[];
    /** 纯文本内容（便于全文检索） */
    rawText?: string;
}

/** 生成选项 */
export interface DocxGenerateOptions {
    /** 纸张尺寸，默认 A4 */
    pageSize?: 'A4' | 'Letter' | 'A3' | 'A5' | 'Legal' | { width: number; height: number };
    /** 页面方向，默认纵向 */
    orientation?: 'portrait' | 'landscape';
    /** 页边距（磅），默认 72（1 英寸） */
    margin?: number | { top?: number; right?: number; bottom?: number; left?: number };
    /** 默认正文字体 */
    defaultFont?: string;
    /** 默认字号（磅） */
    defaultSize?: number;
    /**
     * 可用资源表：用于解析块上的资源引用（`{ assetId }`）
     *
     * 典型用法是「解析 → 改 → 生成」时把 `docxToJson` 返回的 `assets` 原样传回，
     * 图片字节便不必转成 data URL 再从 JSON 里取回来。
     *
     * @example
     * ```js
     * const parsed = await docxToJson(file, { assets: 'extract' });
     * const bytes  = await jsonToDocx(parsed, { assets: parsed.assets });
     * ```
     */
    assets?: DocAsset[] | Record<string, DocAsset>;
}