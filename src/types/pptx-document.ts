/**
 * PPTX JSON 标准格式（统一契约）
 *
 * 本文件定义「PPTX <-> JSON」双向互通的**权威 JSON 模型**，作为 pptxToJson（解析端）
 * 与 jsonToPptx（生成端）共同遵守的单一事实来源（single source of truth）。
 *
 * 设计目标：
 * - 解析端 pptxToJson 与生成端 jsonToPptx 使用同一套结构，实现 JSON 级 round-trip。
 * - 语义优先：元素以 type 区分（text/shape/image/chart/table/diagram），坐标统一为 px。
 * - 无损兜底：每个元素可选携带 __raw（原始 OOXML 片段），序列化时优先用语义字段，
 *   语义层未覆盖的细节回退到 __raw，保证往返不丢信息。
 *
 * 坐标单位：所有 x/y/width/height 均为 px（96 DPI）。
 *   解析端用 SLIDE_FACTOR(96/914400) 将 EMU 换算为 px；生成端用 pxToEmu 逆向换算。
 *
 * 与既有类型的关系：
 * - SerializerElement / SerializerSlide（serializer/element-builders.ts）是生成端当前实现，
 *   应逐步收敛到本文件的 PptxElement / PptxSlide。
 * - SlideElement / ComposerSlide（compatibility-types.ts）是下游兼容类型，结构与此一致，
 *   后续可改用此处的联合类型。
 *
 * @module types/pptx-document
 */

import type { SlideSize } from '../index';

/** 规范版本（递增，便于消费端判断能力） */
export type PptxDocumentVersion = '1.0';

/** 文档级元数据（与 pptxToJson 返回、Composer.metadata 互通） */
export interface PptxMetadata {
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
    [key: string]: string | undefined;
}

/** 文本对齐方式 */
export type TextAlign = 'left' | 'center' | 'right' | 'justify';
/** 垂直对齐方式 */
export type VAlign = 'top' | 'middle' | 'bottom';

/** 文本运行（单段内联样式） */
export interface PptxTextRun {
    text: string;
    fontSize?: number;     // pt
    color?: string;        // 颜色（#RRGGBB 或颜色名）
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    href?: string;         // 外部链接或内部跳转 '#N'
}

/** 段落（可显式 runs，或用 text 配合元素级默认样式） */
export interface PptxParagraph {
    text?: string;
    runs?: PptxTextRun[];
    align?: TextAlign;
    bullet?: boolean;
}

/** 形状填充：颜色串 / { color } / 'none'(无填充) / null(继承主题) */
export type PptxFill = string | { color?: string } | 'none' | null;

/** 形状边框：{ color, width(pt) } / 'none'(无边框) / null(继承) */
export type PptxLine = { color?: string; width?: number } | 'none' | null;

/** 背景填充 */
export type PptxBackground =
    | string                                   // 纯色（#RRGGBB 或颜色名），等价 { type:'solid', color }
    | { type: 'solid'; color: string }
    | { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[] }
    | { type: 'image'; data?: string; src?: string; extension?: string };

/** 幻灯片过渡效果（解析自 p:transition） */
export interface PptxTransition {
    type: string;        // fade/blind/cover/wipe/push/...
    duration: number;    // 毫秒
}

/** 图表系列 */
export interface PptxChartSeries {
    name?: string;
    values?: number[];   // 非散点图
    x?: number[];        // 散点图 X
    y?: number[];        // 散点图 Y
    color?: string;
}

/** 图表类型 */
export type PptxChartType =
    | 'barChart' | 'lineChart' | 'areaChart'
    | 'pieChart' | 'pie3DChart' | 'scatterChart' | string;

/** __raw 载荷：原始 OOXML 子树及其依赖关系（解析端产出，生成端原样回写） */
export interface PptxRawPayload {
    /** 原始节点标签名（如 p:graphicFrame / p:sp） */
    tag: string;
    /** 原始节点内容（tXml simplify 形态） */
    node: unknown;
    /** 节点引用的关系：旧 rId → { type, target, external }，回写时重新登记为新 rId */
    rels?: Record<string, { type: string; target: string; external?: boolean }>;
    /** 关系指向的部件内容（如 SmartArt 的 diagrams/*.xml），保证回写自包含 */
    parts?: { path: string; content?: string; base64?: string; contentType: string; media?: boolean }[];
}

/** 元素公共字段 */
interface PptxElementBase {
    /**
     * 原始 OOXML 载荷（解析端提取时附上）
     * - 语义层不支持的类型（diagram/组合/OLE 等）序列化时自动回退；
     * - 语义层已支持的类型需配合 rawFallback:true 才会回退。
     *
     * **依赖携带策略（重要）**：解析端默认只为「语义层不支持的类型」附带
     * `__raw.rels` / `__raw.parts`（避免每个图片元素都内联一份 base64 使 JSON 膨胀）。
     * 因此语义类型（text/shape/image/chart/table）默认只有 `{ tag, node }`，
     * 对其设置 rawFallback:true 会因缺少依赖而使 r:embed / r:id 等引用悬空。
     * 若确需对这些类型做原始回写，请用 `pptxToStandard(file, { rawDeps: 'all' })` 解析。
     */
    __raw?: PptxRawPayload;
    /**
     * 强制以 __raw 回写（即使 type 已受语义层支持）
     * 注意：需配合 rawDeps:'all' 解析出的载荷，否则引用类属性会悬空（见 __raw 说明）。
     */
    rawFallback?: boolean;
    /** 元素名称（可选，便于编辑区分） */
    name?: string;
}

/** 文本元素 */
export interface PptxTextElement extends PptxElementBase {
    type: 'text';
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    align?: TextAlign;
    valign?: VAlign;
    /** 文本来源三选一：paragraphs > runs > text */
    paragraphs?: PptxParagraph[];
    runs?: PptxTextRun[];
    text?: string;
    /** 元素级默认运行样式（作用于无显式样式的 run） */
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    href?: string;
}

/** 形状元素 */
export interface PptxShapeElement extends PptxElementBase {
    type: 'shape';
    shapeType: string;   // rect/roundRect/ellipse/... 见 OOXML prstGeom
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    fill?: PptxFill;
    line?: PptxLine;
}

/** 图片元素 */
export interface PptxImageElement extends PptxElementBase {
    type: 'image';
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    /** 内联数据：dataURL 或裸 base64 */
    data?: string;
    /** 远程 URL（生成端会下载为媒体） */
    src?: string;
    extension?: string;
    /** 图片级超链接（内部跳转用 '#N'） */
    href?: string;
}

/** 图表元素 */
export interface PptxChartElement extends PptxElementBase {
    type: 'chart';
    chartType: PptxChartType;
    x: number;
    y: number;
    width: number;
    height: number;
    title?: string;
    legend?: boolean;
    varyColors?: boolean;
    barDir?: 'bar' | 'col';
    categories?: string[];
    series?: PptxChartSeries[];
}

/** 表格单元格 */
export interface PptxTableCell {
    /** 单元格文本（无 runs 时使用） */
    text?: string;
    /** 富文本段落（优先于 text） */
    paragraphs?: PptxParagraph[];
    /** 跨列数（OOXML gridSpan，默认 1） */
    colSpan?: number;
    /** 跨行数（OOXML rowSpan，默认 1） */
    rowSpan?: number;
    /** 单元格底色 */
    fill?: string;
    /** 文本水平对齐 */
    align?: TextAlign;
    /** 文本垂直对齐（OOXML anchor） */
    valign?: VAlign;
    /** 单元格级文本样式 */
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
}

/** 表格行 */
export interface PptxTableRow {
    /** 行高（px，可选） */
    height?: number;
    cells: PptxTableCell[];
}

/**
 * 表格元素（解析自 p:graphicFrame/a:graphic/a:graphicData/a:tbl）
 *
 * 行/列尺寸可选：缺省时生成端按均分处理。
 */
export interface PptxTableElement extends PptxElementBase {
    type: 'table';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 列宽（px，长度即列数） */
    colWidths?: number[];
    /** 行高（px，长度即行数） */
    rowHeights?: number[];
    rows: PptxTableRow[];
}

/**
 * SmartArt / 图示元素（解析自 p:graphicFrame/a:graphicData[uri=diagram]）
 *
 * SmartArt 的几何布局由 diagrams/layoutN.xml 驱动，无法在语义层无损表达，
 * 因此标准格式仅保留可读文本内容（texts）与原始节点（__raw），
 * 其还原依赖 __raw 回退（语义层不保证保真）。
 */
export interface PptxDiagramElement extends PptxElementBase {
    type: 'diagram';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 图示数据部件（ppt/diagrams/dataN.xml）中的文本内容，按文档顺序 */
    texts?: string[];
    /** 数据部件路径（便于调试与二次读取） */
    dataPath?: string;
}

/** 幻灯片元素（联合类型） */
/**
 * 原始元素（解析兜底）
 *
 * 元素解析失败（未知标签 / 结构异常）时的占位：语义未知，仅由 __raw 承载原始节点，
 * 生成端始终按 __raw 原样回写，保证不丢信息。
 */
export interface PptxRawElement extends PptxElementBase {
    type: 'raw';
    x: number;
    y: number;
    width: number;
    height: number;
}

export type PptxElement =
    | PptxTextElement
    | PptxShapeElement
    | PptxImageElement
    | PptxChartElement
    | PptxTableElement
    | PptxDiagramElement
    | PptxRawElement;

/** 媒体资源（当元素不内联 data 时，通过 id 引用本表） */
export interface PptxMediaResource {
    /** dataURL 或裸 base64 */
    base64: string;
    mime: string;
}

/** 幻灯片 */
export interface PptxSlide {
    /** 背景（缺省继承主题） */
    background?: PptxBackground;
    /** 过渡效果（解析端自 p:transition 产出，生成端写回 p:transition） */
    transition?: PptxTransition;
    /** 演讲者备注（解析端自 notesContent 产出，生成端写回 notesSlide 部件） */
    notes?: string;
    elements: PptxElement[];
}

/** 可选主题覆盖（高级样式；标准 v1.0 暂为宽松结构，后续细化） */
export interface PptxTheme {
    [key: string]: unknown;
}

/**
 * PPTX 文档标准 JSON（双向统一格式）
 */
export interface PptxDocument {
    /** 规范版本 */
    version: PptxDocumentVersion;
    /** 幻灯片尺寸（px） */
    slideSize: SlideSize;
    /** 文档元数据 */
    metadata?: PptxMetadata;
    /** 幻灯片列表（顺序即显示顺序） */
    slides: PptxSlide[];
    /**
     * 媒体资源表（可选）：当元素用 { ref: '<id>' } 引用而非内联时生效。
     * 解析端默认内联 dataURL 到元素，故通常省略；生成端也可用此表避免重复内联。
     */
    media?: Record<string, PptxMediaResource>;
    /** 主题覆盖（可选，高级） */
    theme?: PptxTheme;
}
