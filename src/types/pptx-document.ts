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
    /** 列表样式：true=项目符号；'number'=自动编号；{ type:'number', fmt?, start? } 或 { type:'bullet', char? } */
    bullet?: boolean | 'number' | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    /** 行距：数字=百分比(100=单倍) 或 { type:'pt', value } / { type:'percent', value } */
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    /** 段前间距 pt */
    spaceBefore?: number;
    /** 段后间距 pt */
    spaceAfter?: number;
    /** 左缩进 pt（a:pPr@marL） */
    indentLeft?: number;
    /** 右缩进 pt（a:pPr@marR） */
    indentRight?: number;
    /** 悬挂缩进 pt（a:pPr@indent，项目符号相对文本的缩进） */
    indent?: number;
}

/** 渐变填充色标 */
export interface PptxGradientStop { color: string; position: number; }
/** 纯色填充（可带透明度 0-100） */
export interface PptxFillSolid { type?: 'solid'; color?: string; transparency?: number; }
/** 渐变填充 */
export interface PptxFillGradient { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: PptxGradientStop[]; }
/** 形状填充：颜色串 / {color} / {type:'solid',...} / {type:'gradient',...} / 'none' / null */
export type PptxFill = string | PptxFillSolid | PptxFillGradient | 'none' | null;

/**
 * 图片填充的源图裁剪（a:srcRect）：从图片各边裁掉的比例，取值 0~1。
 * 四者皆为 0 表示不裁剪（等于不传）。
 */
export interface PptxImageSrcRect { l?: number; t?: number; r?: number; b?: number; }
/**
 * 图片填充的平铺（a:tile）：sx/sy 为每格占图片原始尺寸的比例，tx/ty 为平铺偏移，取值 0~1。
 * 指定 tile 即使用平铺，否则为拉伸铺满（a:stretch）。
 */
export interface PptxImageTile { sx?: number; sy?: number; tx?: number; ty?: number; }

/** 形状边框：{ color, width(pt), transparency, dashType } / 'none'(无边框) / null(继承) */
export interface PptxLineStyle { color?: string; width?: number; transparency?: number; dashType?: string; }
export type PptxLine = PptxLineStyle | 'none' | null;

/** 形状阴影 */
export interface PptxShadow {
    type?: 'outer' | 'inner';
    color?: string;
    blur?: number;          // pt
    distance?: number;      // pt
    angle?: number;         // 度
    transparency?: number; // 0-100
}
/** 形状发光 */
export interface PptxGlow { color?: string; blur?: number; } // blur 单位 pt
/** 形状特效集合（对应 a:effectLst） */
export interface PptxShapeEffects {
    shadow?: PptxShadow | boolean;  // true = 默认外阴影
    glow?: PptxGlow | boolean;      // true = 默认发光
}

/** 背景填充 */
export type PptxBackground =
    | string                                   // 纯色（#RRGGBB 或颜色名），等价 { type:'solid', color }
    | { type: 'solid'; color: string }
    | { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[] }
    | {
        type: 'image'; data?: string; src?: string; extension?: string;
        /** 源图裁剪（a:srcRect），0~1 */
        srcRect?: PptxImageSrcRect;
        /** 平铺（a:tile），0~1；不传则为拉伸铺满 */
        tile?: PptxImageTile;
    };

/** 幻灯片过渡效果（解析自 p:transition） */
export interface PptxTransition {
    type: string;        // fade/blind/cover/wipe/push/...
    duration: number;    // 毫秒
    /** 是否允许点击切换（默认 true；false 表示仅自动播放） */
    advanceOnClick?: boolean;
}

/** 元素进入动画（解析自 p:timing 的 p:spTgt） */
export interface PptxAnimation {
    /** 目标元素索引（slide.elements 中的位置） */
    target: number;
    /** 动画类型 */
    type: 'fade' | 'flyIn' | 'zoom' | 'wipe';
    /** 持续时间（秒） */
    duration: number;
}

/** 图表系列 */
export interface PptxChartSeries {
    name?: string;
    values?: number[];   // 非散点图
    x?: number[];        // 散点图 X
    y?: number[];        // 散点图 Y
    open?: number[];     // 股票图：开盘
    high?: number[];     // 股票图：最高
    low?: number[];      // 股票图：最低
    close?: number[];    // 股票图：收盘
    color?: string;
}

/** 图表分组方式（堆叠/百分比堆叠等） */
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

/** 图表类型（ECMA-376 plotArea 下全部图表节点） */
export type PptxChartType =
    | 'barChart' | 'bar3DChart'
    | 'lineChart' | 'line3DChart'
    | 'areaChart' | 'area3DChart'
    | 'pieChart' | 'pie3DChart' | 'doughnutChart' | 'ofPieChart'
    | 'scatterChart' | 'bubbleChart'
    | 'radarChart' | 'stockChart'
    | 'surfaceChart' | 'surface3DChart'
    | string;

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
    /** 段落级默认样式（纯 text 模式透传给每个段落）：列表/行距/段间距/缩进 */
    bullet?: boolean | 'number' | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    spaceBefore?: number;
    spaceAfter?: number;
    indentLeft?: number;
    indentRight?: number;
    indent?: number;
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
    /** 文本框内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
    /**
     * 文字方向（a:bodyPr/@vert，ST_TextVerticalType）：'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'（默认横排）。
     * 中文竖排用 'eaVert'，逐字堆积用 'wordArtVert'。
     */
    textDirection?: string;
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
    /** 形状特效：阴影 / 发光（对应 a:effectLst） */
    effects?: PptxShapeEffects;
    /**
     * 几何调整值，key 为该预设形状的 OOXML gd 名（roundRect/snip 用 `adj`，
     * 箭头/标注/星形用 `adj1`/`adj2`/…），如 { adj: 25000 }、{ adj1: 50000, adj2: 40000 }
     */
    adjust?: Record<string, number>;
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
    /** 图片裁剪（百分比 0-100）：{ l, r, t, b } */
    crop?: { l?: number; r?: number; t?: number; b?: number };
    /** 图片调整：{ brightness(-100..100), contrast(-100..100), transparency(0..100) } */
    imageAdjust?: { brightness?: number; contrast?: number; transparency?: number };
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
    /** 分组/堆叠方式：stacked、percentStacked 等（bar/line/area 系有效） */
    grouping?: ChartGrouping;
    /** 甜甜圈内径百分比（0-100，默认 50） */
    holeSize?: number;
    /** 折线/散点是否平滑 */
    smooth?: boolean;
    /** 折线/散点是否显示数据标记 */
    marker?: boolean;
    /** 子母饼图（ofPieChart）的第二绘图区类型 */
    ofPieType?: 'pie' | 'bar';
    /** 数值轴/数据标签的数字格式码（如 0.00%、#,##0） */
    numberFormat?: string;
    /** 气泡图：立体显示 */
    bubble3D?: boolean;
    /** 气泡图：显示负气泡 */
    showNegBubbles?: boolean;
    /** 气泡图：气泡缩放百分比（默认 100） */
    bubbleScale?: number;
    /** 曲面图：线框模式 */
    wireframe?: boolean;
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
    /** 四边统一边框 */
    border?: { color?: string; width?: number };
    /** 分边边框（覆盖统一边框） */
    borders?: {
        left?: { color?: string; width?: number } | 'none';
        right?: { color?: string; width?: number } | 'none';
        top?: { color?: string; width?: number } | 'none';
        bottom?: { color?: string; width?: number } | 'none';
    };
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
    /** 单元格内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
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
    /** 表格级默认边框（四边统一） */
    border?: { color?: string; width?: number };
    /** 表格级分边默认边框 */
    borders?: {
        left?: { color?: string; width?: number } | 'none';
        right?: { color?: string; width?: number } | 'none';
        top?: { color?: string; width?: number } | 'none';
        bottom?: { color?: string; width?: number } | 'none';
        /** 对角线边框：'tlBr' / 'blTr' / 'both' */
        diagonal?: 'tlBr' | 'blTr' | 'both';
    };
    /** 表格级单元格内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
    /** 表格样式 ID（引用内置 tableStyles.xml） */
    tableStyleId?: string;
    rows: PptxTableRow[];
}

/** SmartArt 图示节点（层级结构，叶子含 text） */
export interface PptxDiagramNode {
    /** 节点文本 */
    text: string;
    /** 子节点（层级） */
    children?: PptxDiagramNode[];
}

/**
 * SmartArt / 图示元素（解析自 p:graphicFrame/a:graphicData[uri=diagram]，或创作生成）
 *
 * 创作时提供 diagramType + nodes 即可生成原生 diagrams/* 部件；
 * 解析端仅保留可读文本（texts）与原始节点（__raw），还原依赖 __raw 回退。
 */
export interface PptxDiagramElement extends PptxElementBase {
    type: 'diagram';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 图示类型：'list' | 'hierarchy' | 'process' | 'cycle' | 'pyramid'（创作端） */
    diagramType?: string;
    /** 图示节点层级（创作端） */
    nodes?: PptxDiagramNode[];
    /** 图示数据部件（ppt/diagrams/dataN.xml）中的文本内容，按文档顺序（解析端） */
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

/** 分组元素（组合多个子元素，对应 p:grpSp） */
export interface PptxGroupElement extends PptxElementBase {
    type: 'group';
    x: number;
    y: number;
    width: number;
    height: number;
    /**
     * 子元素坐标体系：
     * - 'local'（默认，OOXML 标准）：children 的 x/y 为相对组左上角的局部坐标
     * - 'page'：children 的 x/y 为页绝对坐标，生成时自动减 group 偏移做相对化
     */
    childrenCoordinates?: 'local' | 'page';
    /** 子元素（默认相对组左上角的局部坐标；childrenCoordinates:'page' 时为页绝对坐标） */
    children: PptxElement[];
}

/** 视频元素（mp4 等，p:pic + 媒体关系） */
export interface PptxVideoElement extends PptxElementBase {
    type: 'video';
    x: number;
    y: number;
    width: number;
    height: number;
    data?: string;
    src?: string;
    extension?: string;
    /** 视频海报/预览图 */
    poster?: { data?: string; src?: string; extension?: string };
}

/** 音频元素（mp3/m4a 等，p:pic + 媒体关系） */
export interface PptxAudioElement extends PptxElementBase {
    type: 'audio';
    x: number;
    y: number;
    width: number;
    height: number;
    data?: string;
    src?: string;
    extension?: string;
}

export type PptxElement =
    | PptxTextElement
    | PptxShapeElement
    | PptxImageElement
    | PptxChartElement
    | PptxTableElement
    | PptxDiagramElement
    | PptxGroupElement
    | PptxVideoElement
    | PptxAudioElement
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
    /** 幻灯片批注（生成端写回 ppt/comments/commentsN.xml） */
    comments?: PptxComment[];
    /** 隐藏幻灯片（解析自 p:sld show="0"，生成端写回） */
    hidden?: boolean;
    /** 自动播放：停留毫秒后切换（解析自 p:timing 的 afterTime，生成端写回） */
    advanceTime?: number;
    /** 元素进入动画（解析自 p:timing 的 p:spTgt，生成端写回） */
    animations?: PptxAnimation[];
    elements: PptxElement[];
}

/** 幻灯片批注（commentsN.xml 的 p:cm） */
export interface PptxComment {
    /** 作者（用于 commentAuthors；缺省 'Author'） */
    author?: string;
    /** 批注正文 */
    text: string;
    /** 批注时间（ISO 8601）；缺省取当前时间 */
    dt?: string;
    /** 批注锚点位置（EMU）；缺省 1 英寸处 */
    pos?: { x?: number; y?: number };
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
    /**
     * 文档自定义属性（解析端自 docProps/custom.xml 产出，生成端写回 docProps/custom.xml）
     * 键为属性名，值为字符串（其他类型统一转为字符串）
     */
    customProps?: Record<string, string>;
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
