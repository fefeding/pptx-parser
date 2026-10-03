/**
 * 元素构建器模块
 *
 * 将幻灯片元素（文本/形状/图片）的 JSON 描述转换为 OOXML 节点。
 * 构建过程中通过 ctx 收集图片媒体文件与超链接关系，
 * 供上层（jsonToPptx / editPptx.addSlide）打包 zip 时使用。
 *
 * 支持的元素 JSON 格式：
 * - 文本：{ type:'text', x,y,width,height, text | runs | paragraphs,
 *          align, valign, fontSize(pt), color, bold, italic, underline, fontFace, href }
 * - 形状：{ type:'shape', shapeType:'rect'|'roundRect'|'ellipse'|..., x,y,width,height,
 *          fill:{color}|'none', line:{color,width(pt)}|'none', rotation(deg) }
 * - 图片：{ type:'image', x,y,width,height, data(dataURL/base64)|src(URL), extension, href }
 *
 * @module serializer/element-builders
 */

import { xmlNode, rawXml, pxToEmu, ptToSz, ptToEmu, degToRot, colorToHex, NS, escapeXml, type BuilderNode } from './xml-builder';
import { REL_TYPES, DEFAULT_TABLE_STYLE_ID } from './templates';
import type {
    PptxBackground, PptxTransition, PptxImageSrcRect, PptxImageTile,
    PptxTrendline, Pptx3D, PptxCustomGeometry, PptxAutofit, PptxGeometryPath, PptxGeometryCommand
} from '../types/pptx-document';

/**
 * OOXML 预设几何（ST_PresetGeometryType）白名单。
 * 序列化时若 shapeType 不在该集合内，则回退为 'rect'，避免写出非标准 prst 导致渲染端落入未实现分支。
 */
const PRESET_GEOMETRIES = new Set<string>([
    'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3', 'accentCallout1', 'accentCallout2', 'accentCallout3',
    'actionButtonBackPrevious', 'actionButtonBeginning', 'actionButtonBlank', 'actionButtonDocument', 'actionButtonEnd',
    'actionButtonForwardNext', 'actionButtonHelp', 'actionButtonHome', 'actionButtonInformation', 'actionButtonMovie',
    'actionButtonReturn', 'actionButtonSound', 'arc', 'bentArrow', 'bentUpArrow', 'bevel', 'blockArc', 'bracePair',
    'bracketPair', 'callout1', 'callout2', 'callout3', 'can', 'chartPlus', 'chartStar', 'chartX', 'chevron', 'chord',
    'circularArrow', 'cloud', 'cloudCallout', 'corner', 'cube', 'curvedDownArrow', 'curvedLeftArrow', 'curvedRightArrow',
    'curvedUpArrow', 'decagon', 'diagonalStripe', 'diamond', 'dodecagon', 'donut', 'doubleWave', 'downArrow',
    'downArrowCallout', 'ellipse', 'ellipseRibbon', 'ellipseRibbon2', 'flowChartAlternateProcess', 'flowChartCollate', 'rect',
    'flowChartConnector', 'flowChartDecision', 'flowChartDelay', 'flowChartDisplay', 'flowChartDocument', 'flowChartExtract',
    'flowChartInputOutput', 'flowChartInternalStorage', 'flowChartMagneticDrum', 'flowChartMagneticTape', 'flowChartManualInput',
    'flowChartManualOperation', 'flowChartMerge', 'flowChartMultidocument', 'flowChartOfflineStorage', 'flowChartOnlineStorage',
    'flowChartOr', 'flowChartPredefinedProcess', 'flowChartPreparation', 'flowChartProcess', 'flowChartPunchedCard',
    'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartTerminator', 'folderTab', 'frame', 'funnel', 'gear6', 'gear9',
    'halfFrame', 'heart', 'heptagon', 'hexagon', 'homePlate', 'horizontalScroll', 'irregularSeal1', 'irregularSeal2',
    'leftArrow', 'leftArrowCallout', 'leftBrace', 'leftBracket', 'leftCircularArrow', 'leftRightArrow', 'leftRightArrowCallout',
    'leftRightCircularArrow', 'leftRightUpArrow', 'leftUpArrow', 'lightningBolt', 'line', 'lineInv', 'moon',
    'nonIsoscelesTrapezoid', 'notchedRightArrow', 'octagon', 'parallelogram', 'pentagon', 'pentagonBlock', 'pie', 'pieWedge',
    'plaque', 'plus', 'plusMinus', 'quadArrow', 'quadArrowCallout', 'rectangularCallout', 'ribbon', 'ribbon2', 'rightArrow',
    'rightArrowCallout', 'rightBrace', 'rightBracket', 'round1Rect', 'round2DiagRect', 'round2SameRect', 'roundRect',
    'rtTriangle', 'snip1Rect', 'snip2DiagRect', 'snip2SameRect', 'snipRoundRect', 'sun', 'swooshArrow', 'teardrop', 'trapezoid',
    'triangle', 'upArrow', 'upArrowCallout', 'upDownArrow', 'upDownArrowCallout', 'uturnArrow', 'verticalScroll', 'wave',
    'wedgeEllipseCallout', 'wedgeRectCallout', 'wedgeRoundRectCallout', 'x', 'foldedCorner', 'smileyFace',
    // 连接线（p:cxnSp 专用几何；缺省会被 normalizeShapeType 回退成 rect）
    'straightConnector1',
    'bentConnector2', 'bentConnector3', 'bentConnector4', 'bentConnector5',
    'curvedConnector2', 'curvedConnector3', 'curvedConnector4', 'curvedConnector5'
]);

/** 校验 shapeType 是否为合法 OOXML 预设几何，非法则回退 'rect' */
function normalizeShapeType(t?: string): string {
    if (!t) return 'rect';
    return PRESET_GEOMETRIES.has(t) ? t : 'rect';
}

/** 关系记录（写入 slide rels） */
export interface SerializerRel {
    relId: string;
    type: string;
    target: string;
    external: boolean;
}
/** 媒体文件记录 */
export interface SerializerMedia {
    name: string;
    base64: string;
}
/** 图表嵌入工作簿的行列数据（用于生成 xlsx） */
export interface ChartWorkbookData {
    /** 表头行：系列名称（A1=系列名，B1/C1...=各列名） */
    headers: string[];
    /** 数据行：每行为 [cat, val1, val2, ...] 或 [x, y, size, ...] */
    rows: (string | number)[][];
}
/** 图表部件记录 */
export interface SerializerChart {
    name: string;
    xml: string;
    /** 嵌入工作簿数据（用于生成 WPS 兼容的 xlsx） */
    workbook?: ChartWorkbookData;
}
/** SmartArt 图示部件记录（data/layout/colors/quickStyle 四件套 + drawing，编号共享） */
export interface SerializerDiagram {
    /** 部件编号（与 dataN/layoutN/colorsN/quickStyleN/drawingN 的 N 一致） */
    index: number;
    dataXml: string;
    layoutXml: string;
    colorsXml: string;
    quickStyleXml: string;
    /** 缓存绘图部件 drawingN.xml（Microsoft 标准 dsp:drawing，含按布局算好的形状/填充/文字） */
    drawingXml: string;
}
/** SmartArt 图示节点（层级结构，叶子含 text） */
export interface DiagramNode {
    /** 节点文本 */
    text: string;
    /** 子节点（层级） */
    children?: DiagramNode[];
}
/**
 * __raw 回退所需的附属部件（如 SmartArt 的 diagrams/*.xml）
 * media=true 表示二进制资源（以 base64 落盘，走 Default 扩展名声明）
 */
export interface SerializerPart {
    path: string;
    content?: string;
    base64?: string;
    contentType: string;
    media?: boolean;
}
/** __raw 载荷：原始 OOXML 子树 + 其依赖的关系与部件 */
export interface SerializerRawPayload {
    /** 原始节点标签名（如 p:graphicFrame） */
    tag: string;
    /** tXml simplify 形态的节点内容 */
    node: unknown;
    /** 节点引用的关系：旧 rId → { type, target, external } */
    rels?: Record<string, { type: string; target: string; external?: boolean }>;
    /** 关系指向的部件内容（自包含，保证 round-trip 不丢件） */
    parts?: SerializerPart[];
}
/** 元素构建上下文：构建过程中收集 rels / media / charts / parts */
export interface SerializerContext {
    rels: SerializerRel[];
    media: SerializerMedia[];
    charts: SerializerChart[];
    diagrams: SerializerDiagram[];
    parts: SerializerPart[];
    nextElementId: number;
    nextRelId: number;
    mediaIndex: number;
    chartIndex: number;
    diagramIndex: number;
    /**
     * 本次生成引用到的表格样式 ID：用于在 tableStyles.xml 中补上等价定义，
     * 否则 WPS/PowerPoint 找不到该 GUID 会退化成「无样式无网格」
     */
    tableStyleIds?: Set<string>;
}
/** createElementContext 的选项 */
export interface ElementContextOptions {
    /** 媒体文件起始编号（避免与已有文件冲突） */
    startMediaIndex?: number;
    /** 图表部件起始编号（避免与已有部件冲突） */
    startChartIndex?: number;
    /** 图示部件起始编号（避免与已有部件冲突） */
    startDiagramIndex?: number;
}
/** 运行级样式（文本默认样式 / 单段 run 样式均使用） */
export interface RunStyle {
    align?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    href?: string;
    lang?: string;
    /**
     * 字段类型（a:fld@type）：如 'slidenum'（页码）/ 'datetime'（日期）。
     * 设置后该 run 写出 <a:fld type id><a:rPr/><a:t>text</a:t></a:fld> 而非 <a:r>。
     */
    field?: string;
}
/** 文本运行：{ text, options } 规范格式，或扁平简写格式 */
export interface TextRunSpec extends RunStyle {
    text?: string;
    options?: RunStyle;
}
/** 段落 */
export interface ParagraphSpec {
    text?: string;
    runs?: TextRunSpec[];
    align?: string;
    bullet?: boolean | string | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    spaceBefore?: number;
    spaceAfter?: number;
    indentLeft?: number;
    indentRight?: number;
    indent?: number;
}
/** 单元格/表格边框：颜色 + 线宽(pt) */
export interface CellBorder {
    color?: string;
    width?: number;
}
/** 分边边框：每边为 CellBorder 或 'none'（显式无该边） */
export interface TableBorders {
    left?: CellBorder | 'none';
    right?: CellBorder | 'none';
    top?: CellBorder | 'none';
    bottom?: CellBorder | 'none';
    /** 对角线边框：'tlBr'(左上-右下) / 'blTr'(左下-右上) / 'both' */
    diagonal?: 'tlBr' | 'blTr' | 'both';
}

/** 形状填充/边框/特效（供 SerializerElement 使用） */
export interface ShapeFillSolid { type?: 'solid'; color?: string; transparency?: number; }
export interface ShapeFillGradient { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[]; }
export interface ShapeFillImage {
    type: 'image'; data?: string; src?: string; extension?: string;
    /** 源图裁剪（a:srcRect），各边裁掉的比例，取值 0~1 */
    srcRect?: PptxImageSrcRect;
    /** 平铺（a:tile）：sx/sy 每格占图片的比例，tx/ty 偏移，取值 0~1；不传则拉伸铺满 */
    tile?: PptxImageTile;
}
export interface ShapeFillPattern { type: 'pattern'; prst: string; fg?: string; bg?: string; }
export interface ShapeLineSpec { color?: string; width?: number; transparency?: number; dashType?: string; }
export interface ShapeShadowSpec { type?: 'outer' | 'inner'; color?: string; blur?: number; distance?: number; angle?: number; transparency?: number; }
export interface ShapeGlowSpec { color?: string; blur?: number; }
export interface ShapeEffectsSpec { shadow?: ShapeShadowSpec | boolean; glow?: ShapeGlowSpec | boolean; }

/** 表格单元格 */
export interface SerializerTableCell {
    text?: string;
    paragraphs?: ParagraphSpec[];
    colSpan?: number;
    rowSpan?: number;
    /**
     * 被水平合并吞并（a:tcPr@hMerge="1"）。
     * OOXML 中被合并区域仍需保留单元格节点，否则 PowerPoint 打开会报结构错乱。
     */
    hMerge?: boolean;
    /** 被垂直合并吞并（a:tcPr@vMerge="1"） */
    vMerge?: boolean;
    fill?: string;
    /** 四边统一边框（优先于表格级默认） */
    border?: CellBorder;
    /** 分边边框（最高优先级，覆盖统一边框） */
    borders?: TableBorders;
    align?: string;
    valign?: string;
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
export interface SerializerTableRow {
    height?: number;
    cells: SerializerTableCell[];
}
/** 图表系列 */
export interface ChartSeriesSpec {
    name?: string;
    /** 非散点图数值 */
    values?: number[];
    /** 散点图 x / y 值 */
    x?: number[];
    y?: number[];
    /** 股票图：开盘/最高/最低/收盘值（open/high/low/close） */
    open?: number[];
    high?: number[];
    low?: number[];
    close?: number[];
    color?: string;
    /** 绑定到主/次数值轴（次坐标轴）；需图表级 secondaryValueAxis 配合 */
    axis?: 'primary' | 'secondary';
    /** 该系列是否显示数据标签（覆盖图表级 dataLabels） */
    dataLabels?: boolean;
    /** 该系列趋势线（c:trendline） */
    trendlines?: PptxTrendline[];
}
/** 幻灯片元素 JSON（text / shape / image / chart） */
export interface SerializerElement {
    type?: string;
    name?: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    rotation?: number;
    /** 文本：直接文本 */
    text?: string;
    runs?: TextRunSpec[];
    paragraphs?: ParagraphSpec[];
    align?: string;
    valign?: string;
    /** 段落级默认样式（纯 text 模式透传给每个段落）：列表/行距/段间距/缩进 */
    bullet?: boolean | string | { char?: string; type?: 'number' | 'bullet'; fmt?: string; start?: number };
    lineSpacing?: number | { type: 'pt' | 'percent'; value: number };
    spaceBefore?: number;
    spaceAfter?: number;
    indentLeft?: number;
    indentRight?: number;
    indent?: number;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontFace?: string;
    lang?: string;
    href?: string;
    /** 形状填充：颜色串 / {type:'solid',color,transparency} / {type:'gradient',...} / {type:'image',...} / {type:'pattern',prst} / 'none' / null */
    fill?: string | ShapeFillSolid | ShapeFillGradient | ShapeFillImage | ShapeFillPattern | null;
    /** 形状边框：{ color, width, transparency, dashType } / 'none' / null */
    line?: ShapeLineSpec | 'none' | null;
    /** 形状特效（阴影 / 发光） */
    effects?: ShapeEffectsSpec | null;
    shapeType?: string;
    /** 图片：dataURL / base64 / 远程 URL */
    data?: string;
    src?: string;
    extension?: string;
    /** 视频封面：{ data, src, extension }（仅媒体元素） */
    poster?: { data?: string; src?: string; extension?: string };
    chartType?: string;
    categories?: string[];
    series?: ChartSeriesSpec[];
    varyColors?: boolean;
    barDir?: string;
    title?: string;
    legend?: boolean;
    /** 图表数据标签：true=显示值，或细粒度控制各显示项 */
    dataLabels?: boolean | { showValue?: boolean; showPercent?: boolean; showSeries?: boolean; showCategory?: boolean };
    /** 图表分组/堆叠方式：bar 系默认 clustered，line/area 系默认 standard；支持 stacked / percentStacked */
    grouping?: string;
    /** 甜甜圈内径百分比（0-100，默认 50） */
    holeSize?: number;
    /** 折线/散点平滑线 */
    smooth?: boolean;
    /** 折线/散点数据标记 */
    marker?: boolean;
    /** 子母饼图（ofPieChart）第二绘图区类型，默认 pie */
    ofPieType?: string;
    /** 数值轴与数据标签的数字格式码（如 0.00% / #,##0） */
    numberFormat?: string;
    /** 气泡图：立体显示 */
    bubble3D?: boolean;
    /** 气泡图：显示负气泡 */
    showNegBubbles?: boolean;
    /** 气泡图：气泡缩放百分比（默认 100） */
    bubbleScale?: number;
    /** 曲面图：线框模式 */
    wireframe?: boolean;
    /** 分组（type:'group'）：子元素坐标体系。'local'=相对组左上角的局部坐标（OOXML 标准，默认）；'page'=页绝对坐标（构建时减 group 偏移做相对化） */
    childrenCoordinates?: 'local' | 'page';
    /** 表格：行数据 */
    rows?: SerializerTableRow[];
    /** 表格：列宽（px，缺省均分） */
    colWidths?: number[];
    /** 表格：行高（px，缺省均分） */
    rowHeights?: number[];
    /** 表格级默认边框（四边统一） */
    border?: CellBorder;
    /** 表格级分边默认边框 */
    borders?: TableBorders;
    /** 分组：子元素列表（type:'group' 时） */
    children?: SerializerElement[];
    /** 水平翻转 */
    flipH?: boolean;
    /** 垂直翻转 */
    flipV?: boolean;
    /**
     * 几何调整值，key 必须是该预设形状在 OOXML 中的 gd 名（不是自拟别名），
     * 否则 PowerPoint/WPS 会忽略并回退到默认值。
     * 常见：roundRect/snip 用 `adj`；箭头/标注/星形用 `adj1`/`adj2`/`adj3`…
     * 例：{ adj: 25000 }、{ adj1: 50000, adj2: 40000 }（rightArrow 的箭身厚度与箭头长度）
     */
    adjust?: Record<string, number>;
    /** 文本框内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
    /**
     * 文字方向（a:bodyPr/@vert，ST_TextVerticalType）：'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'（默认横排）。
     * 中文竖排用 'eaVert'，逐字堆积用 'wordArtVert'。非枚举值会被 PowerPoint/WPS 忽略（回退横排），因此这里会做归一/丢弃。
     */
    textDirection?: string;
    /** 图片裁剪（百分比 0-100）：{ l, r, t, b } */
    crop?: { l?: number; r?: number; t?: number; b?: number };
    /** 图片调整：{ brightness(-100..100), contrast(-100..100), transparency(0..100) } */
    imageAdjust?: { brightness?: number; contrast?: number; transparency?: number };
    /** 表格级单元格内边距（px）：{ l, r, t, b }（单元格未显式时继承） */
    tableInset?: { l?: number; r?: number; t?: number; b?: number };
    /** 表格级对角线边框：'tlBr' / 'blTr' / 'both' */
    tableDiagonal?: 'tlBr' | 'blTr' | 'both';
    /** 表格样式 ID（引用内置 tableStyles.xml） */
    tableStyleId?: string;
    /** SmartArt 图示类型：'list' | 'hierarchy' | 'process' | 'cycle' | 'pyramid' */
    diagramType?: string;
    /** SmartArt 图示节点（层级结构，叶子含 text） */
    nodes?: DiagramNode[];
    /** 原始 OOXML 载荷（解析端产出，语义层未覆盖时用于无损回写） */
    __raw?: unknown;
    /** 强制以 __raw 回写（即使 type 已受语义层支持） */
    rawFallback?: boolean;
    /** 替代文本（无障碍，p:cNvPr@descr） */
    descr?: string;
    /** 分栏数（a:bodyPr@numCol，默认 1；需 >1 才写出） */
    numCol?: number;
    /** 栏间距 pt（a:bodyPr@spcCol） */
    spcCol?: number;
    /** 文本框自动适配：'none'(a:noAutofit) / 'normal'(a:normAutofit) / 'shape'(a:spAutoFit) */
    autofit?: PptxAutofit;
    /** 字号缩放百分比（a:normAutofit@fontScale） */
    fontScale?: number;
    /** 行距缩减百分比（a:normAutofit@lnSpcReduction） */
    lnSpcReduction?: number;
    /** 艺术字变形预设（a:bodyPr/a:prstTxWarp@prst），如 'textArchUp' */
    prstTxWarp?: string;
    /** 自定义几何（a:custGeom）；指定时优先于 shapeType */
    custGeom?: PptxCustomGeometry;
    /** 三维属性（a:sp3d + a:scene3d） */
    threeD?: Pptx3D;
    /** 连接线（type:'connector'）起点（px），优先于 x/y 推导 */
    start?: { x: number; y: number };
    /** 连接线（type:'connector'）终点（px） */
    end?: { x: number; y: number };
    /** OLE 程序标识（p:oleObj@progId），如 'Excel.Sheet.12' */
    progId?: string;
    /** OLE 嵌入文件路径或部件名 */
    oleTarget?: string;
    /** OLE 显示为图标（p:oleObj@showAsIcon） */
    showAsIcon?: boolean;
    /** 公式 OMML XML（type:'math'） */
    omml?: string;
    /** 坐标轴标题（c:axTitle） */
    axisTitles?: { category?: string; value?: string; secondaryValue?: string };
    /** 启用次数值轴（配合系列 axis:'secondary'） */
    secondaryValueAxis?: boolean;
    /** 网格线（c:majorGridlines / c:minorGridlines） */
    gridlines?: { major?: boolean; minor?: boolean };
}
/** 幻灯片 JSON */
export interface SerializerSlide {
    /** 使用的版式索引（0 基，按 masters[].layouts 展平后的顺序） */
    layout?: number;
    /** 背景：纯色串 / 渐变 / 图片（与 PptxBackground 同源） */
    background?: string | PptxBackground | null;
    /** 演讲者备注 */
    notes?: string;
    /** 过渡效果 */
    transition?: PptxTransition;
    /** 自动播放：停留毫秒后切换（设置后生成 p:timing 计时） */
    advanceTime?: number;
    /** 是否允许点击切换（默认 true；与 advanceTime 配合） */
    advanceOnClick?: boolean;
    /**
     * 元素动画。type 为 OOXML preset 名（'fade'/'flyIn'/'wipe'/'zoom'/'bounce'…），
     * presetClass 支持 entr(进入) / exit(退出) / emph(强调) / path(路径)。
     */
    animations?: Array<{
        target?: number;
        type?: string;
        duration?: number;
        presetClass?: string;
        presetId?: number;
        presetSubtype?: number;
        delay?: number;
        repeat?: number | 'indefinite';
        direction?: string;
        path?: string;
        trigger?: { type?: 'afterPrev' | 'withPrev' | 'onClick'; target?: number; delay?: number };
    }>;
    /** 隐藏幻灯片（show="0"） */
    hidden?: boolean;
    /** 幻灯片批注（生成 ppt/comments/commentsN.xml） */
    comments?: SerializerComment[];
    elements?: SerializerElement[];
}

/** 幻灯片批注（commentsN.xml 的 p:cm） */
export interface SerializerComment {
    /** 作者（用于 commentAuthors；缺省 'Author'） */
    author?: string;
    /** 批注正文 */
    text: string;
    /** 批注时间（ISO 8601）；缺省取当前时间 */
    dt?: string;
    /** 批注锚点位置（EMU）；缺省 1 英寸处 */
    pos?: { x?: number; y?: number };
}

/**
 * 创建元素构建上下文
 * @param {Object} [options]
 * @param {number} [options.startMediaIndex=0] - 媒体文件起始编号（避免与已有文件冲突）
 * @returns {Object} 构建上下文
 */
export function createElementContext(options: ElementContextOptions = {}): SerializerContext {
    return {
        /** 关系列表 {relId, type, target, external} */
        rels: ([] as SerializerRel[]),
        /** 媒体文件列表 {name, base64} */
        media: ([] as SerializerMedia[]),
        /** __raw 回退附属部件列表 */
        parts: ([] as SerializerPart[]),
        /** 幻灯片内元素自增 id（1 被 spTree 根占用） */
        nextElementId: 2,
        /** 关系自增 id（rId1 固定为版式引用） */
        nextRelId: 2,
        /** 媒体文件自增编号 */
        mediaIndex: options.startMediaIndex || 0,
        /** 图表部件列表 {name, xml}（生成后由上层写入 ppt/charts/） */
        charts: ([] as SerializerChart[]),
        /** 图表自增编号 */
        chartIndex: options.startChartIndex || 0,
        /** 图示部件列表（生成后由上层写入 ppt/diagrams/） */
        diagrams: ([] as SerializerDiagram[]),
        /** 图示自增编号 */
        diagramIndex: options.startDiagramIndex || 0,
        /** 引用到的表格样式 ID（供 tableStyles.xml 补定义） */
        tableStyleIds: new Set<string>()
    };
}

/**
 * 添加关系并返回关系 id
 * @param {Object} ctx - 构建上下文
 * @param {string} type - 关系类型（REL_TYPES 值）
 * @param {string} target - 目标路径
 * @param {boolean} [external=false] - 是否为外部链接
 * @returns {string} 关系 id（如 rId3）
 */
function addRelationship(ctx: SerializerContext, type: string, target: string, external?: boolean) {
    const relId = `rId${ctx.nextRelId++}`;
    ctx.rels.push({ relId, type, target, external: !!external });
    return relId;
}

/**
 * 解析图片来源数据
 * @param {Object} el - 图片元素
 * @returns {Promise<{base64: string, ext: string}>} 图片数据
 */
async function resolveImageData(el: SerializerElement) {
    if (el.data) {
        const str = String(el.data);
        // 兼容任意 mime 的 dataURL（如非常规扩展名产生的 data:application/octet-stream;base64,...）
        const dataUrlMatch = str.match(/^data:([^;,]*);base64,(.+)$/is);
        if (dataUrlMatch) {
            return { base64: dataUrlMatch[2], ext: el.extension || mimeToExt(dataUrlMatch[1]) };
        }
        // 裸 base64
        return { base64: str, ext: el.extension || 'png' };
    }

    if (el.src) {
        if (!/^https?:\/\//i.test(String(el.src))) {
            // 非 URL（如包内相对路径）无法下载，给出明确报错而不是交给 fetch 失败
            throw new Error(`图片元素 src 不是可下载的 URL: ${el.src}。请改用 data（dataURL/base64）或绝对 http(s) 地址`);
        }
        if (typeof fetch !== 'function') {
            throw new Error(`无法获取远程图片 ${el.src}：当前环境不支持 fetch。请将图片下载后以 data(base64/dataURL) 方式提供`);
        }
        const resp = await fetch(el.src);
        if (!resp.ok) {
            throw new Error(`下载远程图片失败: ${el.src} (HTTP ${resp.status})`);
        }
        const buffer = await resp.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);
        const mimeFromHeader = (resp.headers.get('content-type') || '').match(/^image\/([a-z0-9.+-]+)/i);
        let ext = mimeFromHeader ? mimeFromHeader[1].toLowerCase() : null;
        if (ext === 'jpg') ext = 'jpeg';
        if (ext === 'svg+xml') ext = 'svg';
        return { base64, ext: el.extension || ext || guessExtFromUrl(el.src) || 'png' };
    }

    throw new Error("图片元素需要提供 data（dataURL/base64）或 src（远程 URL）字段");
}

/**
 * mime 类型 → 文件扩展名（未知 mime 回退 png）
 * @param {string} mime - mime 类型，如 image/png
 * @returns {string} 扩展名（不含点）
 */
function mimeToExt(mime: string): string {
    const map: Record<string, string> = {
        'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif',
        'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/tiff': 'tiff', 'image/webp': 'webp',
        'image/x-icon': 'ico', 'image/vnd.microsoft.icon': 'ico',
        'image/x-emf': 'emf', 'image/x-wmf': 'wmf'
    };
    const m = String(mime || '').toLowerCase().trim();
    if (map[m]) return map[m];
    const sub = m.split('/')[1];
    if (!sub) return 'png';
    return sub.replace(/\+.*$/, '').replace(/^x-/, '');
}

/**
 * 从 URL 猜测图片扩展名
 * @param {string} url - 图片 URL
 * @returns {string|null} 扩展名
 */
function guessExtFromUrl(url: string) {
    const match = String(url).split('?')[0].match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : null;
}

/**
 * ArrayBuffer 转 base64（兼容浏览器与 Node）
 * @param {ArrayBuffer} buffer - 二进制数据
 * @returns {string} base64 字符串
 */
function arrayBufferToBase64(buffer: ArrayBuffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
        binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)));
    }
    if (typeof btoa === 'function') {
        return btoa(binary);
    }
    return Buffer.from(bytes).toString('base64');
}

/**
 * 构建位置节点（a:xfrm）
 * @param {Object} el - 含 x/y/width/height/rotation 的元素
 * @returns {Object} a:xfrm 节点
 */
function buildXfrm(el: SerializerElement) {
    const xfrmAttrs: Record<string, number | string | null> = { rot: el.rotation ? degToRot(el.rotation) : null };

    // 连接线：用起点/终点推导包围盒，方向由 flipH/flipV 表达（OOXML cxnSp 语义）
    if (el.start && el.end) {
        const sx = el.start.x ?? 0;
        const sy = el.start.y ?? 0;
        const ex = el.end.x ?? 0;
        const ey = el.end.y ?? 0;
        const flipH = ex < sx;
        const flipV = ey < sy;
        if (flipH || el.flipH) xfrmAttrs.flipH = 1;
        if (flipV || el.flipV) xfrmAttrs.flipV = 1;
        return xmlNode('a:xfrm',
            xfrmAttrs,
            xmlNode('a:off', { x: pxToEmu(Math.min(sx, ex)), y: pxToEmu(Math.min(sy, ey)) }),
            xmlNode('a:ext', { cx: pxToEmu(Math.abs(ex - sx)), cy: pxToEmu(Math.abs(ey - sy)) })
        );
    }

    if (el.flipH) xfrmAttrs.flipH = 1;
    if (el.flipV) xfrmAttrs.flipV = 1;
    return xmlNode('a:xfrm',
        xfrmAttrs,
        xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
        xmlNode('a:ext', { cx: pxToEmu(el.width || 0), cy: pxToEmu(el.height || 0) })
    );
}

/**
 * 构建超链接 rPr 子节点
 * @param {Object} ctx - 构建上下文
 * @param {string} href - 链接（http(s):// 外部；'#N' 内部跳转到第 N 页）
 * @returns {Object|null} a:hlinkClick 节点
 */
function buildHyperlink(ctx: SerializerContext, href: string | undefined) {
    if (!href) return null;
    const internalMatch = String(href).match(/^#(\d+)$/);
    if (internalMatch) {
        // 内部跳转：指向对应 slide 部件（slide rels 位于 ppt/slides/_rels/，同目录相对路径）
        const relId = addRelationship(ctx, REL_TYPES.slide, `slide${internalMatch[1]}.xml`);
        return xmlNode('a:hlinkClick', { 'r:id': relId, action: 'ppaction://hlinksldjump' });
    }
    const relId = addRelationship(ctx, REL_TYPES.hyperlink, String(href), true);
    return xmlNode('a:hlinkClick', { 'r:id': relId });
}

/**
 * 构建文本运行（a:r）
 * @param {Object} ctx - 构建上下文
 * @param {string} text - 运行文本
 * @param {Object} opts - 运行选项（fontSize/color/bold/italic/underline/fontFace/href）
 * @returns {Object} a:r 节点
 */
function buildTextRun(ctx: SerializerContext, text: string | undefined, opts: RunStyle = {}) {
    const rPrChildren = [];

    if (opts.color) {
        rPrChildren.push(xmlNode('a:solidFill', colorNode(opts.color)));
    }
    if (opts.fontFace) {
        rPrChildren.push(xmlNode('a:latin', { typeface: opts.fontFace }));
    }
    const hlink = buildHyperlink(ctx, opts.href);
    if (hlink) rPrChildren.push(hlink);

    const rPr = xmlNode('a:rPr',
        {
            lang: opts.lang || 'zh-CN',
            sz: opts.fontSize !== undefined ? ptToSz(opts.fontSize) : null,
            b: opts.bold ? 1 : null,
            i: opts.italic ? 1 : null,
            u: opts.underline ? 'sng' : null,
            dirty: 0
        },
        ...rPrChildren
    );

    // 字段（页码/日期等）：写 a:fld 而非 a:r，保留动态语义（否则会被固化为静态文本）
    if (opts.field) {
        return xmlNode('a:fld',
            { id: `{${generateGuid()}}`, type: opts.field },
            rPr,
            xmlNode('a:t', null, String(text ?? ''))
        );
    }

    return xmlNode('a:r', null, rPr, xmlNode('a:t', null, String(text)));
}

/**
 * 生成 RFC4122 v4 形式 GUID（a:fld@id 要求带花括号的 GUID 字符串）
 */
function generateGuid(): string {
    const hex = '0123456789ABCDEF';
    let out = '';
    // 8-4-4-4-12
    for (let i = 0; i < 36; i++) {
        if (i === 8 || i === 13 || i === 18 || i === 23) {
            out += '-';
        } else if (i === 14) {
            out += '4'; // version 4
        } else if (i === 19) {
            out += hex[(Math.random() * 4 | 0) + 8]; // variant 10xx
        } else {
            out += hex[Math.random() * 16 | 0];
        }
    }
    return out;
}

/**
 * a:buAutoNum/@type 的合法取值（ECMA-376 ST_TextAutonumberScheme）
 * 只列常用项；不在其中的值一律按别名映射或回退到 arabicPeriod，
 * 避免写出 WPS/PowerPoint 无法识别的 type（会被当成默认中文编号渲染）。
 */
const VALID_AUTONUM_TYPES = [
    'alphaLcParenBoth', 'alphaUcParenBoth', 'alphaLcParenR', 'alphaUcParenR', 'alphaLcPeriod', 'alphaUcPeriod',
    'arabicParenBoth', 'arabicParenR', 'arabicPeriod', 'arabicPlain',
    'romanLcParenBoth', 'romanUcParenBoth', 'romanLcParenR', 'romanUcParenR', 'romanLcPeriod', 'romanUcPeriod',
    'circleNumDbPlain', 'circleNumWdWhitePlain', 'circleNumWdBlackPlain',
    'ea1ChsPeriod', 'ea1ChsPlain', 'ea1ChtPeriod', 'ea1ChtPlain', 'ea1JpnChsDbPeriod', 'ea1JpnKorPeriod', 'ea1JpnKorPlain',
    'chineseCounting', 'chineseLegalSimplified', 'chineseCountingThousand', 'ideographDigital',
    'hebrew1', 'hebrew2'
];

/** 友好写法 → 合法 ST_TextAutonumberScheme 值 */
const AUTONUM_TYPE_ALIASES: Record<string, string> = {
    arabic: 'arabicPeriod', decimal: 'arabicPeriod', numeric: 'arabicPeriod', number: 'arabicPeriod', '1': 'arabicPeriod',
    arabicPlain: 'arabicPlain', decimalPlain: 'arabicPlain',
    alphaLc: 'alphaLcPeriod', alphaUc: 'alphaUcPeriod', alpha: 'alphaLcPeriod', letter: 'alphaLcPeriod',
    romanLc: 'romanLcPeriod', romanUc: 'romanUcPeriod', roman: 'romanUcPeriod',
    // 中文编号
    chinese: 'chineseCounting', chineseCounting: 'chineseCounting',
    chineseLegal: 'chineseLegalSimplified',
    ea1Chs: 'ea1ChsPeriod', ea1Cht: 'ea1ChtPeriod', ea1JpnKor: 'ea1JpnKorPeriod'
};

/**
 * 归一化自动编号类型：已经是合法值则原样返回，友好写法按别名表映射，其余回退 arabicPeriod
 * @param {string|undefined} fmt - 用户传入的编号格式
 * @returns {string} 合法的 ST_TextAutonumberScheme 值
 */
function normalizeAutoNumType(fmt: string | undefined): string {
    if (typeof fmt === 'string' && fmt !== '') {
        if (VALID_AUTONUM_TYPES.indexOf(fmt) !== -1) return fmt;
        if (AUTONUM_TYPE_ALIASES[fmt] !== undefined) return AUTONUM_TYPE_ALIASES[fmt];
    }
    return 'arabicPeriod';
}

/**
 * 构建段落（a:p）
 * @param {Object} ctx - 构建上下文
 * @param {Object} paragraph - 段落 { text, runs, align, bullet }
 * @param {Object} defaults - 元素级默认运行选项
 * @returns {Object} a:p 节点
 */
function buildParagraph(ctx: SerializerContext, paragraph: ParagraphSpec, defaults: RunStyle) {
    const alignMap: Record<string, string | null> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
    const p = paragraph || {};

    const pPrAttrs: Record<string, number | string | null> = { algn: alignMap[p.align || defaults.align || 'left'] || null };
    if (p.indentLeft != null) pPrAttrs.marL = ptToEmu(p.indentLeft);
    if (p.indentRight != null) pPrAttrs.marR = ptToEmu(p.indentRight);
    if (p.indent != null) pPrAttrs.indent = ptToEmu(p.indent);

    const pPrChildren: BuilderNode[] = [];

    // 行距 / 段间距
    if (p.lineSpacing != null) {
        const ls = p.lineSpacing;
        if (typeof ls === 'number') {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPct', { val: Math.round(ls * 1000) })));
        } else if (ls.type === 'pt') {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPts', { val: ptToSz(ls.value) })));
        } else {
            pPrChildren.push(xmlNode('a:lnSpc', null, xmlNode('a:spcPct', { val: Math.round(ls.value * 1000) })));
        }
    }
    if (p.spaceBefore != null) pPrChildren.push(xmlNode('a:spcBef', null, xmlNode('a:spcPts', { val: ptToSz(p.spaceBefore) })));
    if (p.spaceAfter != null) pPrChildren.push(xmlNode('a:spcAft', null, xmlNode('a:spcPts', { val: ptToSz(p.spaceAfter) })));

    // 列表符号：自动编号 / 项目符号 / 无
    const b = p.bullet;
    if (b === 'number' || (b && typeof b === 'object' && b.type === 'number')) {
        // fmt 需是合法的 ST_TextAutonumberScheme（'decimal' 这类写法会让 WPS/PowerPoint 回退成中文编号）
        const fmt = normalizeAutoNumType(b && typeof b === 'object' ? b.fmt : undefined);
        const start = (b && typeof b === 'object' && b.start != null) ? b.start : 1;
        pPrChildren.push(xmlNode('a:buAutoNum', { type: fmt, startAt: start }));
    } else if (b === true || b === 'bullet' || (b && typeof b === 'object' && (b.type === 'bullet' || b.char))) {
        const char = (b && typeof b === 'object' && b.char) ? b.char : '•';
        pPrChildren.push(xmlNode('a:buFont', { typeface: 'Arial' }));
        pPrChildren.push(xmlNode('a:buChar', { char }));
    } else {
        pPrChildren.push(xmlNode('a:buNone'));
    }

    const pPr = xmlNode('a:pPr', pPrAttrs, ...pPrChildren);

    // 运行列表：显式 runs 优先，否则用 text + 元素级默认样式
    let runs;
    if (Array.isArray(p.runs) && p.runs.length > 0) {
        runs = p.runs.map((r: TextRunSpec) => {
            // 兼容两种格式：{ text, options: {...} }（规范）与 { text, color, ... }（扁平简写）
            const { text, options, ...flat } = r;
            return buildTextRun(ctx, text, { ...defaults, ...(options || {}), ...flat });
        });
    } else {
        runs = [buildTextRun(ctx, p.text !== undefined ? p.text : '', defaults)];
    }

    return xmlNode('a:p', null, pPr, ...runs);
}

/**
 * 规范化段落列表
 * @param {Object} el - 文本元素
 * @returns {Array<Object>} 段落列表
 */
function normalizeParagraphs(el: SerializerElement): ParagraphSpec[] {
    let paras: ParagraphSpec[];
    if (Array.isArray(el.paragraphs) && el.paragraphs.length > 0) {
        paras = el.paragraphs;
    } else if (Array.isArray(el.runs) && el.runs.length > 0) {
        paras = [{ runs: el.runs }];
    } else if (el.text !== undefined) {
        paras = String(el.text).split('\n').map(t => ({ text: t }));
    } else {
        paras = [{ text: '' }];
    }
    // 元素级段落默认样式（纯 text 模式）：段落自身未显式设置时继承
    const pDefaults: ParagraphSpec = {
        bullet: el.bullet,
        lineSpacing: el.lineSpacing,
        spaceBefore: el.spaceBefore,
        spaceAfter: el.spaceAfter,
        indentLeft: el.indentLeft,
        indentRight: el.indentRight,
        indent: el.indent
    };
    return paras.map(p => ({ ...pDefaults, ...p }));
}

/**
 * 构建文本框元素（p:sp）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 文本元素 JSON
 * @returns {Object} p:sp 节点
 */
/** a:bodyPr/@vert 的合法取值（ST_TextVerticalType） */
const TEXT_VERTICAL_TYPES = new Set([
    'horz', 'vert', 'vert270', 'wordArtVert', 'eaVert', 'mongolianVert', 'wordArtVertRtl'
]);
/** 历史误用值 → 合法枚举值（写成非枚举值会被 PowerPoint/WPS 静默忽略而回退横排） */
const TEXT_VERTICAL_ALIASES: Record<string, string> = {
    wordArtVertical: 'wordArtVert',
    eaVertical: 'eaVert'
};

/**
 * 归一化文字方向：别名转正、非法值丢弃（返回 undefined 时不下发 vert 属性）。
 * @param {string} [value] - 文字方向原始值
 * @returns {string|undefined} 合法枚举值
 */
function normalizeTextVerticalType(value?: string): string | undefined {
    if (!value) return undefined;
    const v = TEXT_VERTICAL_ALIASES[value] ?? value;
    return TEXT_VERTICAL_TYPES.has(v) ? v : undefined;
}

function buildTextElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const anchorMap: Record<string, string | null> = { top: null, middle: 'ctr', bottom: 'b' };
    const bodyPrAttrs: Record<string, number | string | null> = { wrap: 'square', rtlCol: 0 };
    const anchor = anchorMap[el.valign ?? 'top'];
    if (anchor) bodyPrAttrs.anchor = anchor;
    const vert = normalizeTextVerticalType(el.textDirection);
    if (vert) bodyPrAttrs.vert = vert;
    if (el.inset) {
        if (el.inset.l != null) bodyPrAttrs.lIns = pxToEmu(el.inset.l);
        if (el.inset.r != null) bodyPrAttrs.rIns = pxToEmu(el.inset.r);
        if (el.inset.t != null) bodyPrAttrs.tIns = pxToEmu(el.inset.t);
        if (el.inset.b != null) bodyPrAttrs.bIns = pxToEmu(el.inset.b);
    }
    // 分栏（a:bodyPr@numCol / @spcCol）；numCol 需 >1 才有意义，spcCol 单位为 pt
    if (el.numCol && el.numCol > 1) bodyPrAttrs.numCol = el.numCol;
    if (el.spcCol != null) bodyPrAttrs.spcCol = ptToEmu(el.spcCol);

    // a:bodyPr 的子元素顺序（CT_TextBodyProperties）：prstTxWarp → autofit → scene3d → sp3d
    const bodyPrChildren: BuilderNode[] = [];
    if (el.prstTxWarp) {
        bodyPrChildren.push(xmlNode('a:prstTxWarp', { prst: el.prstTxWarp }));
    }
    if (el.autofit) {
        if (el.autofit === 'none') {
            bodyPrChildren.push(xmlNode('a:noAutofit'));
        } else if (el.autofit === 'shape') {
            bodyPrChildren.push(xmlNode('a:spAutoFit'));
        } else {
            bodyPrChildren.push(xmlNode('a:normAutofit', {
                fontScale: el.fontScale != null ? Math.round(el.fontScale * 1000) : null,
                lnSpcReduction: el.lnSpcReduction != null ? Math.round(el.lnSpcReduction * 1000) : null
            }));
        }
    }
    const defaults = {
        align: el.align,
        fontSize: el.fontSize,
        color: el.color,
        bold: el.bold,
        italic: el.italic,
        underline: el.underline,
        fontFace: el.fontFace,
        href: el.href,
        lang: el.lang
    };

    return xmlNode('p:sp',
        null,
        xmlNode('p:nvSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `TextBox ${id - 1}`, descr: el.descr || null }),
            xmlNode('p:cNvSpPr', { txBox: 1 }),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
        ),
        xmlNode('p:txBody',
            null,
            xmlNode('a:bodyPr', bodyPrAttrs, ...bodyPrChildren),
            xmlNode('a:lstStyle'),
            ...normalizeParagraphs(el).map((p: ParagraphSpec) => buildParagraph(ctx, p, defaults))
        )
    );
}

/**
 * 构建形状元素（p:sp）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 形状元素 JSON
 * @returns {Object} p:sp 节点
 */
/** 构造颜色节点：支持 scheme:<name>（a:schemeClr）与绝对色（a:srgbClr） */
function colorNode(color: string): BuilderNode {
    if (typeof color === 'string' && color.startsWith('scheme:')) {
        return xmlNode('a:schemeClr', { val: color.slice(7) });
    }
    return xmlNode('a:srgbClr', { val: colorToHex(color) });
}

/** 0~1 的比例 → OOXML 千分比整数（a:srcRect / a:tile 的 l/t/r/b/sx/sy/tx/ty 均为此单位） */
function ratioThousandth(v: unknown, def: number): number {
    const n = Number(v);
    if (!isFinite(n)) return Math.round(def * 100000);
    return Math.round(Math.min(1, Math.max(0, n)) * 100000);
}

/**
 * 构造 a:blipFill 中除 a:blip 之外的子节点。
 * 顺序遵循 CT_BlipFillProperties：srcRect 在前，其后 stretch 与 tile 二选一。
 * 未指定 tile 时写 a:stretch/a:fillRect（拉伸铺满），与 PowerPoint 默认一致。
 */
function buildBlipFillRects(srcRect?: PptxImageSrcRect | null, tile?: PptxImageTile | null): BuilderNode[] {
    const rects: BuilderNode[] = [];
    if (srcRect) {
        const l = ratioThousandth(srcRect.l, 0);
        const t = ratioThousandth(srcRect.t, 0);
        const r = ratioThousandth(srcRect.r, 0);
        const b = ratioThousandth(srcRect.b, 0);
        if (l !== 0 || t !== 0 || r !== 0 || b !== 0) {
            rects.push(xmlNode('a:srcRect', { l, t, r, b }));
        }
    }
    rects.push(tile
        ? xmlNode('a:tile', {
            sx: ratioThousandth(tile.sx, 1),
            sy: ratioThousandth(tile.sy, 1),
            tx: ratioThousandth(tile.tx, 0),
            ty: ratioThousandth(tile.ty, 0)
        })
        : xmlNode('a:stretch', null, xmlNode('a:fillRect')));
    return rects;
}

/** 构造填充节点（solidFill / gradFill / blipFill / pattFill / noFill）；transparency 写入 a:alpha */
async function buildFillNode(ctx: SerializerContext, fill: SerializerElement['fill']): Promise<BuilderNode | null> {
    if (fill === undefined) return null; // 未指定填充则继承主题
    if (fill === 'none' || fill === null) return xmlNode('a:noFill');
    if (typeof fill === 'string') return xmlNode('a:solidFill', colorNode(fill));
    if ((fill as ShapeFillGradient).type === 'gradient') {
        const g = fill as ShapeFillGradient;
        const stops = (g.stops || []).map((s) =>
            xmlNode('a:gs', { pos: Math.round((s.position || 0) * 100000) }, colorNode(s.color))
        );
        const ang = (g.direction === 'vertical' ? 90 : g.direction === 'diagonal' ? 45 : 0) * 60000;
        return xmlNode('a:gradFill', null, xmlNode('a:gsLst', null, ...stops), xmlNode('a:lin', { ang, scaled: 1 }));
    }
    if ((fill as ShapeFillImage).type === 'image') {
        const img = fill as ShapeFillImage;
        const { base64, ext } = await resolveImageData({ type: 'image', data: img.data, src: img.src, extension: img.extension } as SerializerElement);
        ctx.mediaIndex++;
        const mediaName = `image${ctx.mediaIndex}.${ext}`;
        ctx.media.push({ name: mediaName, base64 });
        const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);
        return xmlNode('a:blipFill', null, xmlNode('a:blip', { 'r:embed': embedRelId }), ...buildBlipFillRects(img.srcRect, img.tile));
    }
    if ((fill as ShapeFillPattern).type === 'pattern') {
        const p = fill as ShapeFillPattern;
        const fg = p.fg ? colorNode(p.fg) : xmlNode('a:schemeClr', { val: 'tx1' });
        const bg = p.bg ? colorNode(p.bg) : xmlNode('a:schemeClr', { val: 'bg1' });
        return xmlNode('a:pattFill', { prst: p.prst }, xmlNode('a:fgClr', null, fg), xmlNode('a:bgClr', null, bg));
    }
    // solid（{color} 或 {type:'solid'}）
    const color = (fill as ShapeFillSolid).color;
    if (!color) return null;
    const alpha = (fill as ShapeFillSolid).transparency;
    // 注意：a:alpha 必须嵌套在颜色节点（srgbClr/schemeClr）内部，不能放在 a:solidFill 下，否则不符合 OOXML schema
    const srgb = colorNode(color);
    if (alpha != null) (srgb.children as BuilderNode[]).push(xmlNode('a:alpha', { val: Math.round((100 - alpha) * 1000) }));
    return xmlNode('a:solidFill', srgb);
}

/**
 * 自定义几何的默认坐标空间（a:path@w/@h）。
 * 用户既可直接给该空间内的坐标，也可给 0~1 归一化坐标（自动放大到本空间）。
 */
const CUSTOM_GEOM_SPACE = 100000;

/**
 * 构建自定义几何（a:custGeom）
 *
 * 位置在 CT_ShapeProperties 中与 a:prstGeom 互斥（二选一）。
 * @param {Object} geom - 自定义几何定义
 * @returns {Object} a:custGeom 节点
 */
function buildCustomGeometry(geom: PptxCustomGeometry): BuilderNode {
    const paths = (geom.paths || []).map((p: PptxGeometryPath) => {
        const w = p.w ?? CUSTOM_GEOM_SPACE;
        const h = p.h ?? CUSTOM_GEOM_SPACE;

        // 归一化坐标（0~1）识别与放大：所有坐标绝对值 ≤1 时按归一化处理
        const rawCoords: number[] = [];
        for (const c of p.commands) {
            if ('x' in c) rawCoords.push(c.x);
            if ('y' in c) rawCoords.push(c.y);
            if ('x1' in c) rawCoords.push(c.x1);
            if ('y1' in c) rawCoords.push(c.y1);
            if ('x2' in c) rawCoords.push(c.x2);
            if ('y2' in c) rawCoords.push(c.y2);
        }
        const isNormalized = rawCoords.length > 0 && rawCoords.every(v => Math.abs(v) <= 1.0001);
        const conv = (v: number) => isNormalized ? Math.round(v * CUSTOM_GEOM_SPACE) : Math.round(v);

        const children: BuilderNode[] = [];
        for (const c of p.commands) {
            switch (c.type) {
                case 'moveTo':
                    children.push(xmlNode('a:moveTo', null, xmlNode('a:pt', { x: conv(c.x), y: conv(c.y) })));
                    break;
                case 'lnTo':
                    children.push(xmlNode('a:lnTo', null, xmlNode('a:pt', { x: conv(c.x), y: conv(c.y) })));
                    break;
                case 'cubicBezTo':
                    children.push(xmlNode('a:cubicBezTo', null,
                        xmlNode('a:pt', { x: conv(c.x1), y: conv(c.y1) }),
                        xmlNode('a:pt', { x: conv(c.x2), y: conv(c.y2) }),
                        xmlNode('a:pt', { x: conv(c.x), y: conv(c.y) })));
                    break;
                case 'quadBezTo':
                    children.push(xmlNode('a:quadBezTo', null,
                        xmlNode('a:pt', { x: conv(c.x1), y: conv(c.y1) }),
                        xmlNode('a:pt', { x: conv(c.x), y: conv(c.y) })));
                    break;
                case 'arcTo':
                    children.push(xmlNode('a:arcTo', { wR: conv(c.wR), hR: conv(c.hR), stAng: c.stAng, swAng: c.swAng }));
                    break;
                case 'close':
                    children.push(xmlNode('a:close'));
                    break;
            }
        }
        // fill:'norm' 表示闭合填充；显式 closed:false 时用 'none'（仅描边）
        return xmlNode('a:path', { w, h, fill: p.closed === false ? 'none' : 'norm' }, ...children);
    });

    return xmlNode('a:custGeom', null,
        xmlNode('a:avLst'),
        xmlNode('a:gdLst'),
        xmlNode('a:ahLst'),
        xmlNode('a:cxnLst'),
        xmlNode('a:rect', { l: 'l', t: 't', r: 'r', b: 'b' }),
        xmlNode('a:pathLst', null, ...paths)
    );
}

/**
 * 构建三维效果节点（a:scene3d + a:sp3d）
 *
 * 必须位于 CT_ShapeProperties 的 effectLst 之后、extLst 之前，否则 PowerPoint 会判为非法。
 * @param {Object} threeD - 三维定义
 * @returns {Array} 节点数组（可能为空）
 */
function build3DNodes(threeD?: Pptx3D): BuilderNode[] {
    if (!threeD) return [];
    const nodes: BuilderNode[] = [];

    const scene = threeD.scene;
    if (scene) {
        const cameraChildren: BuilderNode[] = [];
        if (scene.rotX != null || scene.rotY != null || scene.rotZ != null) {
            // a:rot 的角度单位为 1/60000 度
            cameraChildren.push(xmlNode('a:rot', {
                lat: scene.rotX != null ? Math.round(scene.rotX * 60000) : null,
                lon: scene.rotY != null ? Math.round(scene.rotY * 60000) : null,
                rev: scene.rotZ != null ? Math.round(scene.rotZ * 60000) : null
            }));
        }
        nodes.push(xmlNode('a:scene3d', null,
            xmlNode('a:camera', {
                prst: scene.camera || 'orthographicFront',
                fov: scene.fov != null ? Math.round(scene.fov * 60000) : null,
                zoom: scene.zoom != null ? Math.round(scene.zoom * 1000) : null
            }, ...cameraChildren),
            xmlNode('a:lightRig', {
                rig: scene.lightRig || 'balanced',
                dir: scene.lightDir || 't'
            })
        ));
    }

    const shape = threeD.shape;
    if (shape) {
        const children: BuilderNode[] = [];
        if (shape.bevelTop) {
            children.push(xmlNode('a:bevelT', {
                w: shape.bevelTop.width != null ? ptToEmu(shape.bevelTop.width) : null,
                h: shape.bevelTop.height != null ? ptToEmu(shape.bevelTop.height) : null,
                prst: shape.bevelTop.preset || null
            }));
        }
        if (shape.bevelBottom) {
            children.push(xmlNode('a:bevelB', {
                w: shape.bevelBottom.width != null ? ptToEmu(shape.bevelBottom.width) : null,
                h: shape.bevelBottom.height != null ? ptToEmu(shape.bevelBottom.height) : null,
                prst: shape.bevelBottom.preset || null
            }));
        }
        if (shape.extrusionColor) {
            children.push(xmlNode('a:extrusionClr', null, colorNode(shape.extrusionColor)));
        }
        if (shape.contourColor) {
            children.push(xmlNode('a:contourClr', null, colorNode(shape.contourColor)));
        }
        nodes.push(xmlNode('a:sp3d', {
            // 挤出高度与轮廓线宽均为 ST_PositiveCoordinate（EMU）
            extrusionH: shape.extrusionHeight != null ? ptToEmu(shape.extrusionHeight) : null,
            contourW: shape.contourWidth != null ? ptToEmu(shape.contourWidth) : null,
            prstMaterial: shape.material || null
        }, ...children));
    }

    return nodes;
}

async function buildShapeElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;

    const fillNode = await buildFillNode(ctx, el.fill);

    // 边框
    let lineNode;
    if (el.line === 'none' || el.line === null) {
        lineNode = xmlNode('a:ln', null, xmlNode('a:noFill'));
    } else if (el.line) {
        const w = el.line.width !== undefined ? el.line.width : 1;
        const srgb = colorNode(el.line.color || '#000000');
        if (el.line.transparency != null) (srgb.children as BuilderNode[]).push(xmlNode('a:alpha', { val: Math.round((100 - el.line.transparency) * 1000) }));
        const lnChildren: BuilderNode[] = [xmlNode('a:solidFill', srgb)];
        if (el.line.dashType && el.line.dashType !== 'solid') lnChildren.push(xmlNode('a:prstDash', { val: el.line.dashType }));
        lineNode = xmlNode('a:ln', { w: ptToEmu(w) }, ...lnChildren);
    }

    // 特效（a:effectLst：阴影 / 发光）
    let effectNode: BuilderNode | null = null;
    if (el.effects) {
        const effChildren: BuilderNode[] = [];
        const sh = el.effects.shadow;
        if (sh && sh !== true) {
            const shType = sh.type || 'outer';
            const blur = sh.blur !== undefined ? sh.blur : 4;
            const dist = sh.distance !== undefined ? sh.distance : 3;
            const ang = sh.angle !== undefined ? sh.angle : 90;
            const color = sh.color || '#000000';
            const alpha = sh.transparency != null ? Math.round((100 - sh.transparency) * 1000) : 60000;
            // 注意：OOXML 中阴影的模糊半径属性名是 blurRad（不是 blur），发光的半径属性名是 rad
            effChildren.push(xmlNode(shType === 'inner' ? 'a:innerShdw' : 'a:outerShdw',
                { blurRad: ptToEmu(blur), dist: ptToEmu(dist), dir: Math.round(ang * 60000) },
                xmlNode('a:srgbClr', { val: colorToHex(color) }, xmlNode('a:alpha', { val: alpha }))
            ));
        } else if (sh === true) {
            effChildren.push(xmlNode('a:outerShdw', { blurRad: ptToEmu(4), dist: ptToEmu(3), dir: 5400000 },
                xmlNode('a:srgbClr', { val: '000000' }, xmlNode('a:alpha', { val: 60000 }))
            ));
        }
        const gw = el.effects.glow;
        if (gw && gw !== true) {
            effChildren.push(xmlNode('a:glow', { rad: ptToEmu(gw.blur !== undefined ? gw.blur : 5) },
                xmlNode('a:srgbClr', { val: colorToHex(gw.color || '#FFFF00') })));
        } else if (gw === true) {
            effChildren.push(xmlNode('a:glow', { rad: ptToEmu(5) },
                xmlNode('a:srgbClr', { val: colorToHex('#FFFF00') })));
        }
        if (effChildren.length) effectNode = xmlNode('a:effectLst', null, ...effChildren);
    }

    return xmlNode('p:sp',
        null,
        xmlNode('p:nvSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Shape ${id - 1}`, descr: el.descr || null }),
            xmlNode('p:cNvSpPr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            // 自定义几何与预设几何互斥：有 custGeom 时优先
            el.custGeom
                ? buildCustomGeometry(el.custGeom)
                : xmlNode('a:prstGeom', { prst: normalizeShapeType(el.shapeType) },
                    el.adjust && Object.keys(el.adjust).length
                        ? xmlNode('a:avLst', null, ...Object.entries(el.adjust).map(([name, val]) => xmlNode('a:gd', { name, fmla: `val ${val}` })))
                        : xmlNode('a:avLst')),
            fillNode,
            lineNode,
            ...(effectNode ? [effectNode] : []),
            // scene3d / sp3d 必须排在 effectLst 之后
            ...build3DNodes(el.threeD)
        )
    );
}

/**
 * 构建图片元素（p:pic）
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图片元素 JSON
 * @returns {Promise<Object>} p:pic 节点
 */
async function buildImageElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const { base64, ext } = await resolveImageData(el);

    ctx.mediaIndex++;
    const mediaName = `image${ctx.mediaIndex}.${ext}`;
    ctx.media.push({ name: mediaName, base64 });

    const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);

    // 图片级超链接挂在 cNvPr 上
    let cNvPrChildren = null;
    if (el.href && !/^#\d+$/.test(String(el.href))) {
        const hlinkRelId = addRelationship(ctx, REL_TYPES.hyperlink, String(el.href), true);
        cNvPrChildren = xmlNode('a:hlinkClick', { 'r:id': hlinkRelId });
    }

    // 图片调整 → a:blip 的子节点（CT_Blip 的 EG_EffectProperties 选择组）
    // a:lum 的亮度/对比度是同一个元素的两个属性（bright / contrast），不存在 a:contrast 元素。
    const blipChildren: BuilderNode[] = [];
    let srcRectNode: BuilderNode | null = null;
    if (el.imageAdjust) {
        const adj = el.imageAdjust;
        const lumAttrs: Record<string, number> = {};
        // bright / contrast 均为 ST_FixedPercentage（千分比，-100000..100000），±100% → ±100000
        if (adj.brightness != null) lumAttrs.bright = Math.round(adj.brightness * 1000);
        if (adj.contrast != null) lumAttrs.contrast = Math.round(adj.contrast * 1000);
        if (Object.keys(lumAttrs).length > 0) blipChildren.push(xmlNode('a:lum', lumAttrs));
        // 透明度 → 不透明度：CT_AlphaModulateFixedEffect 的属性是 amt（非 val）
        if (adj.transparency != null) blipChildren.push(xmlNode('a:alphaModFix', { amt: Math.round((100 - adj.transparency) * 1000) }));
    }
    // 裁剪：a:srcRect 是 p:blipFill 的子节点（与 a:blip 同级），不能放进 a:blip 内，
    // 否则 schema 校验失败导致整个 blip 的调整/裁剪被忽略。单位同为千分比。
    if (el.crop) {
        const c = el.crop;
        const srect: Record<string, number> = {};
        if (c.l != null) srect.l = ratioThousandth(c.l, 0);
        if (c.r != null) srect.r = ratioThousandth(c.r, 0);
        if (c.t != null) srect.t = ratioThousandth(c.t, 0);
        if (c.b != null) srect.b = ratioThousandth(c.b, 0);
        if (srect.l || srect.r || srect.t || srect.b) srcRectNode = xmlNode('a:srcRect', srect);
    }

    return xmlNode('p:pic',
        null,
        xmlNode('p:nvPicPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Image ${id - 1}` }, cNvPrChildren),
            xmlNode('p:cNvPicPr', null, xmlNode('a:picLocks', { noChangeAspect: 1 })),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:blipFill',
            null,
            xmlNode('a:blip', { 'r:embed': embedRelId }, ...blipChildren),
            ...(srcRectNode ? [srcRectNode] : []),
            xmlNode('a:stretch', null, xmlNode('a:fillRect'))
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
        )
    );
}

/**
 * 构建图表元素（p:graphicFrame + 原生 c:chartSpace 部件）
 *
 * 输入 el 格式：
 * { type:'chart', x,y,width,height, name,
 *   chartType: 'barChart'|'lineChart'|'areaChart'|'pieChart'|'pie3DChart'|'scatterChart',
 *   title, legend(true), varyColors(true),
 *   categories: ['A','B','C'],
 *   series: [ { name, values:[..], color? }, ... ]            // 非散点
 *   series: [ { name, x:[..], y:[..] } ]                      // 散点
 * }
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图表元素 JSON
 * @returns {Object} p:graphicFrame 节点
 */
function buildChartElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    ctx.chartIndex++;
    const chartNum = ctx.chartIndex;
    const chartName = `chart${chartNum}.xml`;
    const relId = addRelationship(ctx, REL_TYPES.chart, `../charts/${chartName}`);
    const { xml, workbook } = buildChartXml(el);
    ctx.charts.push({ name: chartName, xml, workbook });

    return xmlNode('p:graphicFrame',
        null,
        xmlNode('p:nvGraphicFramePr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Chart ${chartNum}` }),
            xmlNode('p:cNvGraphicFramePr'),
            xmlNode('p:nvPr')
        ),
        buildXfrm(el),
        xmlNode('a:graphic',
            null,
            xmlNode('a:graphicData',
                { uri: NS.c },
                xmlNode('c:chart', { 'xmlns:c': NS.c, 'xmlns:r': NS.r, 'r:id': relId })
            )
        )
    );
}

/** 构造 c:strRef（类别标签缓存） */
function strRefXml(values: unknown[], col: string) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${escapeXml(String(values[i]))}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:strRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:strCache><c:ptCount val="${n}"/>${pts}</c:strCache></c:strRef>`;
}

/** 构造 c:numRef（数值缓存） */
function numRefXml(values: unknown[], col: string) {
    const n = values.length;
    let pts = '';
    for (let i = 0; i < n; i++) {
        pts += `<c:pt idx="${i}"><c:v>${Number(values[i])}</c:v></c:pt>`;
    }
    const last = n > 0 ? n - 1 : 0;
    return `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${2 + last}</c:f>` +
        `<c:numCache><c:fmtCode>General</c:fmtCode><c:ptCount val="${n}"/>${pts}</c:numCache></c:numRef>`;
}

/**
 * 构造趋势线（c:trendline）
 * 子元素顺序遵循 CT_Trendline：name → spPr → trendlineType → order → period
 * → forward → backward → intercept → dispRSqr → dispEq → trendlineLbl
 * @param {Object} t - 趋势线定义
 * @returns {string} c:trendline XML
 */
function buildTrendlineXml(t: PptxTrendline): string {
    const parts: string[] = [];
    if (t.name) parts.push(`<c:name>${escapeXml(t.name)}</c:name>`);
    parts.push(`<c:trendlineType val="${escapeXml(t.type || 'linear')}"/>`);
    if (t.order != null) parts.push(`<c:order val="${Math.round(t.order)}"/>`);
    if (t.period != null) parts.push(`<c:period val="${Math.round(t.period)}"/>`);
    if (t.forward != null) parts.push(`<c:forward val="${Math.round(t.forward)}"/>`);
    if (t.backward != null) parts.push(`<c:backward val="${Math.round(t.backward)}"/>`);
    if (t.intercept != null) parts.push(`<c:intercept val="${t.intercept}"/>`);
    if (t.showRSquared) parts.push('<c:dispRSqr val="1"/>');
    if (t.showEquation) parts.push('<c:dispEq val="1"/>');
    return `<c:trendline>${parts.join('')}</c:trendline>`;
}

/**
 * 生成 c:chartSpace 原生图表 XML（自包含，内联数据缓存，无需外部工作簿）
 * @param {Object} el - 图表元素 JSON
 * @returns {string} chart 部件 XML
 */
function buildChartXml(el: SerializerElement): { xml: string; workbook: ChartWorkbookData } {
    const type = el.chartType || 'barChart';
    // 不能用 /pie/i —— 会把 ofPieChart（子母饼图）误判为普通饼图，导致缺必需的 c:ofPieType
    const isPie = type === 'pieChart' || type === 'pie3DChart';
    const isDoughnut = type === 'doughnutChart';
    const isOfPie = type === 'ofPieChart';
    const isPieLike = isPie || isDoughnut || isOfPie; // 饼类（无坐标轴）
    const isScatter = type === 'scatterChart';
    const isStock = type === 'stockChart';
    const isRadar = type === 'radarChart';
    const isSurface = type === 'surfaceChart' || type === 'surface3DChart';
    const is3D = /3D/i.test(type); // 注意：bar3DChart / pie3DChart 以 3DChart 结尾，不能用 /3D$/
    const isBubble = type === 'bubbleChart';
    const isBarLike = type === 'barChart' || type === 'bar3DChart';
    const isLineArea = type === 'lineChart' || type === 'areaChart' ||
        type === 'line3DChart' || type === 'area3DChart';
    const cats = el.categories || [];
    const series = el.series || [];
    const varyColors = el.varyColors !== undefined
        ? (el.varyColors ? 1 : 0)
        : (isPie || isDoughnut || isOfPie ? 1 : 0);
    // 分组/堆叠：bar 系默认 clustered，line/area 系默认 standard（其余类型无此元素）
    const groupingVal = el.grouping
        ? String(el.grouping)
        : (isBarLike ? 'clustered' : isLineArea ? 'standard' : '');

    /**
     * 构造 c:ser 列表（抽成函数以支持次坐标轴：每对坐标轴对应独立的图表节点，
     * 系列需按 axis 分组，idx/order 在两组间连续编号）
     * @param {Array} serList - 系列子集
     * @param {number} offset - idx/order 起始偏移（保证跨节点唯一）
     * @returns {string} c:ser XML
     */
    const buildSerXml = (serList: ChartSeriesSpec[], offset: number): string => {
        return serList.map((s: ChartSeriesSpec, j: number) => {
            const i = offset + j;
            const tx = `<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f>` +
                `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(s.name || `Series${i + 1}`)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
            let data;
            if (isScatter) {
                data = `<c:xVal>${numRefXml(s.x || [], 'B')}</c:xVal><c:yVal>${numRefXml(s.y || [], 'C')}</c:yVal>`;
            } else if (isBubble) {
                data = `<c:xVal>${numRefXml(s.x || [], 'B')}</c:xVal><c:yVal>${numRefXml(s.y || [], 'C')}</c:yVal><c:bubbleSize>${numRefXml(s.values || [], 'D')}</c:bubbleSize>`;
            } else if (isStock) {
                data = `<c:openVal>${numRefXml(s.open || [], 'B')}</c:openVal>` +
                    `<c:highVal>${numRefXml(s.high || [], 'C')}</c:highVal>` +
                    `<c:lowVal>${numRefXml(s.low || [], 'D')}</c:lowVal>` +
                    `<c:closeVal>${numRefXml(s.close || s.values || [], 'E')}</c:closeVal>`;
            } else {
                data = `<c:cat>${strRefXml(cats, 'A')}</c:cat><c:val>${numRefXml(s.values || [], 'B')}</c:val>`;
            }
            const spPr = s.color ? `<c:spPr><a:solidFill><a:srgbClr val="${colorToHex(s.color)}"/></a:solidFill></c:spPr>` : '';
            // marker 位于数据之前，smooth 位于数据之后（CT_LineSer / CT_ScatterSer 的元素顺序）
            const isSmoothable = isScatter || type === 'lineChart' || type === 'line3DChart';
            const markerXml = (el.marker && isSmoothable) ? '<c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>' : '';
            const smoothXml = (el.smooth && isSmoothable) ? '<c:smooth val="1"/>' : '';
            // 趋势线：CT_*Ser 中位于数据之前（trendline → errBars → cat/val）
            const trendlineXml = (s.trendlines || []).map((t: PptxTrendline) => buildTrendlineXml(t)).join('');
            return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${markerXml}${trendlineXml}${data}${smoothXml}</c:ser>`;
        }).join('');
    };

    // 次坐标轴：同一 plotArea 内需要**两个图表节点**，系列按 axis 归属其一
    const secSeries = el.secondaryValueAxis ? series.filter((s: ChartSeriesSpec) => s.axis === 'secondary') : [];
    const priSeries = el.secondaryValueAxis && secSeries.length
        ? series.filter((s: ChartSeriesSpec) => s.axis !== 'secondary')
        : series;
    // 仅当主/次两组都非空时才拆分（否则退化为单一节点，避免产生空的第二个图表）
    const useSecondary = !!el.secondaryValueAxis && secSeries.length > 0 && priSeries.length > 0;
    const serXml = buildSerXml(useSecondary ? priSeries : series, 0);
    const serXmlSec = useSecondary ? buildSerXml(secSeries, priSeries.length) : '';

    // 数字格式码：同时作用于数据标签与数值轴
    const numFmtXml = el.numberFormat
        ? `<c:numFmt formatCode="${escapeXml(el.numberFormat)}" sourceLinked="0"/>`
        : '';

    // 数据标签必须位于各图表类型节点内部（CT_Chart 本身不含 dLbls 元素）
    const dlObj = (el.dataLabels && typeof el.dataLabels === 'object') ? el.dataLabels : null;
    const dLblsInner = el.dataLabels
        ? `<c:dLbls>${numFmtXml}<c:showVal val="${(el.dataLabels === true || !!dlObj?.showValue) ? 1 : 0}"/>` +
          `<c:showPercent val="${dlObj?.showPercent ? 1 : 0}"/><c:showSer val="${dlObj?.showSeries ? 1 : 0}"/>` +
          `<c:showCatName val="${dlObj?.showCategory ? 1 : 0}"/></c:dLbls>`
        : '';

    // 图表类型特定根（含轴 id；饼类无轴）
    const axIds2 = `<c:axId val="111"/><c:axId val="112"/>`;
    const axIds3 = `<c:axId val="111"/><c:axId val="112"/><c:axId val="113"/>`;
    // 需要序列轴 serAx 的类型：所有 3D 图表 + 曲面图（CT_SurfaceChart 的 serAx 为必需）
    const needsSerAx = is3D || isSurface;
    const axIds = needsSerAx ? axIds3 : axIds2;
    // 各分支均严格按对应 CT_*Chart 的子元素顺序输出，否则 PowerPoint/WPS 会静默忽略
    const makePlot = (serPart: string, axIdsPart: string): string => {
    if (isOfPie) {
        // c:ofPieType 是 CT_OfPieChart 的必需元素，缺失时子母饼图无法识别
        const ofPieType = el.ofPieType === 'bar' ? 'bar' : 'pie';
        return `<c:${type}><c:ofPieType val="${ofPieType}"/><c:varyColors val="${varyColors}"/>${serPart}${dLblsInner}` +
            `<c:gapWidth val="100"/><c:splitType val="auto"/><c:splitPos val="0"/>` +
            `<c:secondPieSize val="75"/><c:serLines/></c:${type}>`;
    } else if (isDoughnut) {
        const holeSize = el.holeSize !== undefined ? el.holeSize : 50;
        return `<c:${type}><c:varyColors val="${varyColors}"/>${serPart}${dLblsInner}<c:holeSize val="${holeSize}"/></c:${type}>`;
    } else if (isPie) {
        return `<c:${type}><c:varyColors val="${varyColors}"/>${serPart}${dLblsInner}</c:${type}>`;
    } else if (isBubble) {
        const bubble3DXml = el.bubble3D ? '<c:bubble3D val="1"/>' : '';
        const bubbleScaleXml = el.bubbleScale !== undefined ? `<c:bubbleScale val="${el.bubbleScale}"/>` : '';
        const showNegXml = el.showNegBubbles ? '<c:showNegBubbles val="1"/>' : '';
        return `<c:${type}>${serPart}${dLblsInner}${bubble3DXml}${bubbleScaleXml}${showNegXml}${axIdsPart}</c:${type}>`;
    } else if (isRadar) {
        return `<c:${type}><c:radarStyle val="standard"/>${serPart}${dLblsInner}${axIdsPart}</c:${type}>`;
    } else if (isStock) {
        // CT_StockChart 无 c:serLines（该元素属于 ofPieChart），高低点连线用 hiLowLines
        return `<c:${type}>${serPart}${dLblsInner}<c:hiLowLines/>${axIdsPart}</c:${type}>`;
    } else if (isSurface) {
        // CT_SurfaceChart 不含 dLbls
        const wireframeXml = el.wireframe ? '<c:wireframe val="1"/>' : '';
        return `<c:${type}>${wireframeXml}${serPart}<c:bandFmts/>${axIdsPart}</c:${type}>`;
    } else {
        const dir = isBarLike ? `<c:barDir val="${el.barDir || 'col'}"/>` : '';
        const grouping = groupingVal ? `<c:grouping val="${groupingVal}"/>` : '';
        const markerXml = (el.marker && isLineArea) ? '<c:marker><c:symbol val="circle"/></c:marker>' : '';
        return `<c:${type}>${dir}${grouping}<c:varyColors val="${varyColors}"/>${serPart}${dLblsInner}${markerXml}${axIdsPart}</c:${type}>`;
    }
    };

    // 主坐标轴对（111/112[+113]）；次坐标轴对固定为 211/212
    const SECONDARY_AX_IDS = `<c:axId val="211"/><c:axId val="212"/>`;
    const plotChart = makePlot(serXml, axIds)
        + (useSecondary ? makePlot(serXmlSec, SECONDARY_AX_IDS) : '');

    // 3D 视图需置于 c:chart 下（ECMA-376：view3D 属于 CT_Chart，位于 plotArea 之前）
    const view3D = is3D ? '<c:view3D><c:rotX val="30"/><c:rotY val="20"/><c:depthPercent val="100"/></c:view3D>' : '';

    // 坐标轴（饼类除外；3D 图表与曲面图还需第三个序列轴 serAx）
    // 散点图/气泡图：X/Y 轴均为值轴（c:valAx），不能用类别轴（ECMA-376 CT_ScatterChart）
    // 所有轴补全 WPS 必需的 c:crosses/c:tickLblPos/c:numFmt 等元素
    const isXYValAx = isScatter || isBubble; // X 轴也是值轴

    // 网格线：默认只画主网格线；显式 gridlines:{ major:false } 可关闭，minor:true 追加次网格线
    const gl = el.gridlines;
    const majorGl = gl && gl.major === false ? '' : '<c:majorGridlines/>';
    const minorGl = gl && gl.minor === true ? '<c:minorGridlines/>' : '';
    // 坐标轴标题（c:axTitle）
    const axisTitleXml = (t?: string) => t
        ? `<c:axTitle><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>${escapeXml(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:axTitle>`
        : '';

    let axes = '';
    if (!isPieLike) {
        const catOrValAxX = isXYValAx
            ? `<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>${majorGl}${minorGl}${axisTitleXml(el.axisTitles?.category)}<c:tickLblPos val="low"/><c:crossAx val="112"/><c:crosses val="autoZero"/></c:valAx>`
            : `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/>${minorGl}${axisTitleXml(el.axisTitles?.category)}<c:tickLblPos val="low"/><c:crossAx val="${needsSerAx ? 113 : 112}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
        const valAxY = `<c:valAx><c:axId val="112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="${escapeXml(el.numberFormat || 'General')}" sourceLinked="0"/>${majorGl}${minorGl}${axisTitleXml(el.axisTitles?.value)}<c:tickLblPos val="low"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
        axes = catOrValAxX + valAxY;
        if (needsSerAx) {
            axes += `<c:serAx><c:axId val="113"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="111"/><c:crosses val="autoZero"/></c:serAx>`;
        }
        // 次坐标轴对（211 分类轴 + 212 数值轴）：分类轴隐藏（delete=1），数值轴置右侧
        if (useSecondary) {
            axes += `<c:catAx><c:axId val="211"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="none"/><c:crossAx val="212"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
            axes += `<c:valAx><c:axId val="212"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="${escapeXml(el.numberFormat || 'General')}" sourceLinked="0"/>${minorGl}${axisTitleXml(el.axisTitles?.secondaryValue)}<c:tickLblPos val="low"/><c:crossAx val="211"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>`;
        }
    }

    const titleXml = el.title
        ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr lang="zh-CN"/><a:t>${escapeXml(el.title)}</a:t></a:r></a:p>` +
          `</c:rich></c:tx><c:overlay val="0"/></c:title>`
        : '';

    const legendXml = el.legend !== false && el.legend !== undefined
        ? `<c:legend><c:legendPos val="${typeof el.legend === 'string' ? el.legend : 'r'}"/><c:overlay val="0"/></c:legend>`
        : '';

    const autoTitleDeleted = `<c:autoTitleDeleted val="${el.title ? 0 : 1}"/>`;

    // 构建嵌入工作簿数据（WPS 要求图表必须关联 xlsx，即使 numCache 已内联数据）
    const workbook = buildChartWorkbookData(type, cats, series);

    return {
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
            `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
            `<c:date1904 val="0"/><c:roundedCorners val="0"/>` +
            `<c:chart>${titleXml}${autoTitleDeleted}${view3D}` +
            `<c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0"/><c:y val="0"/><c:w val="1"/><c:h val="1"/></c:manualLayout></c:layout>${plotChart}${axes}</c:plotArea>` +
            `${legendXml}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`,
        workbook
    };
}

/** 根据图表类型构建嵌入工作簿的行列数据 */
function buildChartWorkbookData(type: string, cats: string[], series: ChartSeriesSpec[]): ChartWorkbookData {
    const isScatter = type === 'scatterChart';
    const isBubble = type === 'bubbleChart';
    const isStock = type === 'stockChart';

    if (isScatter || isBubble) {
        // 散点/气泡：每系列占一组列（X, Y, [Size]），系列名放表头
        const headers: string[] = [];
        const colGroups: number[][][] = []; // 每系列的列数据
        for (const s of series) {
            const cols: number[][] = [];
            cols.push(s.x || []);
            cols.push(s.y || []);
            headers.push(s.name || 'Series');
            headers.push('X');
            headers.push('Y');
            if (isBubble) {
                cols.push(s.values || []);
                headers.push('Size');
            }
            colGroups.push(cols);
        }
        const maxLen = Math.max(...colGroups.flat().map(c => c.length), 0);
        const rows: (string | number)[][] = [];
        for (let r = 0; r < maxLen; r++) {
            const row: (string | number)[] = [];
            for (const cols of colGroups) {
                for (const col of cols) {
                    row.push(col[r] !== undefined ? col[r] : '');
                }
            }
            rows.push(row);
        }
        return { headers, rows };
    }

    if (isStock) {
        // 股票图：Open/High/Low/Close
        const headers = ['Open', 'High', 'Low', 'Close'];
        const s = series[0] || {};
        const open = s.open || [];
        const high = s.high || [];
        const low = s.low || [];
        const close = s.close || s.values || [];
        const maxLen = Math.max(open.length, high.length, low.length, close.length, 0);
        const rows: (string | number)[][] = [];
        for (let r = 0; r < maxLen; r++) {
            rows.push([
                open[r] !== undefined ? open[r] : '',
                high[r] !== undefined ? high[r] : '',
                low[r] !== undefined ? low[r] : '',
                close[r] !== undefined ? close[r] : ''
            ]);
        }
        return { headers, rows };
    }

    // 普通图表：类别在 A 列，每系列占一列
    const headers = ['Category', ...series.map(s => s.name || 'Series')];
    const maxLen = Math.max(cats.length, ...series.map(s => (s.values || []).length), 0);
    const rows: (string | number)[][] = [];
    for (let r = 0; r < maxLen; r++) {
        const row: (string | number)[] = [cats[r] !== undefined ? cats[r] : ''];
        for (const s of series) {
            const vals = s.values || [];
            row.push(vals[r] !== undefined ? vals[r] : '');
        }
        rows.push(row);
    }
    return { headers, rows };
}

/** 表格级边框默认（type:'table' 元素的 border/borders 透传） */
type TableBorderSpec = { border?: CellBorder; borders?: TableBorders };

/**
 * 读取边框对象：
 * - 'none' → 返回 'none' 哨兵（显式无边框，生成 <a:lnX><a:noFill/></a:lnX> 以覆盖表格样式）
 * - 空 → undefined（不生成该边，继承表格样式）
 */
function normalizeCellBorder(b: CellBorder | 'none' | undefined): CellBorder | 'none' | undefined {
    if (b === 'none') return 'none';
    if (!b) return undefined;
    return { color: b.color, width: b.width };
}

/** 解析某条边最终采用的边框：单元格分边 > 单元格统一 > 表格分边 > 表格统一 */
function resolveBorderSide(side: 'L' | 'R' | 'T' | 'B', cell: SerializerTableCell, tableBorder?: TableBorderSpec): CellBorder | 'none' | undefined {
    const key = side === 'L' ? 'left' : side === 'R' ? 'right' : side === 'T' ? 'top' : 'bottom';
    const cSides = cell.borders || {};
    const tSides = (tableBorder && tableBorder.borders) || {};
    if (cSides[key] !== undefined) return normalizeCellBorder(cSides[key]);
    if (cell.border) return { color: cell.border.color, width: cell.border.width };
    if (tSides[key] !== undefined) return normalizeCellBorder(tSides[key]);
    if (tableBorder && tableBorder.border) return { color: tableBorder.border.color, width: tableBorder.border.width };
    return undefined;
}

/** 构造一条单元格边框线 a:ln{side}（含 solidFill 颜色） */
function edgeLineXml(side: 'L' | 'R' | 'T' | 'B', b: CellBorder): BuilderNode {
    const w = b.width !== undefined ? b.width : 1;
    const color = b.color || '#000000';
    return xmlNode(`a:ln${side}`, { w: ptToEmu(w) },
        xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(color) })));
}

/**
 * 构建表格单元格（a:tc）
 * @param {Object} ctx - 构建上下文
 * @param {Object} cell - 单元格 { text | paragraphs, colSpan, rowSpan, fill, border, borders, align, valign, ...样式 }
 * @param {Object} [tableBorder] - 表格级边框默认（type:'table' 的 border/borders）
 * @returns {Object} a:tc 节点
 */
function buildTableCell(ctx: SerializerContext, cell: SerializerTableCell, tableBorder?: TableBorderSpec): BuilderNode {
    const attrs: Record<string, unknown> = {};
    if (cell.colSpan && cell.colSpan > 1) attrs.gridSpan = cell.colSpan;
    if (cell.rowSpan && cell.rowSpan > 1) attrs.rowSpan = cell.rowSpan;
    // 被合并吞并的单元格：仍需输出节点，仅标记 hMerge/vMerge
    if (cell.hMerge) attrs.hMerge = '1';
    if (cell.vMerge) attrs.vMerge = '1';

    const defaults: RunStyle = {
        align: cell.align,
        fontSize: cell.fontSize,
        color: cell.color,
        bold: cell.bold,
        italic: cell.italic,
        underline: cell.underline,
        fontFace: cell.fontFace
    };
    const paragraphs = Array.isArray(cell.paragraphs) && cell.paragraphs.length
        ? cell.paragraphs
        : [{ text: cell.text !== undefined ? cell.text : '' }];

    const tcPrChildren: BuilderNode[] = [];
    // 边框（OOXML 顺序位在填充之前）；显式 'none' 输出 noFill，避免被表格样式网格补上
    for (const side of ['L', 'R', 'T', 'B'] as const) {
        const b = resolveBorderSide(side, cell, tableBorder);
        if (b === 'none') tcPrChildren.push(xmlNode(`a:ln${side}`, null, xmlNode('a:noFill')));
        else if (b) tcPrChildren.push(edgeLineXml(side, b));
    }
    if (cell.fill) {
        tcPrChildren.push(xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(cell.fill) })));
    }
    // 对角线边框（a:lnTlToBr / a:lnBlToTr 本身就是 CT_LineProperties：
    // w/cap/cmpd/algn 与 a:solidFill 直接挂在它上面，不能再嵌一层 a:ln，否则 PowerPoint/WPS 会忽略）
    const diag = cell.borders && cell.borders.diagonal;
    if (diag) {
        const dc = (cell.border && cell.border.color) || '#000000';
        const dw = (cell.border && cell.border.width != null) ? cell.border.width : 1;
        const diagLine = (tag: string) => xmlNode(tag, { w: ptToEmu(dw), cap: 'flat', cmpd: 'sng', algn: 'ctr' },
            xmlNode('a:solidFill', xmlNode('a:srgbClr', { val: colorToHex(dc) })));
        if (diag === 'tlBr' || diag === 'both') tcPrChildren.push(diagLine('a:lnTlToBr'));
        if (diag === 'blTr' || diag === 'both') tcPrChildren.push(diagLine('a:lnBlToTr'));
    }

    // 单元格内边距：OOXML 中是 a:tcPr 的属性（marL/marR/marT/marB），
    // 之前写成自定义元素 <a:tableCellInsets> 属于非法 OOXML，PowerPoint/WPS 会整体忽略
    const anchorMap: Record<string, string | null> = { top: 't', middle: 'ctr', bottom: 'b' };
    const tcPrAttrs: Record<string, unknown> = { anchor: anchorMap[cell.valign ?? 'top'] ?? null };
    if (cell.inset) {
        const ins = cell.inset;
        if (ins.l != null) tcPrAttrs.marL = pxToEmu(ins.l);
        if (ins.r != null) tcPrAttrs.marR = pxToEmu(ins.r);
        if (ins.t != null) tcPrAttrs.marT = pxToEmu(ins.t);
        if (ins.b != null) tcPrAttrs.marB = pxToEmu(ins.b);
    }

    return xmlNode('a:tc', attrs,
        xmlNode('a:txBody', null,
            xmlNode('a:bodyPr', { wrap: 'square', rtlCol: 0 }),
            xmlNode('a:lstStyle'),
            ...paragraphs.map((p) => buildParagraph(ctx, p, defaults))
        ),
        xmlNode('a:tcPr', tcPrAttrs, ...tcPrChildren)
    );
}

/**
 * 构建表格元素（p:graphicFrame + a:graphic/a:graphicData/a:tbl）
 *
 * 输入 el 格式：
 * { type:'table', x,y,width,height, colWidths?: number[], rowHeights?: number[],
 *   rows: [ { height?, cells: [ { text | paragraphs, colSpan, rowSpan, fill, align, valign, ... } ] } ] }
 * 缺省 colWidths/rowHeights 时按整体尺寸均分。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 表格元素 JSON
 * @returns {Object} p:graphicFrame 节点
 */
function buildTableElement(ctx: SerializerContext, el: SerializerElement): BuilderNode {
    const id = ctx.nextElementId++;
    const rows = el.rows || [];
    const width = el.width || 0;
    const height = el.height || 0;

    // 列宽：显式优先，否则均分
    const colCount = el.colWidths && el.colWidths.length
        ? el.colWidths.length
        : Math.max(0, ...rows.map((r) => (r.cells || []).length));
    const colWidths = el.colWidths && el.colWidths.length
        ? el.colWidths
        : new Array(colCount).fill(colCount ? width / colCount : 0);

    // 行高：行级 height > rowHeights > 均分
    const rowHeights = rows.map((r, i) =>
        r.height !== undefined
            ? r.height
            : (el.rowHeights && el.rowHeights[i] !== undefined
                ? el.rowHeights[i]
                : (rows.length ? height / rows.length : 0)));

    const gridCols = colWidths.map((w) => xmlNode('a:gridCol', { w: pxToEmu(w) }));
    // 记录引用到的样式 ID：写 tableStyles.xml 时为其补上等价定义
    const tableStyleId = el.tableStyleId || DEFAULT_TABLE_STYLE_ID;
    if (ctx.tableStyleIds) {
        ctx.tableStyleIds.add(tableStyleId);
    }
    const trNodes = rows.map((row, ri) =>
        xmlNode('a:tr', { h: pxToEmu(rowHeights[ri]) },
            ...(row.cells || []).map((cell) => {
                const merged: SerializerTableCell = { ...cell };
                if (merged.inset == null && el.tableInset != null) merged.inset = el.tableInset;
                const cdiag = merged.borders && merged.borders.diagonal;
                if (el.tableDiagonal && !cdiag) merged.borders = { ...(merged.borders || {}), diagonal: el.tableDiagonal } as TableBorders;
                return buildTableCell(ctx, merged, { border: el.border, borders: el.borders });
            })
        )
    );

    return xmlNode('p:graphicFrame', null,
        xmlNode('p:nvGraphicFramePr', null,
            xmlNode('p:cNvPr', { id, name: el.name || `Table ${id - 1}` }),
            xmlNode('p:cNvGraphicFramePr', null, xmlNode('a:graphicFrameLocks', { noGrp: 1 })),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:xfrm', null,
            xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
            xmlNode('a:ext', { cx: pxToEmu(width), cy: pxToEmu(height) })
        ),
        xmlNode('a:graphic', null,
            xmlNode('a:graphicData', { uri: NS.table },
                xmlNode('a:tbl', null,
                    xmlNode('a:tblPr', { firstRow: 1, bandRow: 1 },
                        // tableStyleId 必须是 a:tblPr 的子元素（写成属性为非法 OOXML，PowerPoint/解析器都无法识别）
                        xmlNode('a:tableStyleId', null, tableStyleId)),
                    xmlNode('a:tblGrid', null, ...gridCols),
                    ...trNodes
                )
            )
        )
    );
}

/** 规整为数组（tXml 单节点为对象、多节点为数组） */
function toArray(v: unknown): unknown[] {
    if (v === undefined || v === null) return [];
    return Array.isArray(v) ? v : [v];
}

/**
 * 将 tXml simplify 形态的原始节点还原为构建器节点（含 rId 重映射）
 * @param {string} tag - 标签名
 * @param {Object|string} node - 原始节点内容
 * @param {Object} remap - 旧 rId → 新 rId 映射
 * @returns {Object} BuilderNode
 */
function rawNodeToBuilder(tag: string, node: unknown, remap: Record<string, string>): BuilderNode {
    if (node === null || node === undefined) return xmlNode(tag, null);
    if (typeof node !== 'object') return xmlNode(tag, null, String(node));

    const src = node as Record<string, unknown>;
    const attrs: Record<string, unknown> = {};
    const srcAttrs = (src.attrs || {}) as Record<string, unknown>;
    for (const k of Object.keys(srcAttrs)) {
        const raw = String(srcAttrs[k]);
        attrs[k] = remap[raw] || raw;
    }

    const children: (BuilderNode | string)[] = [];
    for (const k of Object.keys(src)) {
        if (k === 'attrs') continue;
        for (const child of toArray(src[k])) {
            if (child === null || child === undefined) continue;
            if (typeof child === 'object') children.push(rawNodeToBuilder(k, child, remap));
            else children.push(xmlNode(k, null, String(child)));
        }
    }
    return xmlNode(tag, attrs, ...children);
}

/**
 * 用 __raw 原样回写元素（语义层未覆盖时的无损兜底）
 *
 * 载荷形如 { tag, node, rels, parts }（解析端产出）：
 * - rels 中的旧 rId 会在当前 ctx 中重新登记，节点内的 r:* 属性同步重写为新 rId；
 * - parts 登记到 ctx.parts，由上层（jsonToPptx / addSlide）写入 zip 并声明 Content-Types。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 元素 JSON（需含 __raw）
 * @returns {Object|null} 构建器节点；__raw 缺失或格式不符时返回 null
 */
export function buildRawElement(ctx: SerializerContext, el: SerializerElement): BuilderNode | null {
    const payload = el.__raw as SerializerRawPayload | undefined;
    if (!payload || typeof payload !== 'object' || !payload.tag) return null;

    // 关系重映射：旧 rId → 新 rId
    // 载荷中的 target 为 zip 内绝对路径（ppt/...），写入 slide rels 需转换为 ../ 相对路径
    const remap: Record<string, string> = {};
    for (const [oldRid, rel] of Object.entries(payload.rels || {})) {
        if (!rel || !rel.target) continue;
        const target = rel.target.startsWith('ppt/') ? `../${rel.target.slice(4)}` : rel.target;
        remap[oldRid] = addRelationship(ctx, rel.type, target, rel.external);
    }

    // 附属部件登记
    for (const part of payload.parts || []) {
        if (part && part.path) ctx.parts.push(part);
    }

    return rawNodeToBuilder(payload.tag, payload.node, remap);
}

/**
 * 构建单个幻灯片元素节点
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 元素 JSON
 * @returns {Promise<Object|null>} 元素节点；语义层不支持且无 __raw 时返回 null
 */
/** 构建分组元素（p:grpSp），递归构建子元素 */
async function buildGroupElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const children = el.children || [];
    const childNodes: BuilderNode[] = [];
    // 'page' 模式：子元素坐标是页绝对坐标，构建时减 group 偏移转为组内局部坐标（OOXML 标准）
    const pageCoords = el.childrenCoordinates === 'page';
    const gx = el.x || 0;
    const gy = el.y || 0;
    for (const c of children) {
        const cc = pageCoords ? { ...c, x: (c.x || 0) - gx, y: (c.y || 0) - gy } : c;
        const n = await buildElement(ctx, cc);
        if (n) childNodes.push(n);
    }
    const w = pxToEmu(el.width || 0);
    const h = pxToEmu(el.height || 0);
    return xmlNode('p:grpSp',
        null,
        xmlNode('p:nvGrpSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Group ${id - 1}` }),
            xmlNode('p:cNvGrpSpPr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:grpSpPr',
            null,
            xmlNode('a:xfrm',
                null,
                xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
                xmlNode('a:ext', { cx: w, cy: h }),
                xmlNode('a:chOff', { x: 0, y: 0 }),
                xmlNode('a:chExt', { cx: w, cy: h })
            )
        ),
        // OOXML 要求 group 的子孙必须包在 p:spTree 内，否则 PowerPoint 与解析端均读不到
        xmlNode('p:spTree',
            null,
            xmlNode('p:nvGrpSpPr',
                null,
                xmlNode('p:cNvPr', { id: ctx.nextElementId++, name: `${el.name || 'Group'} Inner` }),
                xmlNode('p:cNvGrpSpPr'),
                xmlNode('p:nvPr')
            ),
            xmlNode('p:grpSpPr',
                null,
                xmlNode('a:xfrm',
                    null,
                    xmlNode('a:off', { x: 0, y: 0 }),
                    xmlNode('a:ext', { cx: w, cy: h }),
                    xmlNode('a:chOff', { x: 0, y: 0 }),
                    xmlNode('a:chExt', { cx: w, cy: h })
                )
            ),
            ...childNodes
        )
    );
}

/** 构建视频/音频媒体元素（p:pic + 媒体关系） */
async function buildMediaElement(ctx: SerializerContext, el: SerializerElement, kind: 'video' | 'audio') {
    const id = ctx.nextElementId++;
    ctx.mediaIndex++;
    const ext = el.extension || (kind === 'video' ? 'mp4' : 'm4a');
    const mediaName = `${kind}${ctx.mediaIndex}.${ext}`;
    const { base64 } = await resolveImageData({ type: 'image', data: el.data, src: el.src, extension: ext } as SerializerElement);
    ctx.media.push({ name: mediaName, base64 });
    const embedRelId = addRelationship(ctx, kind === 'video' ? REL_TYPES.video : REL_TYPES.audio, `../media/${mediaName}`);
    let posterRelId = embedRelId;
    if (el.poster) {
        ctx.mediaIndex++;
        const pExt = el.poster.extension || 'png';
        const pName = `image${ctx.mediaIndex}.${pExt}`;
        const pData = await resolveImageData({ type: 'image', data: el.poster.data, src: el.poster.src, extension: pExt } as SerializerElement);
        ctx.media.push({ name: pName, base64: pData.base64 });
        posterRelId = addRelationship(ctx, REL_TYPES.image, `../media/${pName}`);
    }
    const mediaFileNode = kind === 'video'
        ? xmlNode('a:videoFile', { 'r:link': embedRelId })
        : xmlNode('a:audioFile', { 'r:link': embedRelId });
    return xmlNode('p:pic',
        null,
        xmlNode('p:nvPicPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `${kind} ${id}` }),
            xmlNode('p:cNvPicPr'),
            xmlNode('p:nvPr', null, mediaFileNode)
        ),
        xmlNode('p:blipFill', null, xmlNode('a:blip', { 'r:embed': posterRelId }), xmlNode('a:stretch', null, xmlNode('a:fillRect'))),
        xmlNode('p:spPr', null, xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst')))
    );
}

/** SmartArt 图示命名空间 */
const DGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** 生成确定性的 uniqueId（GUID 形态，便于被 PowerPoint 接受） */
function diagramUniqueId(seed: number): string {
    const hex = (seed & 0xfffffff).toString(16).padStart(8, '0');
    return `{A1B2C3D4-0000-0000-0000-${('00000000' + hex).slice(-12)}}`;
}

/** diagramType → dgm:cat type 缩写 */
function diagramCat(type: string): string {
    const map: Record<string, string> = {
        list: 'list', hierarchy: 'hier', orgChart: 'hier',
        process: 'process', cycle: 'cycle', pyramid: 'pyra', matrix: 'matrix'
    };
    return map[type] || 'list';
}

/** 构建图示数据模型（dataN.xml）：point + connector 树 */
function buildDiagramDataModel(nodes: DiagramNode[]): string {
    let modelId = 0;
    const points = [`<dsdgm:pt modelId="0" type="doc"><dsdgm:extLst/></dsdgm:pt>`];
    const connectors: string[] = [];
    // 同时生成可渲染的 p:drawing/p:spTree（结构级：按层级缩进堆叠，供解析端直接渲染文字）
    const sps: string[] = [];
    let row = 0;
    const EMU = 914400;
    const walk = (ns: DiagramNode[], parentId: number, depth: number) => {
        for (const node of ns) {
            modelId++;
            const mid = modelId;
            const text = escapeXml(node.text || '');
            points.push(
                `<dsdgm:pt modelId="${mid}" type="node">` +
                `<dsdgm:prSet loColr="none" hiColr="none"/>` +
                `<dsdgm:spPr/>` +
                `<dsdgm:t><dsdgm:str val="${text}"/></dsdgm:t>` +
                `<dsdgm:extLst/></dsdgm:pt>`
            );
            const cxnId = 1000 + mid;
            connectors.push(`<dsdgm:cxn modelId="${cxnId}" type="parOf" srcId="${parentId}" destId="${mid}"/>`);
            const x = Math.round(depth * EMU * 1.4);
            const y = Math.round(row * EMU * 0.7);
            row++;
            const fill = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'][depth % 6];
            sps.push(
                `<p:sp><p:nvSpPr><p:cNvPr id="${mid}" name="Node${mid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
                `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2286000" cy="370840"/></a:xfrm>` +
                `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
                `<a:solidFill><a:schemeClr val="${fill}"/></a:solidFill>` +
                `<a:ln><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>` +
                `</p:spPr>` +
                `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
            );
            if (node.children && node.children.length) walk(node.children, mid, depth + 1);
        }
    };
    walk(nodes, 0, 0);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dsdgm:dataModel xmlns:dsdgm="${DGML_NS}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
        `<dsdgm:ptLst>${points.join('')}</dsdgm:ptLst>` +
        `<dsdgm:cxnLst>${connectors.join('')}</dsdgm:cxnLst>` +
        `<p:drawing><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name="Diagram"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        sps.join('') +
        `</p:spTree></p:drawing>` +
        `</dsdgm:dataModel>`;
}

/** 构建图示布局（layoutN.xml）：最小但结构完整的 layoutDef */
function buildDiagramLayout(type: string, seed: number): string {
    const cat = diagramCat(type);
    const uid = diagramUniqueId(seed * 4 + 1);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:layoutDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramStyle" name="${escapeXml(type)}">` +
        `<dgm:title val="${escapeXml(type)}"/><dgm:desc val="pptx-parser generated diagram"/>` +
        `<dgm:catLst><dgm:cat type="${cat}" pri="1000"/></dgm:catLst>` +
        `<dgm:styleData/><dgm:adjLst/><dgm:nodeLst/><dgm:connLst/><dgm:foreachLst/>` +
        `<dgm:layoutNode name="ROOT" styleLbl="node"><dgm:alg type="hierChild"/>` +
        `<dgm:forEach val="__NODE__" refType="begin"><dgm:layoutNode name="NODE" styleLbl="node">` +
        `<dgm:alg type="tx"/><dgm:shp txBox="1" type="rect"/>` +
        `<dgm:forEach val="__CHILD__" refType="child">` +
        `<dgm:layoutNode name="CHILD" styleLbl="node"><dgm:alg type="tx"/><dgm:shp txBox="1" type="rect"/></dgm:layoutNode>` +
        `</dgm:forEach></dgm:layoutNode></dgm:forEach></dgm:layoutNode>` +
        `<dgm:qryLst/><dgm:clrData/><dgm:ruleLst/><dgm:layoutVarLst/><dgm:algLst/>` +
        `</dgm:layoutDef>`;
}

/** 构建图示配色（colorsN.xml）：最小 colorsDef */
function buildDiagramColors(type: string, seed: number): string {
    const uid = diagramUniqueId(seed * 4 + 2);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:colorsDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramColors" name="${escapeXml(type)} colors">` +
        `<dgm:title val="${escapeXml(type)} colors"/><dgm:desc val="generated"/>` +
        `<dgm:catLst/><dgm:varLst/><dgm:styleLblLst/><dgm:sampRagLst/></dgm:colorsDef>`;
}

/** 构建图示快速样式（quickStyleN.xml）：最小 quickStyleDef */
function buildDiagramQuickStyle(type: string, seed: number): string {
    const uid = diagramUniqueId(seed * 4 + 3);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dgm:quickStyleDef xmlns:dgm="${DGML_NS}" uniqueId="${uid}" minVer="2.0.0.0" ` +
        `defStyle="urn:microsoft.com/office/officeart/2005/8/diagramQuickStyle" name="${escapeXml(type)} quick style">` +
        `<dgm:title val="${escapeXml(type)} quick style"/><dgm:desc val="generated"/>` +
        `<dgm:catLst/><dgm:varLst/><dgm:styleLblLst/></dgm:quickStyleDef>`;
}

/** dsp:drawing (缓存绘图) 命名空间 —— 注意是 microsoft 扩展命名空间，解析端会整体 dsp:→p: 替换 */
const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** 单个 SmartArt 节点布局结果（EMU 坐标系，相对 drawing 画布原点） */
interface DiagShape {
    x: number; y: number; w: number; h: number;
    text: string;
    /** 填充主题色（accent1..accent6） */
    fill: string;
    /** 'line' 为连线（细长矩形、无文字无边框），缺省为普通节点 */
    kind?: 'node' | 'line';
}

function countNodes(ns: DiagramNode[]): number {
    let c = 0;
    for (const n of ns) { c++; if (n.children) c += countNodes(n.children); }
    return c;
}
function maxDepth(ns: DiagramNode[], d = 0): number {
    let m = d;
    for (const n of ns) if (n.children) m = Math.max(m, maxDepth(n.children, d + 1));
    return m;
}

/**
 * 按 diagramType 计算各节点的绝对位置/尺寸（EMU）。
 * 这是与 PowerPoint 布局引擎"意图对齐"的简化布局（非像素级），保证：
 * 形状带主题色填充与白色文字、层级/循环/金字塔等结构清晰可见。
 */
function layoutDiagram(type: string, nodes: DiagramNode[], W: number, H: number): DiagShape[] {
    const out: DiagShape[] = [];
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    /** 连线粗细（EMU），用细长矩形充当连接线，避免依赖连接器渲染 */
    const lw = Math.max(1, Math.round(pxToEmu(2)));
    /** 追加一段轴对齐连线（kind='line'）：标准 cxnSp 直线连接器，包围盒退化（一维为 0 表示纯水平/纯垂直） */
    const line = (x: number, y: number, w: number, h: number) => {
        // 允许退化（一维为 0 表示纯水平/垂直直线连接器），两维皆为 0 则跳过
        if (w < 0 || h < 0 || (w === 0 && h === 0)) return;
        out.push({ x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h), text: '', fill: accents[0], kind: 'line' });
    };
    const push = (s: Omit<DiagShape, 'fill'>, depth: number, idx: number) =>
        out.push({ ...s, fill: accents[(type === 'list' || type === 'process') ? idx % 6 : depth % 6] });

    if (type === 'list' || type === 'process') {
        const n = Math.max(1, countNodes(nodes));
        const gap = Math.round(H * 0.04);
        const cellH = Math.round((H - gap * (n - 1)) / n);
        let i = 0;
        const walk = (ns: DiagramNode[]) => {
            for (const node of ns) {
                push({ x: Math.round(W * 0.04), y: i * (cellH + gap), w: Math.round(W * 0.92), h: cellH, text: node.text }, 0, i);
                i++;
                if (node.children) walk(node.children);
            }
        };
        walk(nodes);
        // 相邻节点间的垂直连接线（流程感）
        for (let k = 0; k < n - 1; k++) {
            line(W * 0.5, k * (cellH + gap) + cellH, 0, gap);
        }
    } else if (type === 'hierarchy' || type === 'orgChart') {
        const depthMax = Math.max(1, maxDepth(nodes) + 1);
        const levelH = Math.round(H / depthMax);
        const boxH = Math.round(levelH * 0.72);
        const boxW = Math.round(Math.min(W * 0.26, W / Math.max(2, countNodes(nodes) / depthMax)));
        const links: Array<{ px: number; pyB: number; cx: number; cyT: number }> = [];
        const place = (ns: DiagramNode[], x0: number, x1: number, depth: number) => {
            const span = (x1 - x0) / ns.length;
            ns.forEach((node, idx) => {
                const cx0 = x0 + span * idx, cx1 = x0 + span * (idx + 1);
                const centre = (cx0 + cx1) / 2;
                const x = Math.round(centre - boxW / 2);
                const y = Math.round(depth * levelH + (levelH - boxH) / 2);
                push({ x, y, w: boxW, h: boxH, text: node.text }, depth, 0);
                if (node.children && node.children.length) {
                    const cspan = (cx1 - cx0) / node.children.length;
                    const cyT = Math.round((depth + 1) * levelH + (levelH - boxH) / 2);
                    node.children.forEach((_, ci) => {
                        links.push({ px: centre, pyB: y + boxH, cx: cx0 + cspan * ci + cspan / 2, cyT });
                    });
                    place(node.children, cx0, cx1, depth + 1);
                }
            });
        };
        place(nodes, 0, W, 0);
        // 肘形连线：父底中心 →（竖）层间中线 →（横）子中心 →（竖）子顶中心（对齐 PowerPoint 组织结构图）
        for (const lk of links) {
            const midY = (lk.pyB + lk.cyT) / 2;
            line(lk.px, lk.pyB, 0, midY - lk.pyB);
            if (Math.abs(lk.cx - lk.px) > lw) {
                const x1 = Math.min(lk.px, lk.cx), x2 = Math.max(lk.px, lk.cx);
                line(x1, midY, x2 - x1, 0);
            }
            line(lk.cx, midY, 0, lk.cyT - midY);
        }
    } else if (type === 'cycle') {
        const n = Math.max(1, countNodes(nodes));
        const boxW = Math.round(Math.min(W * 0.22, H * 0.22));
        const boxH = boxW;
        const cx = W / 2, cy = H / 2;
        const rx = W / 2 - boxW / 2 - Math.round(W * 0.05);
        const ry = H / 2 - boxH / 2 - Math.round(H * 0.05);
        let i = 0;
        const walk = (ns: DiagramNode[]) => {
            for (const node of ns) {
                const ang = (2 * Math.PI * i) / n - Math.PI / 2;
                push({ x: Math.round(cx + rx * Math.cos(ang) - boxW / 2), y: Math.round(cy + ry * Math.sin(ang) - boxH / 2), w: boxW, h: boxH, text: node.text }, 0, i);
                i++;
                if (node.children) walk(node.children);
            }
        };
        walk(nodes);
    } else if (type === 'pyramid') {
        const n = Math.max(1, countNodes(nodes));
        const cellH = Math.round(H / n);
        let i = 0;
        const walk = (ns: DiagramNode[]) => {
            for (const node of ns) {
                const t = (i + 0.5) / n;
                const w = Math.round(W * (0.34 + 0.62 * t));
                push({ x: Math.round((W - w) / 2), y: Math.round(i * cellH), w, h: cellH - Math.round(H * 0.02), text: node.text }, 0, i);
                i++;
                if (node.children) walk(node.children);
            }
        };
        walk(nodes);
    } else {
        return layoutDiagram('list', nodes, W, H);
    }
    return out;
}

/**
 * 构建缓存绘图部件 drawingN.xml（dsp:drawing 根）。
 * 写入按 layoutDiagram 算好的每个节点 dsp:sp（a:xfrm 绝对坐标 + a:solidFill 主题色 + 白色文字）。
 * 这是与 PowerPoint 对齐的关键：PowerPoint 打开即呈现完整图示，解析端走 diagramDrawing 回退路径渲染同款。
 * @param type diagramType
 * @param seed 部件编号（= diagramIndex）
 * @param nodes 节点树
 * @param W drawing 画布宽（EMU，= graphicFrame 宽）
 * @param H drawing 画布高（EMU，= graphicFrame 高）
 */
function buildDiagramDrawing(type: string, seed: number, nodes: DiagramNode[], W: number, H: number): string {
    const shapes = layoutDiagram(type, nodes, W, H);
    let spid = 2;
    const sps = shapes.map((s, idx) => {
        const id = spid++;
        const modelId = diagramUniqueId(seed * 100 + idx + 1);
        const xfrm = `<a:xfrm><a:off x="${s.x}" y="${s.y}"/><a:ext cx="${s.w}" cy="${s.h}"/></a:xfrm>`;
        if (s.kind === 'line') {
            // 连线：标准 dsp:cxnSp（解析端重写为 p:cxnSp）直线连接器，a:prstGeom=straightConnector1，
            // 包围盒退化（一维为 0 表示纯水平/垂直），线宽与颜色由 a:ln 表达（无填充、无文字）。
            const lwEmu = Math.max(9525, Math.round(pxToEmu(2)));
            return `<dsp:cxnSp modelId="${escapeXml(modelId)}"><dsp:nvCxnSpPr><dsp:cNvPr id="${id}" name="Connector ${id}"/><dsp:cNvCxnSpPr><a:stCxn id="0" idx="0"/><a:endCxn id="0" idx="0"/></dsp:cNvCxnSpPr><dsp:nvPr/></dsp:nvCxnSpPr>` +
                `<dsp:spPr bwMode="auto">${xfrm}` +
                `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>` +
                `<a:ln w="${lwEmu}"><a:solidFill><a:schemeClr val="${s.fill}"/></a:solidFill></a:ln>` +
                `</dsp:spPr>` +
                `</dsp:cxnSp>`;
        }
        const sz = Math.max(900, Math.min(2200, Math.round((s.h / 914400) * 1200))); // 字号随框高自适应（EMU→pt*100）
        return `<dsp:sp modelId="${escapeXml(modelId)}"><dsp:nvSpPr><dsp:cNvPr id="${id}" name="Node ${id}"/><dsp:cNvSpPr/></dsp:nvSpPr>` +
            `<dsp:spPr bwMode="auto">${xfrm}` +
            `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
            `<a:solidFill><a:schemeClr val="${s.fill}"/></a:solidFill>` +
            `<a:ln><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>` +
            `</dsp:spPr>` +
            `<dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>` +
            `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="${sz}" b="1"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:rPr>` +
            `<a:t>${escapeXml(s.text)}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}" xmlns:r="${NS.r}">` +
        `<dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="1" name="Diagram"/><dsp:cNvGrpSpPr/><dsp:nvPr/></dsp:nvGrpSpPr>` +
        `<dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></dsp:grpSpPr>` +
        sps +
        `</dsp:spTree></dsp:drawing>`;
}

/**
 * 构建 SmartArt 图示元素（p:graphicFrame + 原生 diagrams/* 部件）
 * 生成 data/layout/colors/quickStyle 四件套，并登记幻灯片 → dataN.xml 关系；
 * 四部件由上层（jsonToPptx / editPptx.addSlide）写入 ppt/diagrams/ 并补 dataN.xml.rels。
 *
 * 限制：生成的 layoutDef/colorsDef/quickStyleDef 为结构级最小骨架，可被本解析器 round-trip、
 * 可被 PowerPoint 打开，但 PowerPoint 对 SmartArt 布局引擎校验严格，自定义布局可能不渲染为
 * 完整图示图形。如需保真渲染，应提供完整的标准 layoutDef 模板替换 buildDiagramLayout 等。
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 图示元素 JSON { type:'diagram', diagramType, nodes:[{text,children?}] }
 * @returns {Promise<Object>} p:graphicFrame 节点
 */
async function buildDiagramElement(ctx: SerializerContext, el: SerializerElement) {
    ctx.diagramIndex++;
    const n = ctx.diagramIndex;
    const dgmType = el.diagramType || 'list';
    const nodes: DiagramNode[] = el.nodes || [];
    const dataXml = buildDiagramDataModel(nodes);
    const layoutXml = buildDiagramLayout(dgmType, n);
    const colorsXml = buildDiagramColors(dgmType, n);
    const quickStyleXml = buildDiagramQuickStyle(dgmType, n);
    // 缓存绘图部件（Microsoft 标准 dsp:drawing）：按布局算好的形状/填充/文字，解析端与 PowerPoint 对齐
    const drawW = pxToEmu(el.width || 400);
    const drawH = pxToEmu(el.height || 300);
    const drawingXml = buildDiagramDrawing(dgmType, n, nodes, drawW, drawH);
    // 标准 SmartArt：slide 通过 dgm:relIds 引用 data/layout/colors/quickStyle；另挂 diagramDrawing 关系
    const dataRelId = addRelationship(ctx, REL_TYPES.diagramData, `../diagrams/data${n}.xml`);
    const colorsRelId = addRelationship(ctx, REL_TYPES.diagramColors, `../diagrams/colors${n}.xml`);
    const layoutRelId = addRelationship(ctx, REL_TYPES.diagramLayout, `../diagrams/layout${n}.xml`);
    const quickStyleRelId = addRelationship(ctx, REL_TYPES.diagramQuickStyle, `../diagrams/quickStyle${n}.xml`);
    const drawingRelId = addRelationship(ctx, REL_TYPES.diagramDrawing, `../diagrams/drawing${n}.xml`);
    ctx.diagrams.push({ index: n, dataXml, layoutXml, colorsXml, quickStyleXml, drawingXml });

    const id = ctx.nextElementId++;
    return xmlNode('p:graphicFrame',
        null,
        xmlNode('p:nvGraphicFramePr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Diagram ${id}` }),
            xmlNode('p:cNvGraphicFramePr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:xfrm',
            null,
            xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
            xmlNode('a:ext', { cx: pxToEmu(el.width || 400), cy: pxToEmu(el.height || 300) })
        ),
        xmlNode('a:graphic',
            null,
            xmlNode('a:graphicData', { uri: DGML_NS },
                xmlNode('dgm:relIds', {
                    'xmlns:dgm': DGML_NS,
                    'xmlns:r': NS.r,
                    'r:cs': colorsRelId,
                    'r:dm': dataRelId,
                    'r:lo': layoutRelId,
                    'r:qs': quickStyleRelId
                }))
        )
    );
}

/**
 * 构建连接线元素（p:cxnSp）
 *
 * 与 p:sp 的区别：外层是 p:nvCxnSpPr / p:cNvCxnSpPr，且 a:xfrm 语义为起止两点
 * （若提供 start/end 则由 buildXfrm 推导包围盒 + flipH/flipV 表达方向）。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 连接线元素 JSON
 * @returns {Object} p:cxnSp 节点
 */
function buildConnectorElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const line = (el.line && el.line !== 'none') ? el.line as { color?: string; width?: number; dashType?: string } : null;

    const lnChildren: BuilderNode[] = [];
    if (line && line.color) lnChildren.push(xmlNode('a:solidFill', colorNode(line.color)));
    if (line && line.dashType) lnChildren.push(xmlNode('a:prstDash', { val: line.dashType }));

    const lnNode = el.line === 'none'
        ? xmlNode('a:ln', null, xmlNode('a:noFill'))
        : xmlNode('a:ln', {
            w: line && line.width != null ? ptToEmu(line.width) : null,
            cap: 'flat'
        }, ...lnChildren);

    return xmlNode('p:cxnSp',
        null,
        xmlNode('p:nvCxnSpPr',
            null,
            xmlNode('p:cNvPr', { id, name: el.name || `Connector ${id - 1}`, descr: el.descr || null }),
            xmlNode('p:cNvCxnSpPr'),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr',
            null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: normalizeShapeType(el.shapeType || 'straightConnector1') },
                el.adjust && Object.keys(el.adjust).length
                    ? xmlNode('a:avLst', null, ...Object.entries(el.adjust).map(([name, val]) => xmlNode('a:gd', { name, fmla: `val ${val}` })))
                    : xmlNode('a:avLst')),
            lnNode,
            ...build3DNodes(el.threeD)
        )
    );
}

/**
 * 构建 OLE 嵌入对象元素（p:graphicFrame + a:graphicData[uri=presentationml/ole] + p:oleObj）
 *
 * 结构要点：p:oleObj 必须内嵌一个 p:pic 作为**显示代理图**，否则 PowerPoint 打开时
 * 对象区域会显示为空白（对象本体只在双击激活时由 progId 对应的程序渲染）。
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - OLE 元素 JSON
 * @returns {Object} p:graphicFrame 节点
 */
async function buildOleElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const ext = (el.extension || 'bin').replace(/^\./, '');

    // 嵌入部件落到 ppt/embeddings/（与媒体 ppt/media 区分）
    let oleRelId: string | null = null;
    if (el.data) {
        ctx.parts.push({
            path: `ppt/embeddings/oleObject${ctx.parts.length + 1}.${ext}`,
            base64: el.data,
            contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
            media: true
        });
        oleRelId = addRelationship(ctx, REL_TYPES.oleObject, `../embeddings/oleObject${ctx.parts.length}.${ext}`);
    } else if (el.oleTarget) {
        oleRelId = addRelationship(ctx, REL_TYPES.oleObject, el.oleTarget);
    }

    // 显示代理图（poster）
    let posterRelId: string | null = null;
    if (el.poster) {
        const { base64, ext: pExt } = await resolveImageData({ type: 'image', data: el.poster.data, src: el.poster.src, extension: el.poster.extension } as SerializerElement);
        ctx.mediaIndex++;
        const pName = `image${ctx.mediaIndex}.${pExt}`;
        ctx.media.push({ name: pName, base64 });
        posterRelId = addRelationship(ctx, REL_TYPES.image, `../media/${pName}`);
    }

    const cx = pxToEmu(el.width || 200);
    const cy = pxToEmu(el.height || 150);

    const picChildren: BuilderNode[] = [
        xmlNode('p:nvPicPr', null,
            xmlNode('p:cNvPr', { id: ctx.nextElementId++, name: `${el.name || 'Object'} Display` }),
            xmlNode('p:cNvPicPr', null, xmlNode('a:picLocks', { noGrp: 1, noChangeAspect: 1 })),
            xmlNode('p:nvPr')
        )
    ];
    if (posterRelId) {
        picChildren.push(xmlNode('p:blipFill', null,
            xmlNode('a:blip', { 'r:embed': posterRelId }),
            xmlNode('a:stretch', null, xmlNode('a:fillRect'))
        ));
    }
    picChildren.push(xmlNode('p:spPr', null,
        xmlNode('a:xfrm', null,
            xmlNode('a:off', { x: 0, y: 0 }),
            xmlNode('a:ext', { cx, cy })
        ),
        xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
    ));

    return xmlNode('p:graphicFrame',
        null,
        xmlNode('p:nvGraphicFramePr', null,
            xmlNode('p:cNvPr', { id, name: el.name || `Object ${id - 1}`, descr: el.descr || null }),
            xmlNode('p:cNvGraphicFramePr', null, xmlNode('a:graphicFrameLocks', { noGrp: 1 })),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:xfrm', null,
            xmlNode('a:off', { x: pxToEmu(el.x || 0), y: pxToEmu(el.y || 0) }),
            xmlNode('a:ext', { cx, cy })
        ),
        xmlNode('a:graphic', null,
            xmlNode('a:graphicData', { uri: 'http://schemas.openxmlformats.org/presentationml/2006/ole' },
                xmlNode('p:oleObj', {
                    progId: el.progId || 'Package',
                    'r:id': oleRelId,
                    showAsIcon: el.showAsIcon ? 1 : 0
                },
                ...picChildren
                )
            )
        )
    );
}

/**
 * 构建公式元素（OMML）
 *
 * PresentationML 中 OMML 通过 mc:AlternateContent 挂载：
 * - mc:Choice Requires="a14" 内放 a14:m（OMML 本体，需 m: 命名空间）
 * - mc:Fallback 放纯文本，保证不支持 OMML 的查看器仍能显示
 *
 * @param {Object} ctx - 构建上下文
 * @param {Object} el - 公式元素 JSON
 * @returns {Object} p:sp 节点
 */
function buildMathElement(ctx: SerializerContext, el: SerializerElement) {
    const id = ctx.nextElementId++;
    const text = el.text ?? '';

    const altContent = xmlNode('mc:AlternateContent',
        {
            'xmlns:mc': NS.mc,
            'xmlns:a14': NS.a14,
            'xmlns:m': NS.m
        },
        xmlNode('mc:Choice', { Requires: 'a14' },
            xmlNode('a14:m', null, el.omml ? rawXml(el.omml) : null)
        ),
        xmlNode('mc:Fallback', null,
            xmlNode('a:t', null, text)
        )
    );

    return xmlNode('p:sp',
        null,
        xmlNode('p:nvSpPr', null,
            xmlNode('p:cNvPr', { id, name: el.name || `Math ${id - 1}`, descr: el.descr || null }),
            xmlNode('p:cNvSpPr', { txBox: 1 }),
            xmlNode('p:nvPr')
        ),
        xmlNode('p:spPr', null,
            buildXfrm(el),
            xmlNode('a:prstGeom', { prst: 'rect' }, xmlNode('a:avLst'))
        ),
        xmlNode('p:txBody', null,
            xmlNode('a:bodyPr', { wrap: 'square', rtlCol: 0 }),
            xmlNode('a:lstStyle'),
            xmlNode('a:p', null,
                el.omml ? altContent : xmlNode('a:r', null,
                    xmlNode('a:rPr', { lang: 'zh-CN', dirty: 0 }),
                    xmlNode('a:t', null, text)
                )
            )
        )
    );
}

export async function buildElement(ctx: SerializerContext, el: SerializerElement) {
    if (!el || typeof el !== 'object') return null;
    // 显式回退：已支持的语义类型也可用 __raw 原样回写（语义层可能丢失主题色/动画等细节）
    if (el.rawFallback && el.__raw) return buildRawElement(ctx, el);
    // 形状元素：type 为具体几何名（如 'rect'/'ellipse'）时统一转入 buildShapeElement；
    // 同时保留 type:'shape'（几何写在 shapeType 字段）的旧用法。
    if (el.type && PRESET_GEOMETRIES.has(el.type)) {
        return buildShapeElement(ctx, { ...el, shapeType: el.type });
    }
    switch (el.type) {
        case 'text':
            return buildTextElement(ctx, el);
        case 'shape':
            return buildShapeElement(ctx, el);
        case 'image':
            return buildImageElement(ctx, el);
        case 'video':
        case 'audio':
            return buildMediaElement(ctx, el, el.type as 'video' | 'audio');
        case 'chart':
            return buildChartElement(ctx, el);
        case 'table':
            return buildTableElement(ctx, el);
        case 'group':
            return buildGroupElement(ctx, el);
        case 'connector':
            return buildConnectorElement(ctx, el);
        case 'ole':
            return buildOleElement(ctx, el);
        case 'math':
            return buildMathElement(ctx, el);
        case 'diagram':
            // 有 __raw 载荷（解析端回读）时回落无损回退；否则按语义层生成原生 diagrams 部件
            return el.__raw ? buildRawElement(ctx, el) : buildDiagramElement(ctx, el);
        default:
            // 语义层未覆盖 → 回退 __raw
            return buildRawElement(ctx, el);
    }
}

/**
 * 构建背景节点（p:bg）
 * @param {Object} bg - 背景描述（字符串 / {type:'solid'} / {type:'gradient'} / {type:'image'}）
 * @param {Object} ctx - 构建上下文（图片型背景需登记媒体与关系）
 * @returns {Promise<Object|null>} p:bg 节点，无背景返回 null
 */
export async function buildBackground(bg: string | PptxBackground | null | undefined, ctx: SerializerContext): Promise<BuilderNode | null> {
    if (!bg || bg === 'none') return null;

    let bgPrChildren: (BuilderNode | string)[];
    if (typeof bg === 'string') {
        bgPrChildren = [
            xmlNode('a:solidFill', colorNode(bg)),
            xmlNode('a:effectLst')
        ];
    } else if (bg.type === 'solid') {
        bgPrChildren = [
            xmlNode('a:solidFill', colorNode(bg.color)),
            xmlNode('a:effectLst')
        ];
    } else if (bg.type === 'gradient') {
        const stops = (bg.stops || []).map(s =>
            xmlNode('a:gs', { pos: Math.round((s.position || 0) * 100000) },
                xmlNode('a:srgbClr', { val: colorToHex(s.color) }))
        );
        const ang = bg.direction === 'vertical' ? 90 : bg.direction === 'diagonal' ? 45 : 0;
        bgPrChildren = [
            xmlNode('a:gradFill', null,
                xmlNode('a:gsLst', null, ...stops),
                xmlNode('a:lin', { ang, scaled: 1 })
            )
        ];
    } else {
        // 图片型背景：内联媒体并登记关系
        const imgEl: SerializerElement = { type: 'image', data: bg.data, src: bg.src, extension: bg.extension };
        const { base64, ext } = await resolveImageData(imgEl);
        ctx.mediaIndex++;
        const mediaName = `image${ctx.mediaIndex}.${ext}`;
        ctx.media.push({ name: mediaName, base64 });
        const embedRelId = addRelationship(ctx, REL_TYPES.image, `../media/${mediaName}`);
        bgPrChildren = [
            xmlNode('a:blipFill', null,
                xmlNode('a:blip', { 'r:embed': embedRelId }),
                ...buildBlipFillRects(bg.srcRect, bg.tile)
            )
        ];
    }

    return xmlNode('p:bg', null, xmlNode('p:bgPr', null, ...bgPrChildren));
}

/** 过渡类型 → OOXML p:* 子元素 */
const TRANSITION_TAG: Record<string, string> = {
    fade: 'p:fade', wipe: 'p:wipe', push: 'p:push', cover: 'p:cover',
    blinds: 'p:blinds', checker: 'p:checker', circle: 'p:circle', comb: 'p:comb',
    dissolve: 'p:dissolve', random: 'p:random', split: 'p:split', strips: 'p:strips'
};

/**
 * 构建过渡节点（p:transition）
 * @param {Object} t - 过渡描述 { type, duration(ms) }
 * @returns {Object|null} p:transition 节点
 */
/**
 * 常见进入动画的 presetId（p:cTn@presetId）。
 * PowerPoint 按 preset 名 + presetId 定位具体效果，编号错误会导致动画退化为默认淡入。
 * 未在表中的 preset 回退 1（PowerPoint 会按 preset 名自行修正）。
 */
const ANIMATION_PRESET_IDS: Record<string, number> = {
    appear: 1, flyIn: 2, fly: 2, blinds: 3, blind: 3, box: 4, checkerboard: 5, checker: 5,
    circle: 6, crawl: 7, diamond: 8, dissolve: 9, fade: 10, peek: 11, plus: 12,
    randomBars: 13, random: 13, split: 14, spokes: 15, strips: 16, swivel: 17,
    wedge: 18, wheel: 19, wipe: 20, zoom: 21, bounce: 22, grow: 23, spin: 24,
    pulsate: 1, color: 2, transparency: 3, boldFlash: 4, brush: 5, wave: 6
};

export function buildTransition(t: PptxTransition | undefined, ctx?: SerializerContext): BuilderNode | null {
    if (!t) return null;
    const tag = TRANSITION_TAG[t.type] || 'p:fade';
    const spd = t.duration <= 750 ? '1' : t.duration >= 1500 ? '3' : '2';

    const attrs: Record<string, unknown> = { spd };
    // 自动切换停留时长（毫秒）；与 p:timing 的 afterTime 等价，但 PowerPoint 原生优先读 @advTm
    if (t.advanceAfterTime != null) attrs.advTm = Math.round(t.advanceAfterTime);
    // 禁止点击切换（默认允许，故仅在显式 false 时输出）
    if (t.advanceOnClick === false) attrs.advanceOnClick = 0;

    const children: BuilderNode[] = [];
    children.push(xmlNode(tag, { dir: t.direction || null }));

    // 切换音效（p:sndAc/p:snd），需内嵌音频媒体并登记 hyperlink 之外的 media 关系
    if (t.sound && ctx) {
        const sndAttrs: Record<string, unknown> = {};
        if (t.sound.name) sndAttrs.name = t.sound.name;
        if (t.sound.data) {
            const ext = (t.sound.extension || 'wav').replace(/^\./, '');
            ctx.mediaIndex++;
            const mediaName = `sound${ctx.mediaIndex}.${ext}`;
            ctx.media.push({ name: mediaName, base64: t.sound.data });
            sndAttrs['r:embed'] = addRelationship(ctx, REL_TYPES.audio, `../media/${mediaName}`);
        }
        children.push(xmlNode('p:sndAc', null, xmlNode('p:snd', sndAttrs)));
    }

    return xmlNode('p:transition', attrs, ...children);
}

/**
 * 构建备注幻灯片节点（p:notesSlide）
 * @param {string} notes - 备注文本
 * @returns {Object} p:notesSlide 根节点
 */
export function buildNotesSlide(notes: string): BuilderNode {
    return xmlNode('p:notesSlide',
        { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main' },
        xmlNode('p:cSld', null,
            xmlNode('p:spTree', null,
                xmlNode('p:nvGrpSpPr', null,
                    xmlNode('p:cNvPr', { id: 1, name: '' }),
                    xmlNode('p:cNvGrpSpPr'),
                    xmlNode('p:nvPr')
                ),
                xmlNode('p:grpSpPr', null,
                    xmlNode('a:xfrm', null,
                        xmlNode('a:off', { x: 0, y: 0 }),
                        xmlNode('a:ext', { cx: 0, cy: 0 }),
                        xmlNode('a:chOff', { x: 0, y: 0 }),
                        xmlNode('a:chExt', { cx: 0, cy: 0 })
                    )
                ),
                xmlNode('p:sp', null,
                    xmlNode('p:nvSpPr', null,
                        xmlNode('p:cNvPr', { id: 2, name: 'Notes Placeholder 1' }),
                        xmlNode('p:cNvSpPr', null, xmlNode('p:ph', { type: 'body', idx: 1 })),
                        xmlNode('p:nvPr')
                    ),
                    xmlNode('p:spPr'),
                    xmlNode('p:txBody', null,
                        xmlNode('a:bodyPr', { rtlCol: 0 }),
                        xmlNode('a:lstStyle'),
                        xmlNode('a:p', null,
                            xmlNode('a:r', null,
                                xmlNode('a:rPr', { lang: 'zh-CN' }),
                                xmlNode('a:t', null, notes)
                            )
                        )
                    )
                )
            )
        ),
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}

/**
 * 构建完整幻灯片 XML 根节点（p:sld）
 * @param {Object} ctx - 构建上下文
 * @param {Object} slide - 幻灯片 JSON { background, notes, transition, elements }
 * @returns {Promise<Object>} p:sld 根节点
 */
/** 构建幻灯片计时（自动播放 + 元素动画），生成 p:timing */
function buildTimingNode(slide: SerializerSlide): BuilderNode | null {
    if (!slide) return null;
    // 自动切换时长：slide.advanceTime 与 transition.advanceAfterTime 等价，任一存在即生成计时条件
    const adv = slide.advanceTime ?? slide.transition?.advanceAfterTime;
    const anims = slide.animations || [];
    if (adv == null && anims.length === 0) return null;

    const childNodes: BuilderNode[] = [];
    if (adv != null) {
        childNodes.push(xmlNode('p:cTn',
            { id: 2, fill: 'hold' },
            xmlNode('p:stCondLst', null, xmlNode('p:cond', { type: 'afterTime', val: Math.round(adv) })),
            xmlNode('p:childTnLst', null, xmlNode('p:cTn', { id: 3, pres: 'sld', fill: 'hold' }))
        ));
    }
    let nid = 10;
    for (const a of anims) {
        const spid = a.target != null ? a.target + 2 : 2; // 元素 id 从 2 起连续编号
        // preset 透传真实名称（不再收敛为 4 种），未知类型回退 fade
        const preset = a.type || 'fade';
        const presetClass = (a.presetClass || 'entr') as string;
        // presetId 优先取显式值，否则查常见 preset 编号表
        const presetId = a.presetId != null ? a.presetId : (ANIMATION_PRESET_IDS[preset] ?? 1);

        // 触发时机 → p:stCondLst/p:cond
        const trig = a.trigger?.type || 'afterPrev';
        const condAttrs: Record<string, unknown> = {};
        if (trig === 'onClick') {
            // 「点击时开始」：delay=indefinite 表示等待用户触发
            condAttrs.type = 'begin';
            condAttrs.event = 'delay';
            condAttrs.delay = 'indefinite';
        } else {
            condAttrs.type = trig === 'withPrev' ? 'withPrev' : 'afterPrev';
        }

        const effectChildren: BuilderNode[] = [
            xmlNode('p:stCondLst', null, xmlNode('p:cond', condAttrs))
        ];
        if (a.duration != null) {
            effectChildren.push(xmlNode('p:cTn', { id: nid++, dur: Math.round(a.duration * 1000), fill: 'hold' }));
        }

        const presetAttrs: Record<string, unknown> = {
            id: nid++,
            presetClass,
            presetId,
            type: 'withEffect',
            preset
        };
        if (a.presetSubtype != null) presetAttrs.presetSubtype = a.presetSubtype;
        if (a.delay != null) presetAttrs.delay = Math.round(a.delay * 1000);
        if (a.repeat != null) presetAttrs.repeatCount = a.repeat === 'indefinite' ? 'indefinite' : Math.round(a.repeat * 1000);
        // 路径动画：presetClass='path' 时附加 <p:anim ...> 描述路径（简化为保留 path 文本供上层解析）
        if (presetClass === 'path' && a.path) {
            presetAttrs.presetSubtype = presetAttrs.presetSubtype ?? 0;
        }

        childNodes.push(xmlNode('p:cTn',
            { id: nid++, fill: 'hold' },
            xmlNode('p:tgtEl', null, xmlNode('p:spTgt', { spid })),
            xmlNode('p:childTnLst', null,
                xmlNode('p:cTn', presetAttrs, ...effectChildren)
            )
        ));
    }
    return xmlNode('p:timing',
        null,
        xmlNode('p:tnLst',
            null,
            xmlNode('p:par',
                null,
                xmlNode('p:cTn', { id: 1, fill: 'hold' },
                    xmlNode('p:stCondLst', null, xmlNode('p:cond', { type: 'afterPrev' })),
                    xmlNode('p:childTnLst', null, ...childNodes)
                )
            )
        )
    );
}

export async function buildSlideRoot(ctx: SerializerContext, slide: SerializerSlide) {
    const elementNodes = [];
    for (const el of (slide && slide.elements) || []) {
        const node = await buildElement(ctx, el);
        if (node) elementNodes.push(node);
    }

    // 背景（纯色 / 渐变 / 图片）
    const bgNode = await buildBackground(slide && slide.background, ctx);

    // 过渡效果（传入 ctx 以支持切换音效的媒体关系登记）
    const transitionNode = buildTransition(slide && slide.transition, ctx);
    const timingNode = buildTimingNode(slide);

    return xmlNode('p:sld',
        { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
          show: (slide && slide.hidden) ? 0 : null },
        xmlNode('p:cSld',
            null,
            bgNode,
            xmlNode('p:spTree',
                null,
                xmlNode('p:nvGrpSpPr',
                    null,
                    xmlNode('p:cNvPr', { id: 1, name: '' }),
                    xmlNode('p:cNvGrpSpPr'),
                    xmlNode('p:nvPr')
                ),
                xmlNode('p:grpSpPr',
                    null,
                    xmlNode('a:xfrm',
                        null,
                        xmlNode('a:off', { x: 0, y: 0 }),
                        xmlNode('a:ext', { cx: 0, cy: 0 }),
                        xmlNode('a:chOff', { x: 0, y: 0 }),
                        xmlNode('a:chExt', { cx: 0, cy: 0 })
                    )
                ),
                ...elementNodes
            )
        ),
        transitionNode,
        timingNode,
        xmlNode('p:clrMapOvr', null, xmlNode('a:masterClrMapping'))
    );
}
