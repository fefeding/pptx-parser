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
    /**
     * 东亚字体（a:ea typeface）：中文/日文/韩文字符实际使用的字形。
     * OOXML 把西文与东亚字体分列（a:latin / a:ea），中文 PPT 常写
     * `a:latin="+mn-lt"`（Arial）+ `a:ea="微软雅黑"`；只保留 fontFace 会让中文
     * 回退到系统默认字形（宋体/黑体），与预览端不一致。渲染端应把它追加到
     * font-family 之后，让浏览器按字符集自动回退。
     */
    fontFaceEa?: string;
    href?: string;         // 外部链接或内部跳转 '#N'
    /** 超链接提示（a:hlinkClick@tooltip） */
    hrefTooltip?: string;
    /**
     * 字段（a:fld@type）：动态文本，如 'slidenum' 页码、'datetime' 日期。
     * 解析端读取 a:fld/a:t 作为 text 并回填 type；生成端写出 <a:fld type id><a:t>text</a:t></a:fld>。
     */
    field?: string;
    /** 软换行（a:br）：本 run 之前强制换行，text 为空 */
    break?: boolean;
    /** 文字描边（a:rPr/a:ln）：{ color, width(pt) } */
    outline?: { color?: string; width?: number } | 'none';
    /** 文字外阴影（a:rPr/a:effectLst/a:outerShdw）：{ color, blur(px), x(px), y(px), alpha } */
    shadow?: { color?: string; blur?: number; x?: number; y?: number; alpha?: number };
    /**
     * 基线偏移（上下标）：a:rPr@baseline，1/1000 百分比；>0 上标，<0 下标。0/undefined=正常。
     * 解析端直接复用 OOXML 单位的千分之一比例（如 30000/1000=30% 上标），渲染端据正/负判 super/sub。
     */
    baseline?: number;
    /**
     * 字符间距（字距）：a:rPr/a:spc，单位 px（正=加宽、负=紧缩）。
     * 解析自 a:spcPts（绝对磅值）或 a:spcPct（相对字号比例），统一折算为屏幕像素与渲染端约定一致。
     */
    spacing?: number;
    /** 删除线（a:rPr@strike="1" 或 a:strike 子节点）：CSS text-decoration: line-through */
    strike?: boolean;
    /** 文本高亮（a:rPr/a:highlight）：#RRGGBB 背景色块（非半透明），编辑端渲染为 background-color */
    highlight?: string;
    /**
     * 着重号（a:rPr/a:em@type）：'dot' | 'circle' | 'comma' | 'underDot' 等。
     * 编辑端用 CSS text-emphasis 模拟（位置取 under 以贴近东亚着重号下标习惯）。
     */
    emphasisMark?: string;
    /** 小型大写（a:rPr@cap="small"）：CSS font-variant: small-caps */
    smallCaps?: boolean;
}

/** 段落项目符号：字符（可带符号字体）/ 自动编号 / 图片 */
export interface PptxBullet {
    type: 'number' | 'bullet' | 'picture';
    /** type='bullet'：项目符号字符（配合 font 用符号字体，如 Wingdings 的 U+F0AD） */
    char?: string;
    /** type='bullet'：符号字体名（a:buFont/@typeface），缺省会退化成普通字形 */
    font?: string;
    /** 符号字号百分比（a:buSzPct，100 = 与正文同号） */
    sizePct?: number;
    /** 项目符号颜色（a:buClr）：#RRGGBB，缺省继承文本色 */
    color?: string;
    /** type='number'：编号格式（arabic / romanUpper 等） */
    fmt?: string;
    /** type='number'：起始序号（a:buAutoNum/@startAt） */
    start?: number;
    /** type='picture'：图片 data URL（a:buBlip，已内联 base64） */
    data?: string;
    /** type='picture'：图片关系 id（data 缺失时用于回写定位） */
    rid?: string;
}

/** 段落（可显式 runs，或用 text 配合元素级默认样式） */

export interface PptxParagraph {
    text?: string;
    runs?: PptxTextRun[];
    align?: TextAlign;
    /** 列表样式：true=项目符号；'number'=自动编号；{ type:'number', fmt?, start? } / { type:'bullet', char?, font?, sizePct? } / { type:'picture', data? } */
    bullet?: boolean | 'number' | PptxBullet;
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
    /** 从右到左段落（a:pPr@rtl="1"）；缺省 LTR。RTL 段落默认右对齐 */
    rtl?: boolean;
}

/** 渐变填充色标 */
export interface PptxGradientStop { color: string; position: number; }
/** 纯色填充（可带透明度 0-100） */
export interface PptxFillSolid { type?: 'solid'; color?: string; transparency?: number; }
/** 渐变填充 */
export interface PptxFillGradient {
    type: 'gradient';
    /** 线性渐变方向（a:lin@ang），径向渐变时仍保留以兼容旧数据 */
    direction?: 'horizontal' | 'vertical' | 'diagonal';
    stops: PptxGradientStop[];
    /** 渐变类型：a:lin 为线性（缺省），a:path 为径向/矩形渐变 */
    gradientType?: 'linear' | 'radial';
    /** 径向渐变路径（a:path@path）：circle（默认）/ rect / shape */
    gradientPath?: string;
}
/** 形状填充：颜色串 / {color} / {type:'solid',...} / {type:'gradient',...} / {type:'pattern',...} / {type:'image',...} / 'none' / null */
export type PptxFill = string | PptxFillSolid | PptxFillGradient | PptxFillPattern | PptxFillImage | 'none' | null;

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
/** 图案填充（a:pattFill）：prst 为 OOXML 预设图案名，fg/bg 为前景/背景色 */
export interface PptxFillPattern { type: 'pattern'; prst: string; fg?: string; bg?: string; }
/**
 * 图片填充（a:blipFill）：data 为内联 dataURL，src 为外链地址。
 * 与 p:pic 一致，round-trip 时自包含。
 */
export interface PptxFillImage {
    type: 'image'; data?: string; src?: string; extension?: string;
    srcRect?: PptxImageSrcRect;
    tile?: PptxImageTile;
}

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
    /**
     * 反射（a:effectLst/a:reflection）：水平镜像倒影。
     * 单位：blur/distance 为 pt，angle 为度，alpha 为起始透明度 0-100（越大越淡），scale 为缩放百分比。
     * 编辑端用 -webkit-box-reflect 镜像近似渲染。
     */
    reflection?: { blur?: number; distance?: number; angle?: number; alpha?: number; scale?: number };
    /** 柔化边缘（a:effectLst/a:softEdge@rad），单位 pt */
    softEdge?: { radius: number };
    /** 模糊（a:effectLst/a:blur@rad），单位 pt */
    blur?: { radius: number };
}

/**
 * 三维格式（a:sp3d）：挤出与斜面
 * 单位：extrusionHeight / contourWidth / bevel 尺寸为 pt（生成端转 EMU）
 */
export interface PptxShape3D {
    /** 挤出高度（pt） */
    extrusionHeight?: number;
    /** 轮廓线宽（pt） */
    contourWidth?: number;
    /** 挤出颜色（顶面/侧面材质色） */
    extrusionColor?: string;
    /** 轮廓颜色 */
    contourColor?: string;
    /** 顶部斜面（a:bevelT） */
    bevelTop?: { width?: number; height?: number; preset?: string };
    /** 底部斜面（a:bevelB） */
    bevelBottom?: { width?: number; height?: number; preset?: string };
    /** 材质类型（a:sp3d@prstMaterial） */
    material?: string;
}

/**
 * 三维场景（a:scene3d）：相机与光照
 */
export interface PptxScene3D {
    /** 相机预设（a:camera@prst，如 'orthographicFront' / 'perspectiveRelaxed'） */
    camera?: string;
    /** 相机视场角（a:camera@fov，度） */
    fov?: number;
    /** 相机缩放（a:camera@zoom，百分比） */
    zoom?: number;
    /** 旋转：绕 X 轴（度） */
    rotX?: number;
    /** 旋转：绕 Y 轴（度） */
    rotY?: number;
    /** 旋转：绕 Z 轴（度） */
    rotZ?: number;
    /** 光照预设（a:lightRig@rig，如 'threePt' / 'balanced'） */
    lightRig?: string;
    /** 光照方向（a:lightRig@dir） */
    lightDir?: string;
}

/** 形状三维属性集合 */
export interface Pptx3D {
    /** 形状自身 3D（a:sp3d） */
    shape?: PptxShape3D;
    /** 场景相机/光照（a:scene3d） */
    scene?: PptxScene3D;
}

/**
 * 自定义几何路径（a:custGeom）
 * 坐标已归一化到 0~1（相对形状宽高），避免依赖 EMU 坐标空间。
 */
export interface PptxCustomGeometry {
    /** 路径填充模式；缺省使用形状自身 fill */
    paths: PptxGeometryPath[];
}

/** 自定义几何的单条路径（a:path） */
export interface PptxGeometryPath {
    /** 路径宽（归一化坐标空间） */
    w?: number;
    /** 路径高（归一化坐标空间） */
    h?: number;
    /** 是否闭合（a:path@fill 之外，OOXML 用最后一个 a:close） */
    closed?: boolean;
    /** 路径指令序列 */
    commands: PptxGeometryCommand[];
}

/** 自定义几何指令（a:moveTo / a:lnTo / a:cubicBezTo / a:quadBezTo / a:arcTo / a:close） */
export type PptxGeometryCommand =
    | { type: 'moveTo'; x: number; y: number }
    | { type: 'lnTo'; x: number; y: number }
    | { type: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
    | { type: 'quadBezTo'; x1: number; y1: number; x: number; y: number }
    | { type: 'arcTo'; wR: number; hR: number; stAng: number; swAng: number }
    | { type: 'close' };

/** 文本框自动适配模式（a:bodyPr 的子元素） */
export type PptxAutofit = 'none' | 'normal' | 'shape';

/** 背景填充 */
export type PptxBackground =
    | string                                   // 纯色（#RRGGBB 或颜色名），等价 { type:'solid', color }
    | { type: 'solid'; color: string }
    | { type: 'gradient'; direction?: 'horizontal' | 'vertical' | 'diagonal'; stops: { color: string; position: number }[]; gradientType?: 'linear' | 'radial'; gradientPath?: string }
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
    /**
     * 自动切换停留时长（毫秒，p:transition@advTm）。
     * 与 PptxSlide.advanceTime（走 p:timing 的 stCondLst/cond@afterTime）互为等价表达：
     * 解析时两者都回填，生成时优先写 @advTm（PowerPoint 原生语义），
     * 只有在未指定 advanceOnClick 时才额外补 timing 条件。
     */
    advanceAfterTime?: number;
    /** 切换方向（p:transition@dir 子元素属性，如 8 向 blinds/wipe） */
    direction?: string;
    /**
     * 切换伴随声音（p:snd）：内嵌音频关系目标或 { name } 内置音效名。
     * 生成端对内置名写出 p:snd 的 r:embed 关系（需提供 data），
     * 否则仅记录 name 供上层处理。
     */
    sound?: { name?: string; data?: string; extension?: string };
}

/** 动画类别（p:cTn@presetClass，ECMA-376 ST_TLAnimateBehavior 分类） */
export type PptxAnimationClass = 'entr' | 'exit' | 'emph' | 'path' | 'mediacall';

/** 动画触发时机（p:cond@type / p:cond@delay 组合） */
export type PptxAnimationTrigger =
    /** 上一动画之后 */
    | { type: 'afterPrev'; delay?: number }
    /** 与上一动画同时 */
    | { type: 'withPrev'; delay?: number }
    /** 点击时 */
    | { type: 'onClick'; target?: number; delay?: number };

/** 元素动画（解析自 p:timing 的 p:spTgt，生成端写回 p:timing） */
export interface PptxAnimation {
    /**
     * 目标元素标识。
     * - 字符串：OOXML 形状 spid（cNvPr id），解析端直接透传，生成端直接写入 p:spTgt@spid。
     *   这是推荐形式，对 group 嵌套场景也正确（spid 全局唯一，不依赖元素排列顺序）。
     * - 数字：slide.elements 中的扁平位置索引（向后兼容），生成端按 `index + 2` 推算 spid，
     *   仅适用于无 group 的扁平布局。
     */
    target: number | string;
    /**
     * 动画类型（OOXML preset 名）。
     * 解析端透传真实 preset（如 'flyIn'、'wipe'、'bounce'、'path'…），
     * 不再收敛为 4 种；无法识别时保留原值。
     */
    type: string;
    /** 持续时间（秒） */
    duration: number;
    /**
     * 动画类别：进入(entr) / 退出(exit) / 强调(emph) / 路径(path)。
     * 缺省按 'entr' 处理。
     */
    presetClass?: PptxAnimationClass;
    /** preset 子类型编号（p:cTn@presetId，如 flyIn 的方向变体）；缺省按 type 查表 */
    presetId?: number;
    /** 子类型（p:cTn@presetSubtype） */
    presetSubtype?: number;
    /** 触发时机；缺省 { type:'afterPrev' } */
    trigger?: PptxAnimationTrigger;
    /** 延迟（秒，p:cTn@delay） */
    delay?: number;
    /** 重复次数（p:cTn@repeatCount）；'indefinite' 表示循环 */
    repeat?: number | 'indefinite';
    /** 方向（如 flyIn 的 'l'/'r'/'t'/'b'） */
    direction?: string;
    /** 路径动画的 SVG path 数据（presetClass:'path' 时生效） */
    path?: string;
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
    /** 逐点填充色（c:dPt）：下标对应 values，undefined 表示该点无覆盖 */
    pointColors?: (string | PptxGradientFill | undefined)[];
    /**
     * 系列绑定到哪条数值轴（次坐标轴场景）。
     * 'secondary' 时生成端写出 c:ser/c:order + c:ser 挂到第二个 c:valAx（除 bar 系用 c:catAx 组合外），
     * 需要配合图表级 secondaryValueAxis:true 生效。
     */
    axis?: 'primary' | 'secondary';
    /** 数据标签（c:dLbls）覆盖 */
    dataLabels?: boolean;
    /** 该系列的趋势线（c:trendline） */
    trendlines?: PptxTrendline[];
}

/** 趋势线类型（c:trendlineType@val，ECMA-376） */
export type PptxTrendlineType =
    | 'linear' | 'exp' | 'log' | 'poly' | 'movingAvg' | 'power';

/** 趋势线（c:trendline） */
export interface PptxTrendline {
    type?: PptxTrendlineType;
    /** 显示名称（c:trendline/c:trendlineLbl） */
    name?: string;
    /** 多项式阶数（type:'poly' 时生效，c:order） */
    order?: number;
    /** 移动平均周期（type:'movingAvg' 时生效，c:period） */
    period?: number;
    /** 向前/向后预测周期数（c:forward / c:backward） */
    forward?: number;
    backward?: number;
    /** 显示公式 */
    showEquation?: boolean;
    /** 显示 R² 值 */
    showRSquared?: boolean;
    /** 截距（c:intercept） */
    intercept?: number;
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
    /** 替代文本（无障碍，p:cNvPr@descr） */
    descr?: string;
    /** 是否为装饰性元素（无障碍，p:cNvPr@title 之外的隐藏标记） */
    decorative?: boolean;
    /**
     * 来自版式/母版的非占位符形状（版式装饰设计）。
     * PowerPoint 会把这些形状画在 slide 底层，解析端为还原视觉一并导入；
     * 它们不属于 slide 自身的 spTree，**生成端必须跳过**，否则会写进 slide XML 造成重复。
     */
    inherited?: boolean;
}

/** 文本元素 */
export interface PptxTextElement extends PptxElementBase {
    type: 'text';
    x: number;
    y: number;
    width: number;
    height: number;
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    /** 底层形状类型（带文字的形状，如椭圆/饼图/弧线；编辑器据此还原形状底） */
    shapeType?: string;
    /** 纯文本框标记（p:cNvSpPr@txBox="1"）。预览端对纯文本框在无 a:lnSpc 时兜底 line-height 1.3，
     *  带形状底的文字形状不适用，故需与 shapeType 分开记录。 */
    txBox?: boolean;
    /** 自定义几何（带文字的形状使用了 a:custGeom，如艺术字/特殊剪裁；编辑器据此还原路径） */
    custGeom?: PptxCustomGeometry;
    /** 底层形状填充（与 PptxShapeElement.fill 同构） */
    fill?: PptxFill;
    /** 底层形状边框 */
    line?: PptxLine;
    /** 几何调整值（带文字的形状，如 arc/pie 的 adj） */
    adjust?: Record<string, number>;
    align?: TextAlign;
    valign?: VAlign;
    /** 段落级默认样式（纯 text 模式透传给每个段落）：列表/行距/段间距/缩进 */
    bullet?: boolean | 'number' | PptxBullet;
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
    /** 东亚字体（a:ea typeface），中文实际字形；见 PptxTextRun.fontFaceEa */
    fontFaceEa?: string;
    href?: string;
    /** 文本框内边距（px）：{ l, r, t, b } */
    inset?: { l?: number; r?: number; t?: number; b?: number };
    /**
     * 文字方向（a:bodyPr/@vert，ST_TextVerticalType）：'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl'（默认横排）。
     * 中文竖排用 'eaVert'，逐字堆积用 'wordArtVert'。
     */
    textDirection?: string;
    /** 分栏数（a:bodyPr@numCol，默认 1） */
    numCol?: number;
    /** 是否从右向左排布列（a:bodyPr@rtlCol，影响 RTL 文本折行） */
    rtlCol?: boolean;
    /** 栏间距 pt（a:bodyPr@spcCol） */
    spcCol?: number;
    /**
     * 自动适配（a:bodyPr 的子元素）：
     * - 'none' → a:noAutofit（不缩放）
     * - 'normal' → a:normAutofit（按 fontScale/lnSpcReduction 缩小字号）
     * - 'shape' → a:spAutoFit（缩放形状以贴合文本）
     */
    autofit?: PptxAutofit;
    /** 字号缩放百分比（a:normAutofit@fontScale，配合 autofit:'normal'） */
    fontScale?: number;
    /** 行距缩减百分比（a:normAutofit@lnSpcReduction，配合 autofit:'normal'） */
    lnSpcReduction?: number;
    /**
     * 艺术字变形预设（a:bodyPr@prstTxWarp 的 a:prstTxWarp@prst），
     * 如 'textArchUp' / 'textWave' / 'textCircle' / 'textTriangle'。
     */
    prstTxWarp?: string;
    /**
     * 不换行（a:bodyPr@wrap="none"）：文本不自动折行（与预览端一致）。
     */
    noWrap?: boolean;
    /** 底层形状特效（阴影/发光），来自 spPr/a:effectLst 或 p:style/a:effectRef 主题样式 */
    effects?: PptxShapeEffects;
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
    flipH?: boolean;
    flipV?: boolean;
    fill?: PptxFill;
    line?: PptxLine;
    /** 形状特效：阴影 / 发光（对应 a:effectLst） */
    effects?: PptxShapeEffects;
    /**
     * 几何调整值，key 为该预设形状的 OOXML gd 名（roundRect/snip 用 `adj`，
     * 箭头/标注/星形用 `adj1`/`adj2`/…），如 { adj: 25000 }、{ adj1: 50000, adj2: 40000 }
     */
    adjust?: Record<string, number>;
    /** 自定义几何（a:custGeom）；指定时优先于 shapeType 的预设几何 */
    custGeom?: PptxCustomGeometry;
    /** 三维属性（a:sp3d + a:scene3d） */
    threeD?: Pptx3D;
    /**
     * 语义层未建模的 spPr 特效原始节点（3D 属性 / 双色调 a:duotone / 填充覆盖 a:fillOverlay /
     * 反射 a:reflection / 柔化 a:softEdge / 模糊 a:blur 等），保留以便生成端原样回写，
     * 保证 round-trip 不丢信息。PowerPoint 要求 scene3d/sp3d 位于 a:effectLst 之后，
     * 生成端将其追加在 spPr 末尾以符合顺序约束。
     */
    effectsRaw?: Array<{ tag: string; node: any }>;
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
    /**
     * 语义层未建模的 spPr 特效原始节点（3D 属性 / 双色调 / 填充覆盖 等），
     * 保留以便生成端原样回写，保证 round-trip 不丢信息。
     */
    effectsRaw?: Array<{ tag: string; node: any }>;
    /**
     * 图片 blipFill 级特效原始节点（仅 a:duotone）。
     * 双色调在 OOXML 里只合法于 a:blip 内，绝不能写进 p:spPr（否则结构非法）。
     * 生成端会把它回写到 p:blipFill/a:blip，与 effectsRaw（spPr 级）分开放置。
     */
    blipFx?: Array<{ tag: string; node: any }>;
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
    /** 图例位置（c:legend/c:legendPos@val）：如 'b'/'r'/'t'/'l' */
    legendPosition?: string;
    /** 图表区填充（c:chartSpace/c:spPr）：'none' 或 #RRGGBB；缺省透明 */
    spaceFill?: string;
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
    /** 三维视角（c:view3D）：旋转/厚度/直角轴 */
    view3D?: {
        /** 俯仰角（0-90） */
        rotX?: number;
        /** 旋转角（0-360） */
        rotY?: number;
        /** 厚度百分比（默认 100） */
        depthPercent?: number;
        /** 是否为直角轴（正交投影） */
        rAngAx?: boolean;
    };
    /** 坐标轴标题（c:axTitle） */
    axisTitles?: {
        /** 分类轴标题 */
        category?: string;
        /** 数值轴标题 */
        value?: string;
        /** 次数值轴标题（secondaryValueAxis 生效时） */
        secondaryValue?: string;
    };
    /**
     * 启用次数值轴（第二个 c:valAx + 第二个 c:catAx，配合系列 axis:'secondary'）。
     * 生成端会写出 c:valAx/c:catAx 的第二个实例并分配 axId，同时写 c:barChart 的 c:axIdList。
     */
    secondaryValueAxis?: boolean;
    /** 是否显示数据标签（c:dLbls，系列级可覆盖） */
    dataLabels?: boolean;
    /** 网格线：主要/次要（c:majorGridlines / c:minorGridlines） */
    gridlines?: { major?: boolean; minor?: boolean };
}

/** 渐变填充（形状填充、图表逐点填充 c:dPt 共用） */
export interface PptxGradientFill {
    type: 'gradient';
    /** 线性渐变方向（a:lin@ang 换算：90°→vertical、45°→diagonal、其余 horizontal） */
    direction?: 'horizontal' | 'vertical' | 'diagonal';
    /** 渐变色标（position 为 0~1） */
    stops: { color: string; position: number }[];
    /** 缺省为线性（a:lin）；'radial' 写 a:path */
    gradientType?: 'linear' | 'radial';
    /** 径向渐变路径（a:path@path）：circle / rect / shape */
    gradientPath?: string;
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
    /**
     * 被水平合并吞并的单元格（a:tcPr@hMerge="1"）。
     * OOXML 中被合并区域**仍然存在**这些单元格节点，仅以此标记隐藏，
     * 缺失会导致 PowerPoint 打开时表格结构错乱。
     */
    hMerge?: boolean;
    /** 被垂直合并吞并的单元格（a:tcPr@vMerge="1"），语义同 hMerge */
    vMerge?: boolean;
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
    /**
     * 从右到左段落（a:pPr@rtl="1"）。
     * 单段单元格用 text 简写时无法携带段落级 rtl，故在单元格级冗余一份，
     * 使生成端能写回 rtl="1"（仅 algn="r" 不足以让 RTL 表格复现右对齐渲染）。
     */
    rtl?: boolean;
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
 * 图示缓存绘图中的单个形状（解析自 Microsoft 缓存绘图部件 ppt/diagrams/drawingN.xml）。
 * 坐标为 px、相对图示框左上角；连接线沿包围盒对角绘制（flipH/flipV 决定方向）。
 */
export interface PptxDiagramShape {
    x: number;
    y: number;
    width: number;
    height: number;
    /** OOXML 预设几何：roundRect / ellipse / rect / straightConnector1 ... */
    prst?: string;
    /** 预设几何调整值（解析自 a:prstGeom/a:avLst，如 arc 的 adj1/adj2 角度） */
    adjust?: Record<string, number>;
    /** 连接线（无填充文字，仅描边） */
    connector?: boolean;
    /** 填充色（#RRGGBB），'none' 表示无填充 */
    fill?: string;
    /** 描边色（#RRGGBB） */
    lineColor?: string;
    /** 描边宽度（pt） */
    lineWidth?: number;
    /** 连接线箭头（a:headEnd/a:tailEnd 的 type，如 'triangle'、'stealth'；仅 connector 形状有） */
    startArrow?: string;
    endArrow?: string;
    /** 节点文字（多段以 \n 连接） */
    text?: string;
    /** 字号（pt） */
    fontSize?: number;
    /** 文字颜色（#RRGGBB） */
    color?: string;
    bold?: boolean;
    /** 水平对齐：'l' | 'ctr' | 'r' */
    align?: string;
    /** 垂直对齐：'t' | 'ctr' | 'b' */
    anchor?: string;
    flipH?: boolean;
    flipV?: boolean;
}

/**
 * SmartArt / 图示元素（解析自 p:graphicFrame/a:graphicData[uri=diagram]，或创作生成）
 *
 * 创作时提供 diagramType + nodes 即可生成原生 diagrams/* 部件；
 * 解析端保留可读文本（texts）、缓存绘图形状（shapes）与原始节点（__raw）。
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
    /** 缓存绘图形状（解析自 drawingN.xml，含树形布局坐标，供编辑器/渲染端还原） */
    shapes?: PptxDiagramShape[];
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
    childrenCoordinates?: 'local' | 'page' | 'relative';
    /** 子元素（默认相对组左上角的局部坐标；childrenCoordinates:'page' 时为页绝对坐标） */
    children: PptxElement[];
    /**
     * 组合级旋转（度，顺时针）。对应 grpSpPr/a:xfrm/@rot。
     * 作用于整个组容器（与 PowerPoint 一致：旋转中心为组边界框中心，子元素随组旋转）。
     */
    rotation?: number;
    /** 组合级水平翻转（grpSpPr/a:xfrm/@flipH，1/0/"true"/"false"） */
    flipH?: boolean;
    /** 组合级垂直翻转（grpSpPr/a:xfrm/@flipV，1/0/"true"/"false"） */
    flipV?: boolean;
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
    /** 音频图标/占位图（OOXML 中 a:blip 指向的图，PowerPoint 用它表示音频） */
    poster?: { data?: string; src?: string; extension?: string };
}

/**
 * 连接线元素（p:cxnSp）
 *
 * OOXML 连接线用 a:xfrm 描述**起止两点**（而非 left/top/width/height），
 * 因此这里同时提供：
 * - `x`/`y`/`width`/`height`：由起止点换算的包围盒（与其他元素坐标体系一致）
 * - `start`/`end`：精确端点（px），生成端优先用它写出 a:xfrm@off + a:xfrm@ext
 */
export interface PptxConnectorElement extends PptxElementBase {
    type: 'connector';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 几何类型（如 'straightConnector1' / 'bentConnector3' / 'curvedConnector2'），默认 'straightConnector1' */
    shapeType?: string;
    /** 起点（px） */
    start?: { x: number; y: number };
    /** 终点（px） */
    end?: { x: number; y: number };
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    /** 线型（复用形状边框描述） */
    line?: PptxLine;
    /** 几何调整值（a:avLst） */
    adjust?: Record<string, number>;
}

/** OLE 嵌入对象元素（p:oleObj，通常外裹 p:graphicFrame） */
export interface PptxOleElement extends PptxElementBase {
    type: 'ole';
    x: number;
    y: number;
    width: number;
    height: number;
    /** 程序标识（p:oleObj@progId），如 'Excel.Sheet.12' / 'PowerPoint.Show.12' */
    progId?: string;
    /** 嵌入对象部件名（p:oleObj@r:id 指向 ppt/embeddings/*.xlsx 等） */
    target?: string;
    /** 嵌入文件二进制（base64）；提供时生成端写出部件 */
    data?: string;
    /** 嵌入文件扩展名（如 'xlsx' / 'docx'） */
    extension?: string;
    /** 显示为图标（p:oleObj@showAsIcon="1"） */
    showAsIcon?: boolean;
    /** 图标/预览图（a:blip@r:embed） */
    poster?: { data?: string; src?: string; extension?: string };
}

/** 公式元素（OMML m:oMathPara / m:oMath，通常位于 mc:AlternateContent 内） */
export interface PptxMathElement extends PptxElementBase {
    type: 'math';
    x: number;
    y: number;
    width: number;
    height: number;
    /** OMML XML 字符串（含 m:oMathPara 或 m:oMath 根节点） */
    omml?: string;
    /** 纯文本形式（如 Unicode 线性公式），解析端回退 / 生成端无 omml 时使用 */
    text?: string;
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
    | PptxConnectorElement
    | PptxOleElement
    | PptxMathElement
    | PptxRawElement;

/** 媒体资源（当元素不内联 data 时，通过 id 引用本表） */
export interface PptxMediaResource {
    /** dataURL 或裸 base64 */
    base64: string;
    mime: string;
}

/** 占位符类型（p:ph@type，ECMA-376 ST_PlaceholderType 常用子集） */
export type PptxPlaceholderType =
    | 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'obj'
    | 'ftr' | 'sldNum' | 'dt'
    | 'pic' | 'tbl' | 'chart' | 'media' | 'clipArt' | 'dgm';

/** 母版/版式中的占位符定义（p:sp + p:nvSpPr/p:nvPr/p:ph） */
export interface PptxPlaceholder {
    type: PptxPlaceholderType;
    x: number;
    y: number;
    width: number;
    height: number;
    /** 占位符索引（p:ph@idx），用于与幻灯片元素按 idx 匹配继承 */
    idx?: number;
    /** 提示文本（p:ph 无文本时的灰字提示，仅版式层有效） */
    prompt?: string;
    /** 元素名称 */
    name?: string;
    /** 该占位符的默认文本样式 */
    fontSize?: number;
    color?: string;
    bold?: boolean;
    fontFace?: string;
    align?: TextAlign;
    valign?: VAlign;
    /** 项目符号默认样式（body 占位符常见） */
    bullet?: boolean | 'number' | PptxBullet;
}

/** 幻灯片版式（ppt/slideLayouts/slideLayoutN.xml） */
export interface PptxSlideLayout {
    /** 版式名称（p:cSld@name） */
    name?: string;
    /** 版式背景（缺省继承母版） */
    background?: PptxBackground;
    /** 版式上的常驻元素（非占位符，如 logo、装饰） */
    elements?: PptxElement[];
    /** 占位符定义 */
    placeholders?: PptxPlaceholder[];
    /** 是否显示母版背景图形（p:sldLayout@showMasterSp，默认 true） */
    showMasterSp?: boolean;
}

/** 幻灯片母版（ppt/slideMasters/slideMasterN.xml） */
export interface PptxSlideMaster {
    /** 母版名称 */
    name?: string;
    /** 母版背景 */
    background?: PptxBackground;
    /** 母版上的常驻元素 */
    elements?: PptxElement[];
    /**
     * 母版级占位符（title/body/ftr/sldNum/dt 的**默认位置与样式**）。
     * 版式未覆盖的占位符从这里继承。
     */
    placeholders?: PptxPlaceholder[];
    /** 该母版下的版式列表 */
    layouts?: PptxSlideLayout[];
    /** 该母版引用的主题整串 XML（解析端回读，无损保真；与 presentation 共享或独立 theme 部件） */
    themeXml?: string;
}

/** 幻灯片 */
export interface PptxSlide {
    /**
     * 使用的版式索引（0 基，按 PptxDocument.masters[].layouts 展平后的顺序）。
     * 省略时使用第 0 个版式。需要多母版/多版式时配合 PptxDocument.masters 使用。
     */
    layout?: number;
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

/** 主题配色方案（a:clrScheme 的 12 个色槽） */
export interface PptxThemeColorScheme {
    name?: string;
    dk1?: string; lt1?: string;
    dk2?: string; lt2?: string;
    accent1?: string; accent2?: string; accent3?: string;
    accent4?: string; accent5?: string; accent6?: string;
    hlink?: string; folHlink?: string;
}

/** 主题字体组（a:fontScheme 的 majorFont/minorFont） */
export interface PptxThemeFonts {
    /** 拉丁字体 */
    latin?: string;
    /** 东亚字体（中日韩） */
    ea?: string;
    /** 复杂文种字体（阿拉伯/希伯来等） */
    cs?: string;
}

/** 主题字体方案 */
export interface PptxThemeFontScheme {
    name?: string;
    /** 标题字体（majorFont） */
    major?: PptxThemeFonts;
    /** 正文字体（minorFont） */
    minor?: PptxThemeFonts;
}

/**
 * 主题定义（a:theme）
 *
 * 语义级对象：提供 colors/fonts 时生成端据此构造完整 `themeN.xml`；
 * 其余未知键允许透传（兼容旧的整串 XML 覆盖用法）。
 */
export interface PptxTheme {
    /** 主题名（a:theme@name） */
    name?: string;
    /** 配色方案 */
    colors?: PptxThemeColorScheme;
    /** 字体方案 */
    fonts?: PptxThemeFontScheme;
    [key: string]: unknown;
}

/** 文档节（presentation.xml 的 p:sectPr / p:section） */
export interface PptxSection {
    /** 节名称（p:section@name） */
    name?: string;
    /** 该节包含的幻灯片索引（0 基，对应 PptxDocument.slides 下标） */
    slides: number[];
}

/** 嵌入字体资源（ppt/fonts/fontN.fntdata + ppt/fontTable.xml） */
export interface PptxFontResource {
    /** 字体族名（a:font@typeface，如 'Source Han Sans'） */
    name: string;
    /**
     * 字体数据（base64）。OOXML 要求嵌入字体为 **经 XOR 混淆** 的 fntdata，
     * 生成端会按 ECMA-376 规则对原始 TTF/OTF 做混淆后写入；
     * 解析端回读时自动解混淆并还原为原始字节。
     */
    data: string;
    /** panose 分类（a:font@panose，20 位十六进制串） */
    panose?: string;
    /** 是否为粗体变体 */
    bold?: boolean;
    /** 是否为斜体变体 */
    italic?: boolean;
    /**
     * 嵌入方式：
     * - 'full'（默认）：完整嵌入（a:font@embed='embed'）
     * - 'subset'：仅嵌入用到的字形子集
     */
    embedType?: 'full' | 'subset';
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
    /** 主题覆盖（可选，高级）：字符串=整串 theme XML；对象=语义级主题定义 */
    theme?: PptxTheme | string;
    /** 原始 theme1.xml 整串文本（解析端回读，用于无损回退以保证 fmtScheme/fontScheme 等细节保真） */
    themeXml?: string;
    /**
     * 原始 ppt/tableStyles.xml 整串文本（解析端回读）。
     * 表格样式的 GUID 定义决定网格线颜色/底纹/条带，重新生成的等价定义只有通用黑网格，
     * 会让白网格表格变黑、底纹与条带丢失。
     */
    tableStylesXml?: string;
    /**
     * 原始主题部件整串列表（解析端回读，顺序即 ppt/theme/themeN.xml 的 N）。
     * 多母版/多主题文件（如 Sample_12 的不同页绑定不同主题）按 masters[i].themeXml
     * 各自写回对应主题部件，避免所有页退化为单一主题导致 SmartArt/图表配色错位。
     */
    themeXmls?: string[];
    /**
     * 母版定义（可选）。提供时生成端按此写出多个 slideMasterN.xml 及其版式，
     * 并让幻灯片通过 PptxSlide.layout 指定所用版式。省略时退回单一空白版式。
     */
    masters?: PptxSlideMaster[];
    /** 文档节（presentation.xml 的 p:section）；省略时不分节 */
    sections?: PptxSection[];
    /** 嵌入字体（ppt/fonts/fontN.fntdata + ppt/fontTable.xml） */
    fonts?: PptxFontResource[];
}
