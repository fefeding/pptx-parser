/**
 * OOXML 静态模板模块
 *
 * 存放生成 PPTX 包所需的固定 XML 部件模板：
 * 主题（theme1.xml）、母版（slideMaster1.xml）、版式（slideLayout1.xml）、
 * 演示文稿属性（presProps/viewProps/tableStyles）等。
 * 模板保持最小可用结构，参考标准 Office 输出。
 *
 * @module serializer/templates
 */

import { NS, escapeXml, pxToEmu, ptToSz, colorToHex } from './xml-builder';
import type {
    PptxTheme, PptxThemeColorScheme, PptxThemeFontScheme, PptxThemeFonts,
    PptxSlideMaster, PptxSlideLayout, PptxPlaceholder, PptxBackground
} from '../types/pptx-document';

/** DrawingML 常用命名空间声明串 */
const DRAWING_NS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

/**
 * 主题模板（Office 标准配色/字体方案，最小化 fmtScheme）
 * @returns {string} theme1.xml 内容
 */
/** 内置默认主题（Office 标准配色/字体 + 最小化 fmtScheme） */
const DEFAULT_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="110000"/></a:schemeClr></a:gs><a:gs pos="1000"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="4553000" scaled="0"/></a:gradFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="97000"/><a:satMod val="130000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

/** 主题 12 色槽的默认 sRGB 值（未指定时使用） */
const THEME_COLOR_DEFAULTS: Record<string, string> = {
    dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
    accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
    accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
    hlink: '0563C1', folHlink: '954F72'
};

/**
 * 构造 a:clrScheme（主题配色方案）
 * @param {Object} c - 配色定义（缺省槽位用 Office 默认值补齐）
 * @returns {string} a:clrScheme XML
 */
function buildClrScheme(c: PptxThemeColorScheme): string {
    const inner = Object.keys(THEME_COLOR_DEFAULTS)
        .map(k => {
            const val = (c as Record<string, string | undefined>)[k] || THEME_COLOR_DEFAULTS[k];
            return `<a:${k}><a:srgbClr val="${colorToHex(val)}"/></a:${k}>`;
        })
        .join('');
    return `<a:clrScheme name="${escapeXml(c.name || 'Custom')}">${inner}</a:clrScheme>`;
}

/**
 * 构造 a:fontScheme（主题字体方案）
 * @param {Object} f - 字体定义
 * @returns {string} a:fontScheme XML
 */
function buildFontScheme(f: PptxThemeFontScheme): string {
    const grp = (tag: string, g: PptxThemeFonts | undefined, defLatin: string) =>
        `<a:${tag}><a:latin typeface="${escapeXml(g?.latin || defLatin)}"/>` +
        `<a:ea typeface="${escapeXml(g?.ea || '')}"/><a:cs typeface="${escapeXml(g?.cs || '')}"/></a:${tag}>`;
    return `<a:fontScheme name="${escapeXml(f.name || 'Custom')}">` +
        grp('majorFont', f.major, 'Calibri Light') + grp('minorFont', f.minor, 'Calibri') +
        `</a:fontScheme>`;
}

/**
 * 主题模板：支持语义级主题定义（colors / fonts），也支持整串 XML 覆盖
 * @param {Object|string} [theme] - 主题定义对象或完整 XML 字符串
 * @returns {string} theme1.xml 内容
 */
export function buildThemeXml(theme?: PptxTheme | string) {
    // 整串 XML 覆盖（旧用法）
    if (typeof theme === 'string') return theme;
    let xml = DEFAULT_THEME_XML;
    if (theme?.colors) {
        xml = xml.replace(/<a:clrScheme[^>]*>[\s\S]*?<\/a:clrScheme>/, buildClrScheme(theme.colors));
    }
    if (theme?.fonts) {
        xml = xml.replace(/<a:fontScheme[^>]*>[\s\S]*?<\/a:fontScheme>/, buildFontScheme(theme.fonts));
    }
    if (theme?.name) {
        xml = xml.replace(/<a:theme([^>]*)name="[^"]*"/, `<a:theme$1name="${escapeXml(theme.name)}"`);
    }
    return xml;
}

/** 默认母版模板（空白母版，仅含布局引用与颜色映射） */
const DEFAULT_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${DRAWING_NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

/** 默认版式模板（空白版式） */
const DEFAULT_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${DRAWING_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

/**
 * 构建占位符形状（p:sp + p:nvPr/p:ph）
 *
 * 占位符是母版/版式的核心机制：幻灯片元素通过 p:ph 的 type/idx 继承
 * 位置与默认文本样式；页脚/页码/日期（ftr/sldNum/dt）即由此实现。
 *
 * @param {Object} ph - 占位符定义
 * @param {number} id - 形状 id（母版/版式内唯一，从 2 起）
 * @returns {string} p:sp XML
 */
export function buildPlaceholderSpXml(ph: PptxPlaceholder, id: number): string {
    const anchorMap: Record<string, string> = { top: 't', middle: 'ctr', bottom: 'b' };
    const anchor = ph.valign ? anchorMap[ph.valign] : undefined;
    const bodyPr = `<a:bodyPr${anchor ? ` anchor="${anchor}"` : ''} rtlCol="0"/>`;

    const rPrAttrs: string[] = ['lang="zh-CN"', 'dirty="0"'];
    if (ph.fontSize) rPrAttrs.push(`sz="${ptToSz(ph.fontSize)}"`);
    if (ph.bold) rPrAttrs.push('b="1"');
    if (ph.fontFace) rPrAttrs.push('');
    const rPrChildren = ph.fontFace ? `<a:latin typeface="${escapeXml(ph.fontFace)}"/>` : '';
    const rPr = `<a:rPr ${rPrAttrs.join(' ')}>${rPrChildren}</a:rPr>`;
    // 提示文本（版式层灰字）；无提示时留空段落
    const para = ph.prompt
        ? `<a:p><a:r>${rPr}<a:t>${escapeXml(ph.prompt)}</a:t></a:r></a:p>`
        : '<a:p><a:endParaRPr lang="zh-CN" dirty="0"/></a:p>';

    const idxAttr = ph.idx != null ? ` idx="${ph.idx}"` : '';

    return `<p:sp>` +
        `<p:nvSpPr>` +
        `<p:cNvPr id="${id}" name="${escapeXml(ph.name || `${ph.type} Placeholder ${id - 1}`)}"/>` +
        `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
        `<p:nvPr><p:ph type="${escapeXml(ph.type)}"${idxAttr}/></p:nvPr>` +
        `</p:nvSpPr>` +
        `<p:spPr>` +
        `<a:xfrm><a:off x="${pxToEmu(ph.x || 0)}" y="${pxToEmu(ph.y || 0)}"/>` +
        `<a:ext cx="${pxToEmu(ph.width || 0)}" cy="${pxToEmu(ph.height || 0)}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</p:spPr>` +
        `<p:txBody>${bodyPr}<a:lstStyle/>${para}</p:txBody>` +
        `</p:sp>`;
}

/**
 * 构建背景 XML（p:bg）
 * 支持纯色与线性渐变；图片背景因需登记媒体关系，交由 element-builders 处理。
 * @param {Object|string} bg - 背景定义
 * @returns {string} p:bg XML（不支持时返回空串）
 */
export function buildBackgroundXml(bg?: PptxBackground | string | null): string {
    if (!bg) return '';
    if (typeof bg === 'string') {
        return `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${colorToHex(bg)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    }
    if (bg.type === 'solid') {
        return `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${colorToHex(bg.color)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    }
    if (bg.type === 'gradient') {
        const angMap: Record<string, number> = { horizontal: 0, vertical: 5400000, diagonal: 4500000 };
        const ang = angMap[bg.direction || 'horizontal'] ?? 0;
        const stops = bg.stops
            .map(s => `<a:gs pos="${Math.round(s.position * 100000)}"><a:srgbClr val="${colorToHex(s.color)}"/></a:gs>`)
            .join('');
        return `<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>`;
    }
    return '';
}

/**
 * 母版模板：支持自定义占位符、背景与多个版式引用
 * @param {Object|string} [master] - 母版定义或完整 XML 字符串
 * @param {Array<string>} [layoutRelIds] - 该母版下各版式的关系 id（默认单个 rId1）
 * @param {string} [elementXml] - 母版上常驻元素的 XML（由 element-builders 生成）
 * @returns {string} slideMasterN.xml 内容
 */
export function buildSlideMasterXml(master?: PptxSlideMaster | string, layoutRelIds?: string[], elementXml = ''): string {
    if (typeof master === 'string') return master;
    let xml = DEFAULT_MASTER_XML;

    const relIds = layoutRelIds && layoutRelIds.length ? layoutRelIds : ['rId1'];
    const layoutIds = relIds
        .map((rid, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${escapeXml(rid)}"/>`)
        .join('');
    xml = xml.replace(/<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/, `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>`);

    // 占位符 + 常驻元素插入 spTree
    const placeholders = (master?.placeholders || []);
    const inner = placeholders.map((ph, i) => buildPlaceholderSpXml(ph, 2 + i)).join('') + elementXml;
    if (inner) xml = xml.replace('</p:spTree>', `${inner}</p:spTree>`);

    if (master?.background) {
        const bgXml = buildBackgroundXml(master.background);
        if (bgXml) xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/, bgXml);
    }
    return xml;
}

/**
 * 版式模板：支持自定义占位符、背景与常驻元素
 * @param {Object|string} [layout] - 版式定义或完整 XML 字符串
 * @param {string} [elementXml] - 版式上常驻元素的 XML（由 element-builders 生成）
 * @returns {string} slideLayoutN.xml 内容
 */
export function buildSlideLayoutXml(layout?: PptxSlideLayout | string, elementXml = ''): string {
    if (typeof layout === 'string') return layout;
    let xml = DEFAULT_LAYOUT_XML;

    const placeholders = (layout?.placeholders || []);
    const inner = placeholders.map((ph, i) => buildPlaceholderSpXml(ph, 2 + i)).join('') + elementXml;
    if (inner) xml = xml.replace('</p:spTree>', `${inner}</p:spTree>`);

    if (layout?.name) {
        xml = xml.replace(/<p:cSld([^>]*)name="[^"]*"/, `<p:cSld$1name="${escapeXml(layout.name)}"`);
    }
    if (layout?.showMasterSp === false) {
        xml = xml.replace('<p:sldLayout ', '<p:sldLayout showMasterSp="0" ');
    }
    if (layout?.background) {
        const bgXml = buildBackgroundXml(layout.background);
        if (bgXml) xml = xml.replace('<p:spTree>', `${bgXml}<p:spTree>`);
    }
    return xml;
}

/**
 * 演示文稿属性模板（presProps.xml）
 * @returns {string} presProps.xml 内容
 */
export function buildPresPropsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentationPr xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`;
}

/**
 * 视图属性模板（viewProps.xml）
 * @returns {string} viewProps.xml 内容
 */
export function buildViewPropsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:viewProps xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`;
}

/**
 * 默认表格样式 ID（对应内置 “Table Grid”：纯网格线、无底纹）
 * 使用内置 GUID 可让 PowerPoint 直接解析，同时我们在 tableStyles.xml 中给出等价定义供解析器使用。
 */
export const DEFAULT_TABLE_STYLE_ID = '{5940675A-B579-460E-94D1-54222C63F5DA}';
/**
 * 内置 “Medium Style 2 - Accent 1”：首行/首列强调色底 + 白色文字 + 白色网格线。
 * 同样在 tableStyles.xml 中给出等价定义，避免依赖 WPS/PowerPoint 各自的内置样式表。
 */
export const ACCENT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

/** 六向网格线（颜色以 DrawingML 颜色片段给出，如 `<a:srgbClr val="000000"/>`） */
function tableGridEdges(colorXml: string): string {
    const edge = (side: string) => `<a:${side}><a:ln w="12700" cmpd="sng"><a:solidFill>${colorXml}</a:solidFill></a:ln></a:${side}>`;
    return ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map(edge).join('');
}

/** Table Grid 等价定义（黑网格 + 首行加粗） */
function tableGridStyleXml(styleId: string): string {
    const grid = tableGridEdges('<a:srgbClr val="000000"/>');
    return `<a:tblStyle styleId="${styleId}" styleName="Table Grid">` +
        `<a:wholeTbl>` +
        `<a:tcTxStyle b="off"><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr></a:tcStyle>` +
        `</a:wholeTbl>` +
        `<a:firstRow>` +
        `<a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr></a:tcStyle>` +
        `</a:firstRow>` +
        `</a:tblStyle>`;
}

/** Medium Style 2 - Accent 1 等价定义（强调色底纹 + 白色网格线） */
function accentTableStyleXml(styleId: string): string {
    const grid = tableGridEdges('<a:schemeClr val="lt1"/>');
    const accentFill = `<a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill>`;
    // 交替行：强调色 40%（lumMod 40% + lumOff 60% ≈ 原色 40% 亮度）
    const bandFill = `<a:fill><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr></a:solidFill></a:fill>`;
    return `<a:tblStyle styleId="${styleId}" styleName="Medium Style 2 - Accent 1">` +
        `<a:wholeTbl>` +
        `<a:tcTxStyle b="off"><a:fontRef idx="minor"/><a:schemeClr val="lt1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr>${accentFill}</a:tcStyle>` +
        `</a:wholeTbl>` +
        // 区域名必须是 CT_TableStyle 里的 band1H（写成 band1Horz 会被 PowerPoint/WPS 忽略）
        `<a:band1H><a:tcStyle>${bandFill}</a:tcStyle></a:band1H>` +
        `<a:firstRow>` +
        `<a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="lt1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr>${accentFill}</a:tcStyle>` +
        `</a:firstRow>` +
        `</a:tblStyle>`;
}

/**
 * 表格样式模板（tableStyles.xml）
 *
 * 根元素必须是 DrawingML 命名空间下的 a:tblStyleLst（解析端按 ["a:tblStyleLst"]["a:tblStyle"] 读取）。
 * 除内置的两个样式外，还会为文档实际引用到的每个 tableStyleId 补一份等价定义 ——
 * WPS/PowerPoint 与本解析器遇到未知 GUID 都会退化为「无样式无网格」，表格会整个看不见边框。
 *
 * @param {string[]} [styleIds] - 文档中引用到的表格样式 ID
 * @returns {string} tableStyles.xml 内容
 */
export function buildTableStylesXml(styleIds: string[] = []) {
    const ids = new Set<string>([DEFAULT_TABLE_STYLE_ID, ...styleIds]);
    const styles = [...ids]
        .map((id) => (id === ACCENT_TABLE_STYLE_ID ? accentTableStyleXml(id) : tableGridStyleXml(id)))
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<a:tblStyleLst xmlns:a="${NS.a}" def="${DEFAULT_TABLE_STYLE_ID}">${styles}</a:tblStyleLst>`;
}

/**
 * 生成 docProps/custom.xml（自定义文档属性）
 * @param {Object} pairs - 键值对（名称 → 值）
 * @returns {string} custom.xml 内容
 */
export function buildCustomPropsXml(pairs: Record<string, string>) {
    const props = Object.entries(pairs)
        .map(([name, value], i) =>
            `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${escapeXml(name)}"><vt:lpwstr>${escapeXml(value)}</vt:lpwstr></property>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ` +
        `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${props}</Properties>`;
}

/** 注释作者信息 */
export interface CommentAuthorInfo {
    /** 作者编号（comments 部件通过 authorId 引用） */
    id: number;
    /** 作者显示名 */
    name: string;
    /** 作者缩写（缺省取 name 首字母） */
    initials?: string;
    /** 作者批注颜色（ARGB hex，缺省 FF000000） */
    color?: string;
    /** 该作者最后一条批注 id（缺省 0） */
    lastIdx?: number;
}

/** 单条批注（已解析 authorId） */
export interface CommentInfo {
    /** 引用的作者编号 */
    authorId: number;
    /** 批注正文 */
    text: string;
    /** 批注时间（ISO 8601） */
    dt: string;
    /** 批注锚点 x（EMU） */
    x?: number;
    /** 批注锚点 y（EMU） */
    y?: number;
}

/**
 * 生成 ppt/comments/commentsN.xml（批注列表 p:cmLst）
 * @param {CommentInfo[]} comments - 已解析作者的批注数组
 * @returns {string} commentsN.xml 内容
 */
export function buildCommentsXml(comments: CommentInfo[]) {
    const items = comments
        .map((c, i) =>
            `<p:cm authorId="${c.authorId}" dt="${escapeXml(c.dt)}" id="${i}">` +
            `<p:pos x="${c.x ?? 914400}" y="${c.y ?? 914400}"/>` +
            `<p:text>${escapeXml(c.text)}</p:text></p:cm>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<p:cmLst xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">${items}</p:cmLst>`;
}

/**
 * 生成 ppt/commentAuthors.xml（批注作者列表）
 * @param {CommentAuthorInfo[]} authors - 作者数组（按 id 升序）
 * @returns {string} commentAuthors.xml 内容
 */
export function buildCommentAuthorsXml(authors: CommentAuthorInfo[]) {
    const items = authors
        .map((a) =>
            `<p:cmAuthor id="${a.id}" name="${escapeXml(a.name)}" ` +
            `initials="${escapeXml(a.initials || (a.name ? a.name[0] : 'A'))}" ` +
            `lastIdx="${a.lastIdx ?? 0}" clr="${escapeXml(a.color || 'FF000000')}"/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<p:commentAuthors xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">${items}</p:commentAuthors>`;
}

/**
 * 生成 presentation.xml
 * @param {Object} slideSize - 幻灯片尺寸（px）
 * @param {number} slideSize.width - 宽（px）
 * @param {number} slideSize.height - 高（px）
 * @param {Array<{relId: string}>} slides - 幻灯片引用列表
 * @returns {string} presentation.xml 内容
 */
export function buildPresentationXml(
    slideSize: { width: number; height: number },
    slides: Array<{ relId: string }>,
    opts: {
        /** 各母版在 presentation.xml.rels 中的关系 id（默认单个 rId1） */
        masterRelIds?: string[];
        /** 文档节：每节包含若干 sldId（与 slides 顺序对应的 0 基下标） */
        sections?: { name?: string; slides: number[] }[];
        /** 备注母版在 presentation.xml.rels 中的关系 id（提供时输出 p:notesMasterIdLst） */
        notesMasterRelId?: string;
    } = {}
) {
    const slideEntries = slides
        .map((s, i) => `<p:sldId id="${256 + i}" r:id="${escapeXml(s.relId)}"/>`)
        .join('');

    const masterRelIds = opts.masterRelIds?.length ? opts.masterRelIds : ['rId1'];
    const masterEntries = masterRelIds
        .map((rid, i) => `<p:sldMasterId id="${2147483648 + i}" r:id="${escapeXml(rid)}"/>`)
        .join('');

    // 节（p:sectionLst）：位于 sldIdLst 之后，节内以 sldId 的数值 id 引用幻灯片
    let sectionXml = '';
    if (opts.sections && opts.sections.length) {
        const items = opts.sections.map((sec, i) => {
            const ids = (sec.slides || [])
                .map(idx => `<p:sldId id="${256 + idx}"/>`)
                .join('');
            return `<p:section name="${escapeXml(sec.name || `Section ${i + 1}`)}" id="{${sectionGuid(i)}}"><p:sldIdLst>${ids}</p:sldIdLst></p:section>`;
        }).join('');
        sectionXml = `<p:sectionLst>${items}</p:sectionLst>`;
    }

    // 备注母版：位于 sldMasterIdLst 之后、sldIdLst 之前
    const notesMasterXml = opts.notesMasterRelId
        ? `<p:notesMasterIdLst><p:notesMasterId r:id="${escapeXml(opts.notesMasterRelId)}"/></p:notesMasterIdLst>`
        : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1"><p:sldMasterIdLst>${masterEntries}</p:sldMasterIdLst>${notesMasterXml}<p:sldIdLst>${slideEntries}</p:sldIdLst>${sectionXml}<p:sldSz cx="${pxToEmu(slideSize.width)}" cy="${pxToEmu(slideSize.height)}"/><p:notesSz cx="6858000" cy="914400"/><p:defaultTextStyle/></p:presentation>`;
}

/**
 * OOXML 字体嵌入混淆密钥（ECMA-376 / MS-ODRAWXML 的 16 字节循环 XOR key）
 * 嵌入字体部件必须经此混淆后写入，否则 PowerPoint 无法识别。
 */
const FONT_OBFUSCATION_KEY = [
    0x05, 0x00, 0xEC, 0x03, 0x4B, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
];

/**
 * 对字体数据做 OOXML 嵌入混淆（16 字节循环 XOR）
 * @param {Uint8Array} bytes - 原始字体字节
 * @returns {Uint8Array} 混淆后的字节
 */
export function obfuscateFontData(bytes: Uint8Array): Uint8Array {
    const out = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) {
        out[i] = bytes[i] ^ FONT_OBFUSCATION_KEY[i % 16];
    }
    return out;
}

/**
 * 生成 ppt/fontTable.xml（嵌入字体登记表）
 * @param {Array} fonts - 字体定义 [{ name, panose, bold, italic, embedType }]
 * @param {Array<string>} relIds - 各字体部件在 fontTable.xml.rels 中的关系 id
 * @returns {string} fontTable.xml 内容
 */
export function buildFontTableXml(
    fonts: Array<{ name: string; panose?: string; bold?: boolean; italic?: boolean; embedType?: 'full' | 'subset' }>,
    relIds: string[]
): string {
    const entries = fonts.map((f, i) => {
        const panose = f.panose ? `<a:panose val="${escapeXml(f.panose)}"/>` : '';
        const styleAttrs = `${f.bold ? ' b="1"' : ''}${f.italic ? ' i="1"' : ''}`;
        const subsetted = f.embedType === 'subset' ? '1' : '0';
        const embedded = relIds[i]
            ? `<a:embeddedFont embed="embed" subsetted="${subsetted}"><a:fontData r:id="${escapeXml(relIds[i])}"/></a:embeddedFont>`
            : '';
        return `<a:font script="Latn" typeface="${escapeXml(f.name)}"${styleAttrs}>${panose}` +
            `<a:charset val="0"/><a:family val="roman"/><a:pitch val="variable"/>${embedded}</a:font>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:fontTable ${DRAWING_NS}>${entries}</p:fontTable>`;
}

/**
 * 生成备注母版（ppt/notesMasters/notesMaster1.xml）
 * 备注页必须挂在备注母版上，否则 PowerPoint 打开后备注区无版式可继承。
 * @returns {string} notesMaster1.xml 内容
 */
export function buildNotesMasterXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:notesMaster ${DRAWING_NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
        `<p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN" dirty="0"/></a:p></p:txBody></p:sp>` +
        `</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

/** 节的确定性 GUID（避免每次生成不同导致无意义 diff） */
function sectionGuid(seed: number): string {
    const base = '00000000-0000-4000-8000-000000000000';
    const suffix = String(seed).padStart(12, '0');
    return `${base.slice(0, -suffix.length)}${suffix}`;
}

/**
 * 生成 presentation.xml.rels
 * @param {Array<{relId: string, target: string, type: string, external?: boolean}>} rels - 关系列表
 * @returns {string} presentation.xml.rels 内容
 */
export function buildRelationshipsXml(rels: Array<{ relId: string; type: string; target: string; external?: boolean }>) {
    const entries = rels
        .map((r) => `<Relationship Id="${escapeXml(r.relId)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rel}">${entries}</Relationships>`;
}

/**
 * 生成 [Content_Types].xml
 * @param {Array<string>} mediaExts - 用到的媒体扩展名（如 ['png','jpeg']）
 * @param {number} slideCount - 幻灯片数量
 * @returns {string} [Content_Types].xml 内容
 */
export function buildContentTypesXml(
    mediaExts: Iterable<unknown>,
    slideCount: number,
    opts: { masterCount?: number; layoutCount?: number } = {}
) {
    const MIME_MAP = {
        png: 'image/png',
        jpeg: 'image/jpeg',
        jpg: 'image/jpeg',
        gif: 'image/gif',
        bmp: 'image/bmp',
        svg: 'image/svg+xml',
        tiff: 'image/tiff',
        webp: 'image/webp',
        emf: 'image/x-emf',
        wmf: 'image/x-wmf'
    };
    const defaults = ['rels', 'xml']
        .map(ext => `<Default Extension="${ext}" ContentType="${ext === 'rels' ? 'application/vnd.openxmlformats-package.relationships+xml' : 'application/xml'}"/>`)
        .join('');
    const mediaDefaults = [...new Set(mediaExts)]
        .map(ext => `<Default Extension="${escapeXml(ext)}" ContentType="${(MIME_MAP as Record<string, string>)[String(ext).toLowerCase()] || 'application/octet-stream'}"/>`)
        .join('');
    const slideOverrides = Array.from({ length: slideCount }, (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${mediaDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${Array.from({ length: Math.max(1, opts.masterCount ?? 1) }, (_, i) =>
        `<Override PartName="/ppt/slideMasters/slideMaster${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`
    ).join('')}${Array.from({ length: Math.max(1, opts.layoutCount ?? 1) }, (_, i) =>
        `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`
    ).join('')}<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>${slideOverrides}</Types>`;
}

/**
 * 生成 docProps/core.xml（元数据）
 * @param {Object} metadata - 元数据（title/author/subject/keywords/description/lastModifiedBy/category/status）
 * @returns {string} core.xml 内容
 */
export function buildCorePropsXml(metadata?: Record<string, unknown>) {
    const md: Record<string, unknown> = metadata || {};
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const created = md.created || now;
    const modified = md.modified || now;
    const fields = [
        ['dc:title', md.title],
        ['dc:subject', md.subject],
        ['dc:creator', md.author],
        ['cp:keywords', md.keywords],
        ['dc:description', md.description],
        ['cp:lastModifiedBy', md.lastModifiedBy || md.author],
        ['cp:category', md.category],
        ['cp:contentStatus', md.status]
    ]
        .filter(([, val]) => val !== undefined && val !== null && val !== '')
        .map(([tag, val]) => `<${tag}>${escapeXml(val)}</${tag}>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">${fields}<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${modified}</dcterms:modified></cp:coreProperties>`;
}

/**
 * 生成 docProps/app.xml
 * @param {number} slideCount - 幻灯片数量
 * @returns {string} app.xml 内容
 */
export function buildAppPropsXml(slideCount: number) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Properties xmlns="${NS.ext}" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office PowerPoint</Application><Slides>${slideCount}</Slides><AppVersion>16.0000</AppVersion></Properties>`;
}

/**
 * 生成根关系 _rels/.rels
 * @returns {string} .rels 内容
 */
export function buildRootRelsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

/** 母版关系：rId1 版式、rId2 主题 */
export const MASTER_RELS = [
    { relId: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', target: '../slideLayouts/slideLayout1.xml' },
    { relId: 'rId2', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', target: '../theme/theme1.xml' }
];

/** 版式关系：rId1 母版 */
export const LAYOUT_RELS = [
    { relId: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', target: '../slideMasters/slideMaster1.xml' }
];

/** 关系类型常量 */
export const REL_TYPES = {
    slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
    theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
    viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
    tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
    video: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    audio: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
    comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    commentAuthors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
    diagramData: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
    diagramLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout',
    diagramColors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors',
    diagramQuickStyle: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle',
    diagramDrawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
    oleObject: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
    font: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
    fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
    notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
    handoutMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster',
    thumbnail: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
};
