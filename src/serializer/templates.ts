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

import { NS, escapeXml, pxToEmu } from './xml-builder';

/** DrawingML 常用命名空间声明串 */
const DRAWING_NS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"`;

/**
 * 主题模板（Office 标准配色/字体方案，最小化 fmtScheme）
 * @returns {string} theme1.xml 内容
 */
export function buildThemeXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${NS.a}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="110000"/></a:schemeClr></a:gs><a:gs pos="1000"><a:schemeClr val="phClr"><a:tint val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="4553000" scaled="0"/></a:gradFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="97000"/><a:satMod val="130000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

/**
 * 母版模板（空白母版，仅含布局引用与颜色映射）
 * @returns {string} slideMaster1.xml 内容
 */
export function buildSlideMasterXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${DRAWING_NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}

/**
 * 版式模板（空白版式）
 * @returns {string} slideLayout1.xml 内容
 */
export function buildSlideLayoutXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${DRAWING_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
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
 * 表格样式模板（tableStyles.xml）
 *
 * 根元素必须是 DrawingML 命名空间下的 a:tblStyleLst（解析端按 ["a:tblStyleLst"]["a:tblStyle"] 读取）；
 * 并给出默认样式的完整定义（纯网格线），使表格在解析端也能渲染出完整网格而非只剩单元格自定义边。
 * @returns {string} tableStyles.xml 内容
 */
export function buildTableStylesXml() {
    const edge = (side: string) =>
        `<a:${side}><a:ln w="12700" cmpd="sng"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></a:${side}>`;
    const grid = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map(edge).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<a:tblStyleLst xmlns:a="${NS.a}" def="${DEFAULT_TABLE_STYLE_ID}">` +
        `<a:tblStyle styleId="${DEFAULT_TABLE_STYLE_ID}" styleName="Table Grid">` +
        `<a:wholeTbl>` +
        `<a:tcTxStyle b="off"><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr></a:tcStyle>` +
        `</a:wholeTbl>` +
        `<a:firstRow>` +
        `<a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle>` +
        `<a:tcStyle><a:tcBdr>${grid}</a:tcBdr></a:tcStyle>` +
        `</a:firstRow>` +
        `</a:tblStyle>` +
        `</a:tblStyleLst>`;
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
export function buildPresentationXml(slideSize: { width: number; height: number }, slides: Array<{ relId: string }>) {
    const slideEntries = slides
        .map((s, i) => `<p:sldId id="${256 + i}" r:id="${escapeXml(s.relId)}"/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slideEntries}</p:sldIdLst><p:sldSz cx="${pxToEmu(slideSize.width)}" cy="${pxToEmu(slideSize.height)}"/><p:notesSz cx="6858000" cy="914400"/><p:defaultTextStyle/></p:presentation>`;
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
export function buildContentTypesXml(mediaExts: Iterable<unknown>, slideCount: number) {
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
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${mediaDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>${slideOverrides}</Types>`;
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
    diagramQuickStyle: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle'
};
