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

import { NS, escapeXml, pxToEmu } from './xml-builder.js';

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
 * 表格样式模板（tableStyles.xml）
 * @returns {string} tableStyles.xml 内容
 */
export function buildTableStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:tblStyLst xmlns:a="${NS.a}" xmlns:p="${NS.p}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

/**
 * 生成 presentation.xml
 * @param {Object} slideSize - 幻灯片尺寸（px）
 * @param {number} slideSize.width - 宽（px）
 * @param {number} slideSize.height - 高（px）
 * @param {Array<{relId: string}>} slides - 幻灯片引用列表
 * @returns {string} presentation.xml 内容
 */
export function buildPresentationXml(slideSize, slides) {
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
export function buildRelationshipsXml(rels) {
    const entries = rels
        .map(r => `<Relationship Id="${escapeXml(r.relId)}" Type="${escapeXml(r.type)}" Target="${escapeXml(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`)
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${NS.rel}">${entries}</Relationships>`;
}

/**
 * 生成 [Content_Types].xml
 * @param {Array<string>} mediaExts - 用到的媒体扩展名（如 ['png','jpeg']）
 * @param {number} slideCount - 幻灯片数量
 * @returns {string} [Content_Types].xml 内容
 */
export function buildContentTypesXml(mediaExts, slideCount) {
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
        .map(ext => `<Default Extension="${escapeXml(ext)}" ContentType="${MIME_MAP[ext.toLowerCase()] || 'application/octet-stream'}"/>`)
        .join('');
    const slideOverrides = Array.from({ length: slideCount }, (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${mediaDefaults}<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${slideOverrides}</Types>`;
}

/**
 * 生成 docProps/core.xml（元数据）
 * @param {Object} metadata - 元数据（title/author/subject/keywords/description/lastModifiedBy/category/status）
 * @returns {string} core.xml 内容
 */
export function buildCorePropsXml(metadata) {
    const md = metadata || {};
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
export function buildAppPropsXml(slideCount) {
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
    presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
    viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
    tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
};
