/**
 * JSON → PPTX 序列化模块
 *
 * 提供两个核心 API：
 *
 * 1. jsonToPptx(presentation, options) —— 从演示文稿 JSON 树生成全新的 PPTX 文件。
 *    输入格式与 PPTXComposer.toJSON() 输出一致，也可直接传 Composer 实例。
 *
 * 2. editPptx(fileData) —— 加载已有 PPTX，进行幻灯片管理（删除/重排/追加）、
 *    元数据写回等编辑操作后重新保存。借鉴 nodejs-pptx 的 load/modify/save 模式。
 *
 * 生成的文件结构完全兼容 PowerPoint / WPS / Keynote，
 * 且可被本库的 pptxToJson/pptxToHtml 直接解析（round-trip）。
 *
 * @module serializer/json-to-pptx
 */

import JSZip from 'jszip';
import { toXmlDocument, escapeXml } from './xml-builder';
import {
    buildThemeXml, buildSlideMasterXml, buildSlideLayoutXml,
    buildPresPropsXml, buildViewPropsXml, buildTableStylesXml,
    buildPresentationXml, buildRelationshipsXml, buildContentTypesXml,
    buildCorePropsXml, buildAppPropsXml, buildRootRelsXml, buildCustomPropsXml,
    buildCommentsXml, buildCommentAuthorsXml,
    MASTER_RELS, LAYOUT_RELS, REL_TYPES
} from './templates';
import { createElementContext, buildSlideRoot, buildNotesSlide, type SerializerSlide, type SerializerComment, type SerializerPart, type SerializerRel } from './element-builders';
import { PPTXXmlUtils } from '../utils/xml';

/** 演示文稿 JSON 树 */
interface PresentationJson {
    slides: SerializerSlide[];
    slideSize?: { width: number; height: number };
    metadata?: Record<string, string>;
    [key: string]: unknown;
}

/** 带 toJSON 的对象（如 PPTXComposer 实例） */
interface JsonSerializable {
    toJSON: () => PresentationJson;
}

/** sldIdLst 条目 */
interface SldIdEntry {
    id?: string | number;
    relId: string;
}

/** JSZip 支持的输出类型 */
export type ZipOutputType = 'base64' | 'string' | 'text' | 'binarystring' | 'array' | 'uint8array' | 'arraybuffer' | 'blob' | 'nodebuffer';

/** 已解析的 Relationship（属性集合） */
type ParsedRel = Record<string, string>;

/**
 * 关系目标 → zip 内绝对路径（ppt/...）
 * @param {string} target - 关系目标（../xx 或 ppt/xx 或绝对 http 链接）
 * @returns {string} 绝对路径；外部链接原样返回
 */
function toAbsPath(target: string): string {
    if (!target) return target;
    if (/^https?:/i.test(target)) return target;
    if (target.startsWith('ppt/')) return target;
    return `ppt/${target.replace(/^(\.\.\/)+/, '').replace(/^\/+/, '')}`;
}

/**
 * 规整 __raw 附属部件路径：与已占用路径冲突时重命名，并同步改写 ctx.rels 中指向原路径的关系目标
 *
 * 必须在 slide 关系落盘前调用，否则重命名后的部件不会被关系指向。
 *
 * @param {Object} ctx - 元素构建上下文
 * @param {Set<string>} usedPaths - 已被占用的 zip 内路径（调用方需持续维护）
 * @returns {Array<{part: Object, path: string}>} 部件与最终落盘路径
 */
function dedupeRawParts(ctx: { parts: SerializerPart[]; rels: SerializerRel[] }, usedPaths: Set<string>) {
    const resolved: { part: SerializerPart; path: string }[] = [];
    for (const part of ctx.parts) {
        let path = part.path;
        if (usedPaths.has(path)) {
            const dot = path.lastIndexOf('.');
            const base = dot > 0 ? path.slice(0, dot) : path;
            const ext = dot > 0 ? path.slice(dot) : '';
            let n = 2;
            while (usedPaths.has(`${base}__${n}${ext}`)) n++;
            path = `${base}__${n}${ext}`;
            // 同步改写指向原路径的关系目标
            for (const rel of ctx.rels) {
                if (toAbsPath(rel.target) === part.path) {
                    rel.target = `../${path.replace(/^ppt\//, '')}`;
                }
            }
        }
        usedPaths.add(path);
        resolved.push({ part, path });
    }
    return resolved;
}

/**
 * 规范化演示文稿 JSON 输入
 * @param {Object|{toJSON: Function}} input - 演示文稿 JSON 或 Composer 实例
 * @returns {Object} 演示文稿 JSON 树
 */
function normalizePresentation(input: unknown): PresentationJson {
    let presentation: unknown = input;
    const serializable = presentation as Partial<JsonSerializable> | null;
    if (serializable && typeof serializable.toJSON === 'function') {
        presentation = serializable.toJSON();
    }
    const pres = presentation as PresentationJson | null | undefined;
    if (!pres || typeof pres !== 'object') {
        throw new Error('jsonToPptx: 输入必须为演示文稿 JSON 对象或 PPTXComposer 实例');
    }
    if (!Array.isArray(pres.slides) || pres.slides.length === 0) {
        throw new Error('jsonToPptx: 演示文稿至少需要一页幻灯片（slides 数组为空）');
    }
    return pres;
}

/**
 * 将演示文稿 JSON 序列化为 PPTX
 * @param {Object|{toJSON: Function}} presentation - 演示文稿 JSON 或 Composer 实例
 * @param {Object} [options] - 选项
 * @param {string} [options.outputType='uint8array'] - JSZip 输出类型
 *        （uint8array / arraybuffer / blob / nodebuffer / base64）
 * @returns {Promise<Uint8Array>} PPTX 文件二进制数据
 */
/** 列号转字母（1→A, 2→B, 27→AA） */
function colToLetter(col: number): string {
    let s = '';
    while (col > 0) {
        const m = (col - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        col = Math.floor((col - 1) / 26);
    }
    return s;
}

/** 构建嵌入图表的最小 xlsx 工作簿（WPS 兼容必需） */
async function buildChartXlsx(workbook: { headers: string[]; rows: (string | number)[][] }): Promise<Uint8Array> {
    const xlsxZip = new JSZip();
    // 收集共享字符串
    const strings: string[] = [];
    const strIndex = new Map<string, number>();
    const getStrIdx = (s: string): number => {
        if (!strIndex.has(s)) { strIndex.set(s, strings.length); strings.push(s); }
        return strIndex.get(s)!;
    };

    // 构建 sheet1.xml 行数据
    const allRows = [workbook.headers, ...workbook.rows];
    const sheetRows = allRows.map((row, rIdx) => {
        const cells = row.map((val, cIdx) => {
            const ref = `${colToLetter(cIdx + 1)}${rIdx + 1}`;
            if (typeof val === 'number') {
                return `<c r="${ref}" t="n"><v>${val}</v></c>`;
            }
            const s = String(val);
            if (s === '') return `<c r="${ref}"/>`;
            return `<c r="${ref}" t="s"><v>${getStrIdx(s)}</v></c>`;
        }).join('');
        return `<row r="${rIdx + 1}">${cells}</row>`;
    }).join('');

    const sstXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">` +
        strings.map(s => `<si><t xml:space="preserve">${escapeXml(s)}</t></si>`).join('') +
        `</sst>`;

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<sheetData>${sheetRows}</sheetData></worksheet>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
        `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
        `<borders count="1"><border/></borders>` +
        `<cellStyleXfs count="1"><xf/></cellStyleXfs>` +
        `<cellXfs count="1"><xf/></cellXfs></styleSheet>`;

    const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
        `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
        `</Types>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

    const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;

    xlsxZip.file('[Content_Types].xml', ctXml);
    xlsxZip.file('_rels/.rels', rootRels);
    xlsxZip.file('xl/workbook.xml', workbookXml);
    xlsxZip.file('xl/_rels/workbook.xml.rels', wbRels);
    xlsxZip.file('xl/worksheets/sheet1.xml', sheetXml);
    xlsxZip.file('xl/styles.xml', stylesXml);
    xlsxZip.file('xl/sharedStrings.xml', sstXml);

    return xlsxZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function jsonToPptx(presentation: unknown, options: { outputType?: ZipOutputType } = {}) {
    const pres = normalizePresentation(presentation);
    const slideSize = pres.slideSize || { width: 1280, height: 720 };
    const zip = new JSZip();

    // ===== 逐页构建 slide XML 与关系 =====
    const slideRefs = [];       // presentation.xml 中的引用 { relId, target }
    const allMediaExts = new Set();
    const allChartNames = [];   // 图表部件名（用于 Content-Types 覆盖）
    const allWorkbookNames = []; // 嵌入工作簿名（用于 Content-Types 覆盖）
    const allNotesSlides: number[] = [];  // 备注部件编号（用于 Content-Types 覆盖）
    const commentsSlideIndices: number[] = [];  // 含批注的幻灯片编号（用于 Content-Types 覆盖）
    const allDiagramIndices: number[] = [];  // 图示部件编号（用于 Content-Types 覆盖）
    const allRawParts: { path: string; contentType: string }[] = []; // __raw 回退部件（用于 Content-Types 覆盖）
    let presRelId = 2;         // rId1 为母版

    // 媒体/图表编号需跨页递增：否则各页都从 1 开始，后页会覆盖前页的 image1/chart1
    let mediaIndex = 0;
    let chartIndex = 0;
    let diagramIndex = 0;
    const usedRawParts = new Set<string>();

    // 批注作者收集（跨页去重，按出现顺序编号，供 commentsN.xml 的 authorId 引用）
    const commentAuthors = new Map<string, number>();
    let commentAuthorSeq = 0;
    for (const slide of pres.slides) {
        const slideComments = (slide as any).comments as SerializerComment[] | undefined;
        if (slideComments) {
            for (const c of slideComments) {
                const name = c.author || 'Author';
                if (!commentAuthors.has(name)) commentAuthors.set(name, commentAuthorSeq++);
            }
        }
    }
    const hasComments = commentAuthors.size > 0;

    // 各页引用到的表格样式 ID：写 ppt/tableStyles.xml 时为它们补等价定义，
    // 否则未知 GUID 会让表格在 WPS/PowerPoint 里退化成「无样式无网格」
    const tableStyleIds = new Set<string>();

    for (const i of pres.slides.keys()){
        const slideIndex = i + 1;
        const ctx = createElementContext({ startMediaIndex: mediaIndex, startChartIndex: chartIndex, startDiagramIndex: diagramIndex });
        const slideRoot = await buildSlideRoot(ctx, pres.slides[i]);
        ctx.tableStyleIds?.forEach((id) => tableStyleIds.add(id));

        zip.file(`ppt/slides/slide${slideIndex}.xml`, toXmlDocument(slideRoot));

        // slide 关系：rId1 版式 + 元素产生的图片/超链接/图表关系
        const slideRels = [
            { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
            ...ctx.rels
        ];

        // 演讲者备注 → 生成 notesSlide 部件及其关系（双向标准字段）
        const slideNotes = pres.slides[i] && pres.slides[i].notes;
        if (slideNotes) {
            const notesRelId = `rId${ctx.nextRelId++}`;
            zip.file(`ppt/notesSlides/notesSlide${slideIndex}.xml`, toXmlDocument(buildNotesSlide(slideNotes)));
            zip.file(`ppt/notesSlides/_rels/notesSlide${slideIndex}.xml.rels`,
                buildRelationshipsXml([{ relId: 'rId1', type: REL_TYPES.slide, target: `../slides/slide${slideIndex}.xml` }]));
            slideRels.push({ relId: notesRelId, type: REL_TYPES.notesSlide, target: `../notesSlides/notesSlide${slideIndex}.xml` });
            allNotesSlides.push(slideIndex);
        }

        // 幻灯片批注 → 生成 commentsN.xml 部件及其关系（幻灯片 → commentsN.xml → commentAuthors.xml）
        const slideComments = (pres.slides[i] as any).comments as SerializerComment[] | undefined;
        if (slideComments && slideComments.length) {
            const commentsRelId = `rId${ctx.nextRelId++}`;
            const resolved = slideComments.map((c) => ({
                authorId: commentAuthors.get(c.author || 'Author') ?? 0,
                text: c.text,
                dt: c.dt || new Date().toISOString(),
                x: c.pos?.x,
                y: c.pos?.y
            }));
            zip.file(`ppt/comments/comments${slideIndex}.xml`, buildCommentsXml(resolved));
            zip.file(`ppt/comments/_rels/comments${slideIndex}.xml.rels`,
                buildRelationshipsXml([{ relId: 'rId1', type: REL_TYPES.commentAuthors, target: '../commentAuthors.xml' }]));
            slideRels.push({ relId: commentsRelId, type: REL_TYPES.comments, target: `../comments/comments${slideIndex}.xml` });
            commentsSlideIndices.push(slideIndex);
        }

        // __raw 回退附属部件（SmartArt 的 diagrams/*.xml 等）
        // 不同页的原始部件可能同名（如各自的 diagrams/data1.xml），需去重并同步改写引用关系。
        // 必须先于 slide rels 落盘执行，否则重命名的部件不会被关系指向。
        for (const { part, path } of dedupeRawParts(ctx, usedRawParts)) {
            if (part.base64) {
                zip.file(path, part.base64, { base64: true });
                allMediaExts.add((path.split('.').pop() ?? '').toLowerCase());
            } else if (part.content !== undefined) {
                zip.file(path, part.content);
                allRawParts.push({ path, contentType: part.contentType });
            }
        }

        zip.file(`ppt/slides/_rels/slide${slideIndex}.xml.rels`, buildRelationshipsXml(slideRels));

        // 媒体文件
        for (const media of ctx.media) {
            zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
            allMediaExts.add((media.name.split('.').pop() ?? '').toLowerCase());
        }

        // 图表部件（含嵌入工作簿 + chart rels，WPS 兼容必需）
        for (const chart of ctx.charts) {
            zip.file(`ppt/charts/${chart.name}`, chart.xml);
            allChartNames.push(chart.name);
            // 生成嵌入 xlsx 工作簿
            if (chart.workbook) {
                const wbNum = allChartNames.length; // workbook 编号与 chart 编号一致
                const wbName = `workbook${wbNum}.xlsx`;
                const xlsxData = await buildChartXlsx(chart.workbook);
                zip.file(`ppt/embeddings/${wbName}`, xlsxData);
                allWorkbookNames.push(wbName);
                // chart rels：oleObject 关系指向嵌入工作簿
                zip.file(`ppt/charts/_rels/${chart.name}.rels`,
                    buildRelationshipsXml([
                        { relId: 'rId1', type: REL_TYPES.oleObject, target: `../embeddings/${wbName}` }
                    ]));
            }
        }

        // 图示部件（data/layout/colors/quickStyle 四件套 + dataN.xml.rels）
        for (const d of ctx.diagrams) {
            const dir = `ppt/diagrams/`;
            zip.file(`${dir}data${d.index}.xml`, d.dataXml);
            zip.file(`${dir}layout${d.index}.xml`, d.layoutXml);
            zip.file(`${dir}colors${d.index}.xml`, d.colorsXml);
            zip.file(`${dir}quickStyle${d.index}.xml`, d.quickStyleXml);
            zip.file(`${dir}_rels/data${d.index}.xml.rels`,
                buildRelationshipsXml([
                    { relId: 'rId1', type: REL_TYPES.diagramLayout, target: `layout${d.index}.xml` },
                    { relId: 'rId2', type: REL_TYPES.diagramColors, target: `colors${d.index}.xml` },
                    { relId: 'rId3', type: REL_TYPES.diagramQuickStyle, target: `quickStyle${d.index}.xml` }
                ]));
            allDiagramIndices.push(d.index);
        }

        mediaIndex = ctx.mediaIndex;
        chartIndex = ctx.chartIndex;
        diagramIndex = ctx.diagramIndex;

        slideRefs.push({ relId: `rId${presRelId++}`, target: `slides/slide${slideIndex}.xml` });
    }

    // 批注作者表（所有含批注的页共享一份 commentAuthors.xml，由 comments 部件与 presentation 共同引用）
    if (hasComments) {
        const authors = [...commentAuthors.entries()]
            .sort((a, b) => a[1] - b[1])
            .map(([name, id]) => ({ id, name }));
        zip.file('ppt/commentAuthors.xml', buildCommentAuthorsXml(authors));
    }

    // ===== presentation.xml 及其关系 =====
    zip.file('ppt/presentation.xml', buildPresentationXml(slideSize, slideRefs));

    // presentation.xml.rels：母版 rId1 → 各 slide → theme/presProps/viewProps/tableStyles
    const presRels = [
        { relId: 'rId1', type: REL_TYPES.slideMaster, target: 'slideMasters/slideMaster1.xml' },
        ...slideRefs.map(ref => ({ relId: ref.relId, type: REL_TYPES.slide, target: ref.target }))
    ];
    presRels.push(
        { relId: `rId${presRelId++}`, type: REL_TYPES.theme, target: 'theme/theme1.xml' },
        { relId: `rId${presRelId++}`, type: REL_TYPES.presProps, target: 'presProps.xml' },
        { relId: `rId${presRelId++}`, type: REL_TYPES.viewProps, target: 'viewProps.xml' },
        { relId: `rId${presRelId++}`, type: REL_TYPES.tableStyles, target: 'tableStyles.xml' }
    );
    if (hasComments) {
        presRels.push({ relId: `rId${presRelId++}`, type: REL_TYPES.commentAuthors, target: 'commentAuthors.xml' });
    }
    zip.file('ppt/_rels/presentation.xml.rels', buildRelationshipsXml(presRels));

    // ===== 静态部件 =====
    // 主题 / 母版 / 版式：支持自定义覆盖（options.theme/masterXml/layoutXml 或 pres.theme/slideMaster/slideLayout 提供完整 XML）
    const themeXml = (options && (options as any).theme) || (pres as any).theme;
    zip.file('ppt/theme/theme1.xml', typeof themeXml === 'string' ? themeXml : buildThemeXml());
    const masterXml = (options && (options as any).masterXml) || (pres as any).slideMaster;
    zip.file('ppt/slideMasters/slideMaster1.xml', typeof masterXml === 'string' ? masterXml : buildSlideMasterXml());
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', buildRelationshipsXml(MASTER_RELS));
    const layoutXml = (options && (options as any).layoutXml) || (pres as any).slideLayout;
    zip.file('ppt/slideLayouts/slideLayout1.xml', typeof layoutXml === 'string' ? layoutXml : buildSlideLayoutXml());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', buildRelationshipsXml(LAYOUT_RELS));
    zip.file('ppt/presProps.xml', buildPresPropsXml());
    zip.file('ppt/viewProps.xml', buildViewPropsXml());
    zip.file('ppt/tableStyles.xml', buildTableStylesXml([...tableStyleIds]));

    // ===== docProps =====
    zip.file('docProps/core.xml', buildCorePropsXml(pres.metadata));
    zip.file('docProps/app.xml', buildAppPropsXml(pres.slides.length));

    // ===== 包级文件 =====
    let rootRelsXml = buildRootRelsXml();
    if ((pres as any).customProps) {
        rootRelsXml = rootRelsXml.replace('</Relationships>',
            `<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`);
    }
    zip.file('_rels/.rels', rootRelsXml);
    let contentTypeXml = buildContentTypesXml([...allMediaExts], pres.slides.length);
    if ((pres as any).customProps) {
        contentTypeXml = contentTypeXml.replace('</Types>',
            `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>`);
        zip.file('docProps/custom.xml', buildCustomPropsXml((pres as any).customProps as Record<string, string>));
    }
    // 图表部件 Content-Types 覆盖
    for (const chartName of allChartNames) {
        const override = `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
    }
    // 嵌入工作簿 Content-Types 覆盖（WPS 兼容必需）
    for (const wbName of allWorkbookNames) {
        const override = `<Override PartName="/ppt/embeddings/${wbName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
    }
    // 备注部件 Content-Types 覆盖
    for (const idx of allNotesSlides) {
        const override = `<Override PartName="/ppt/notesSlides/notesSlide${idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
    }
    // 批注部件 Content-Types 覆盖（commentsN.xml + commentAuthors.xml）
    if (hasComments) {
        const cmOverride = `<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${cmOverride}</Types>`);
        for (const idx of commentsSlideIndices) {
            const override = `<Override PartName="/ppt/comments/comments${idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>`;
            contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
        }
    }
    // 图示部件 Content-Types 覆盖（data/layout/colors/quickStyle）
    const DIAGRAM_CT: Record<string, string> = {
        data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
        layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
        colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
        quickStyle: 'application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml'
    };
    for (const idx of allDiagramIndices) {
        for (const kind of ['data', 'layout', 'colors', 'quickStyle'] as const) {
            const override = `<Override PartName="/ppt/diagrams/${kind}${idx}.xml" ContentType="${DIAGRAM_CT[kind]}"/>`;
            contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
        }
    }
    // __raw 回退部件 Content-Types 覆盖
    for (const part of allRawParts) {
        const override = `<Override PartName="/${part.path}" ContentType="${escapeXml(part.contentType)}"/>`;
        contentTypeXml = contentTypeXml.replace('</Types>', `${override}</Types>`);
    }
    zip.file('[Content_Types].xml', contentTypeXml);

    return zip.generateAsync({
        type: options.outputType || 'uint8array',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    });
}

// ---------------------------------------------------------------------------
// 已有 PPTX 编辑器（editPptx）
// ---------------------------------------------------------------------------

/** 从文本中解析 <Relationship .../> 列表 */
function parseRelationships(relsText: string): ParsedRel[] {
    const rels: ParsedRel[] = [];
    const re = /<Relationship\s+([^>]*?)\/>/g;
    const attrRe = /(\w+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(relsText)) !== null) {
        const attrs: ParsedRel = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs.Id) rels.push(attrs);
    }
    return rels;
}

/** 从 presentation.xml 解析 sldIdLst 条目 */
function parseSldIdLst(presentationText: string): SldIdEntry[] {
    const lstMatch = presentationText.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!lstMatch) return [];
    const entries: SldIdEntry[] = [];
    const re = /<p:sldId\s+([^>]*?)\/>/g;
    const attrRe = /([\w:.-]+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(lstMatch[1])) !== null) {
        const attrs: ParsedRel = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs['r:id']) entries.push({ id: attrs.id, relId: attrs['r:id'] });
    }
    return entries;
}

/** 用新条目列表重写 presentation.xml 的 sldIdLst */
function rewriteSldIdLst(presentationText: string, entries: SldIdEntry[]) {
    const inner = entries
        .map((e, i) => `<p:sldId id="${e.id || 256 + i}" r:id="${escapeXml(e.relId)}"/>`)
        .join('');
    return presentationText.replace(
        /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
        `<p:sldIdLst>${inner}</p:sldIdLst>`
    );
}

/** 列出 zip 中匹配 /ppt\/slides\/slide(\d+)\.xml 的幻灯片编号 */
function listSlideNumbers(zip: JSZip) {
    const numbers: number[] = [];
    zip.forEach((path: string, entry: { dir: boolean }) => {
        if (entry.dir) return;
        const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (m) numbers.push(Number(m[1]));
    });
    return numbers.sort((a: number, b: number) => a - b);
}

/** 列出 zip 中 ppt/charts 下的最大图表编号（chartN.xml） */
function maxChartIndex(zip: JSZip) {
    let max = 0;
    zip.forEach((path: string, entry: { dir: boolean }) => {
        if (entry.dir) return;
        const m = path.match(/^ppt\/charts\/[^/]*?(\d+)\.[^/]+$/);
        if (m) max = Math.max(max, Number(m[1]));
    });
    return max;
}

/** 列出 zip 中 ppt/media 下的最大媒体编号（不限 imageN 命名，取文件名中最后一个数字段） */
function maxMediaIndex(zip: JSZip) {
    let max = 0;
    zip.forEach((path: string, entry: { dir: boolean }) => {
        if (entry.dir) return;
        const m = path.match(/^ppt\/media\/[^/]*?(\d+)\.[^/]+$/);
        if (m) max = Math.max(max, Number(m[1]));
    });
    return max;
}

/**
 * 加载已有 PPTX 并返回编辑器
 * @param {ArrayBuffer|Uint8Array|Buffer} fileData - PPTX 文件数据
 * @returns {Promise<Object>} 编辑器实例
 */
async function editPptx(fileData: ArrayBuffer | Uint8Array | string) {
    const zip = await JSZip.loadAsync(fileData);

    const readText = async (name: string) => {
        const f = zip.file(name);
        return f ? await f.async('string') : null;
    };

    /** 读取 presentation.xml 解析后的信息 */
    async function getPresentationInfo() {
        const text = await readText('ppt/presentation.xml');
        if (!text) throw new Error('editPptx: 无效的 PPTX 文件（缺少 ppt/presentation.xml）');
        const relsText = await readText('ppt/_rels/presentation.xml.rels');
        const rels = parseRelationships(relsText || '');
        const relById: Record<string, ParsedRel> = {};
        for (const r of rels) relById[r.Id] = r;
        const entries = parseSldIdLst(text);
        // 解析每页对应的 slide 部件路径
        const slideParts = entries.map(e => {
            const rel = relById[e.relId];
            if (!rel) return null;
            const target = rel.Target.replace(/^\//, '').replace(/^ppt\//, 'ppt/');
            return { relId: e.relId, target: target.startsWith('ppt/') ? target : `ppt/${target}` };
        }).filter(Boolean);
        return { text, rels, relById, entries, slideParts };
    }

    /** 回写 docProps/app.xml 的 <Slides> 计数（不存在则跳过） */
    async function setAppSlideCount(n: number) {
        const app = await readText('docProps/app.xml');
        if (app && /<Slides>\d+<\/Slides>/.test(app)) {
            zip.file('docProps/app.xml', app.replace(/<Slides>\d+<\/Slides>/, `<Slides>${n}</Slides>`));
        }
    }

    /**
     * 保存编辑结果
     * @param {Object} [options] - 选项
     * @param {string} [options.outputType='uint8array'] - JSZip 输出类型
     * @returns {Promise<Uint8Array>} 新的 PPTX 文件数据
     */
    async function save(options: { outputType?: ZipOutputType } = {}) {
        return zip.generateAsync({
            type: options.outputType || 'uint8array',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
    }

    return {
        zip,
        save,

        /**
         * 获取幻灯片数量（逻辑顺序 = sldIdLst 顺序，与 PowerPoint 显示一致）
         * @returns {Promise<number>} 页数
         */
        async getSlideCount() {
            const info = await getPresentationInfo();
            return info.entries.length;
        },

        /**
         * 获取指定页的简化 XML 树（与 pptxToJson 结果中的 slideContent 同构，仅含该页）。
         * slideNum 为逻辑顺序（sldIdLst 顺序），通过关系反查真实文件名。
         * @param {number} slideNum - 页码（1 起，逻辑顺序）
         * @returns {Promise<Object|null>} tXml 简化树
         */
        async getSlide(slideNum: number) {
            const info = await getPresentationInfo();
            const part = info.slideParts[slideNum - 1];
            if (!part) throw new Error(`getSlide: 页码越界（共 ${info.entries.length} 页）`);
            return PPTXXmlUtils.readXmlFile(zip, part.target);
        },

        /**
         * 删除指定页（同步移除 sldIdLst 条目、关系、slide 部件、Content-Types 覆盖项；
         * 关联的 notesSlide 会一并移除，媒体文件保留不清理）
         * @param {number} slideNum - 页码（1 起）
         */
        async deleteSlide(slideNum: number) {
            const info = await getPresentationInfo();
            if (slideNum < 1 || slideNum > info.entries.length) {
                throw new Error(`deleteSlide: 页码越界（共 ${info.entries.length} 页）`);
            }
            if (info.entries.length <= 1) {
                throw new Error('deleteSlide: 至少需保留一页幻灯片，无法删除最后一页');
            }

            const part = info.slideParts[slideNum - 1];
            const slidePath = part!.target; // ppt/slides/slideN.xml

            // 1. 移除 sldIdLst 条目
            const remaining = info.entries.filter((_, i) => i !== slideNum - 1);
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, remaining));

            // 2. 移除 presentation.xml.rels 中的关系
            const relsText = await readText('ppt/_rels/presentation.xml.rels');
            const relRe = new RegExp(`<Relationship\\s+Id="${part!.relId}"[^>]*/>`);
            zip.file('ppt/_rels/presentation.xml.rels', relsText!.replace(relRe, ''));

            // 3. 找到并移除 notesSlide（通过 slide 的 rels）
            const slideRelsText = await readText(`${slidePath.replace('slides/', 'slides/_rels/')}.rels`);
            if (slideRelsText) {
                for (const rel of parseRelationships(slideRelsText)) {
                    if (rel.Type && rel.Type.endsWith('/notesSlide')) {
                        const notesPath = rel.Target.replace('../', 'ppt/');
                        zip.remove(notesPath);
                        zip.remove(`${notesPath.replace('notesSlides/', 'notesSlides/_rels/')}.rels`);
                        await removeContentTypeOverride(notesPath);
                    }
                }
            }

            // 4. 移除 slide 部件与关系文件
            zip.remove(slidePath);
            zip.remove(`${slidePath.replace('slides/', 'slides/_rels/')}.rels`);

            // 5. 移除 Content-Types 覆盖项
            await removeContentTypeOverride(slidePath);

            // 6. 回写 app.xml 幻灯片计数
            await setAppSlideCount(remaining.length);
        },

        /**
         * 重排幻灯片：把第 from 页移动到第 to 页的位置。
         * 除更新 sldIdLst 外，还会物理重编号 slide 文件（slide1..slideN 按新顺序），
         * 以保证 PowerPoint 与本库解析端（按文件名编号排序）顺序一致；
         * 同时重映射各页 rels 中的内部跳转目标。
         * @param {number} from - 原页码（1 起）
         * @param {number} to - 目标页码（1 起）
         */
        async moveSlide(from: number, to: number) {
            const info = await getPresentationInfo();
            const entries = [...info.entries];
            if (from < 1 || from > entries.length || to < 1 || to > entries.length) {
                throw new Error(`moveSlide: 页码越界（共 ${entries.length} 页）`);
            }
            const [moved] = entries.splice(from - 1, 1);
            entries.splice(to - 1, 0, moved);

            // 新逻辑顺序下各条目对应的旧文件编号
            const orderedRels = entries.map(e => info.relById[e.relId]);
            const oldNums = orderedRels.map(rel => {
                const m = rel.Target.match(/slide(\d+)\.xml$/);
                if (!m) throw new Error(`moveSlide: 无法解析 slide 目标 ${rel.Target}`);
                return Number(m[1]);
            });
            const mapping: Record<string, number> = {}; // 旧编号 → 新编号
            oldNums.forEach((oldNum, i) => { mapping[String(oldNum)] = i + 1; });

            // 先读出全部内容，避免读写交叉覆盖
            const contents: (string | null)[] = [];
            const relsContents: (string | null)[] = [];
            for (const oldNum of oldNums) {
                contents.push(await readText(`ppt/slides/slide${oldNum}.xml`));
                relsContents.push(await readText(`ppt/slides/_rels/slide${oldNum}.xml.rels`));
            }

            // 按新编号写回，并重映射 rels 中的内部跳转目标（slideN.xml）
            oldNums.forEach((oldNum, i) => {
                const newNum = i + 1;
                const newRels = (relsContents[i] || '').replace(
                    /Target="slide(\d+)\.xml"/g,
                    (m: string, n: string) => `Target="slide${mapping[n] !== undefined ? mapping[n] : n}.xml"`
                );
                zip.file(`ppt/slides/slide${newNum}.xml`, contents[i]!);
                zip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, newRels);
            });

            // 更新 presentation.xml.rels：各 relId 目标改为新编号文件
            let relsText = await readText('ppt/_rels/presentation.xml.rels');
            orderedRels.forEach((rel, i) => {
                const oldNum = oldNums[i];
                const newNum = i + 1;
                if (oldNum !== newNum) {
                    const re = new RegExp(`(<Relationship\\s+Id="${rel.Id}"[^>]*?Target=")slides/slide${oldNum}\\.xml(")`);
                    relsText = relsText!.replace(re, `$1slides/slide${newNum}.xml$2`);
                }
            });
            zip.file('ppt/_rels/presentation.xml.rels', relsText!);

            // 重写 sldIdLst（新逻辑顺序）
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, entries));
        },

        /**
         * 写回演示文稿元数据（docProps/core.xml）
         * @param {Object} metadata - 元数据（字段同 pptxToJson 返回的 metadata）
         */
        async setMetadata(metadata: Record<string, unknown>) {
            const existing = await readText('docProps/core.xml');
            if (existing) {
                zip.file('docProps/core.xml', buildCorePropsXml(metadata));
            } else {
                // 创建 core.xml 并补齐根关系与 Content-Types
                zip.file('docProps/core.xml', buildCorePropsXml(metadata));
                const rootRels = await readText('_rels/.rels');
                if (rootRels && !rootRels.includes('core-properties')) {
                    const newRel = '<Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
                    zip.file('_rels/.rels', rootRels.replace('</Relationships>', `${newRel}</Relationships>`));
                }
                const ctText = await readText('[Content_Types].xml');
                if (ctText && !ctText.includes('core-properties')) {
                    const override = '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
                    zip.file('[Content_Types].xml', ctText.replace('</Types>', `${override}</Types>`));
                }
            }
        },

        /**
         * 追加一页幻灯片（元素格式同 jsonToPptx 的 slide JSON）
         * @param {Object} slideJson - 幻灯片 JSON { background, elements }
         */
        async addSlide(slideJson: SerializerSlide) {
            const info = await getPresentationInfo();
            const numbers = listSlideNumbers(zip);
            const nextNum = (numbers.length ? numbers[numbers.length - 1] : 0) + 1;

            // 构建新页内容（媒体/图表编号从包内现有最大值续接，避免覆盖已有部件）
            const ctx = createElementContext({
                startMediaIndex: maxMediaIndex(zip),
                startChartIndex: maxChartIndex(zip)
            });
            const slideRoot = await buildSlideRoot(ctx, slideJson);
            zip.file(`ppt/slides/slide${nextNum}.xml`, toXmlDocument(slideRoot));

            // 媒体文件
            const mediaExts = new Set();
            const newCtCharts = [];
            for (const media of ctx.media) {
                zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
                mediaExts.add((media.name.split('.').pop() ?? '').toLowerCase());
            }

            // 图表部件（含嵌入工作簿 + chart rels）
            const newCtWorkbooks = [];
            for (const chart of ctx.charts) {
                zip.file(`ppt/charts/${chart.name}`, chart.xml);
                newCtCharts.push(chart.name);
                if (chart.workbook) {
                    const wbNum = newCtCharts.length;
                    const wbName = `workbook${wbNum}.xlsx`;
                    const xlsxData = await buildChartXlsx(chart.workbook);
                    zip.file(`ppt/embeddings/${wbName}`, xlsxData);
                    newCtWorkbooks.push(wbName);
                    zip.file(`ppt/charts/_rels/${chart.name}.rels`,
                        buildRelationshipsXml([
                            { relId: 'rId1', type: REL_TYPES.oleObject, target: `../embeddings/${wbName}` }
                        ]));
                }
            }

            // 图示部件（data/layout/colors/quickStyle 四件套 + dataN.xml.rels）
            const newCtDiagrams: number[] = [];
            for (const d of ctx.diagrams) {
                const dir = `ppt/diagrams/`;
                zip.file(`${dir}data${d.index}.xml`, d.dataXml);
                zip.file(`${dir}layout${d.index}.xml`, d.layoutXml);
                zip.file(`${dir}colors${d.index}.xml`, d.colorsXml);
                zip.file(`${dir}quickStyle${d.index}.xml`, d.quickStyleXml);
                zip.file(`${dir}_rels/data${d.index}.xml.rels`,
                    buildRelationshipsXml([
                        { relId: 'rId1', type: REL_TYPES.diagramLayout, target: `layout${d.index}.xml` },
                        { relId: 'rId2', type: REL_TYPES.diagramColors, target: `colors${d.index}.xml` },
                        { relId: 'rId3', type: REL_TYPES.diagramQuickStyle, target: `quickStyle${d.index}.xml` }
                    ]));
                newCtDiagrams.push(d.index);
            }

            // __raw 回退附属部件（SmartArt 的 diagrams/*.xml 等）
            // 与包内已有部件重名时重命名并同步改写关系，且必须先于 slide rels 落盘
            const usedPaths = new Set<string>();
            zip.forEach((p: string, entry: { dir: boolean }) => { if (!entry.dir) usedPaths.add(p); });
            const newCtRawParts: { path: string; contentType: string }[] = [];
            for (const { part, path } of dedupeRawParts(ctx, usedPaths)) {
                if (part.base64) {
                    zip.file(path, part.base64, { base64: true });
                    mediaExts.add((path.split('.').pop() ?? '').toLowerCase());
                } else if (part.content !== undefined) {
                    zip.file(path, part.content);
                    newCtRawParts.push({ path, contentType: part.contentType });
                }
            }

            const slideRels = [
                { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
                ...ctx.rels
            ];
            zip.file(`ppt/slides/_rels/slide${nextNum}.xml.rels`, buildRelationshipsXml(slideRels));

            // presentation.xml 追加 sldId
            const maxId = info.entries.reduce((m, e) => Math.max(m, Number(e.id) || 256), 255);
            const usedRelIds = new Set(info.rels.map(r => r.Id));
            let relNum = 1;
            while (usedRelIds.has(`rId${relNum}`)) relNum++;
            const newRelId = `rId${relNum}`;

            const newEntries = [...info.entries, { id: maxId + 1, relId: newRelId }];
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, newEntries));

            // 回写 app.xml 幻灯片计数
            await setAppSlideCount(newEntries.length);

            // presentation.xml.rels 追加关系
            const relsText = await readText('ppt/_rels/presentation.xml.rels');
            const newRel = `<Relationship Id="${newRelId}" Type="${REL_TYPES.slide}" Target="slides/slide${nextNum}.xml"/>`;
            zip.file('ppt/_rels/presentation.xml.rels', relsText!.replace('</Relationships>', `${newRel}</Relationships>`));

            // Content-Types 追加
            const ctText = await readText('[Content_Types].xml');
            let newCt = ctText;
            const override = `<Override PartName="/ppt/slides/slide${nextNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            newCt = newCt!.replace('</Types>', `${override}</Types>`);
            const MIME_MAP: Record<string, string> = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' };
            for (const ext of mediaExts) {
                if (!newCt.includes(`Extension="${ext}"`)) {
                    newCt = newCt.replace('</Types>', `<Default Extension="${ext}" ContentType="${MIME_MAP[String(ext).toLowerCase()] || 'application/octet-stream'}"/></Types>`);
                }
            }
            for (const chartName of newCtCharts) {
                if (!newCt.includes(`/ppt/charts/${chartName}`)) {
                    newCt = newCt.replace('</Types>', `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
                }
            }
            for (const wbName of newCtWorkbooks) {
                if (!newCt.includes(`/ppt/embeddings/${wbName}`)) {
                    newCt = newCt.replace('</Types>', `<Override PartName="/ppt/embeddings/${wbName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>`);
                }
            }
            for (const idx of newCtDiagrams) {
                for (const kind of ['data', 'layout', 'colors', 'quickStyle'] as const) {
                    const partName = `/ppt/diagrams/${kind}${idx}.xml`;
                    if (!newCt.includes(partName)) {
                        const ct = { data: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml', layout: 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml', colors: 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml', quickStyle: 'application/vnd.openxmlformats-officedocument.drawingml.diagramQuickStyle+xml' }[kind];
                        newCt = newCt.replace('</Types>', `<Override PartName="${partName}" ContentType="${ct}"/></Types>`);
                    }
                }
            }
            for (const part of newCtRawParts) {
                if (!newCt.includes(`/${part.path}`)) {
                    newCt = newCt.replace('</Types>', `<Override PartName="/${part.path}" ContentType="${escapeXml(part.contentType)}"/></Types>`);
                }
            }
            zip.file('[Content_Types].xml', newCt);
        }
    };

    /**
     * 移除 [Content_Types].xml 中指定部件的覆盖项
     * @param {string} partPath - 部件路径（如 ppt/slides/slide2.xml）
     */
    async function removeContentTypeOverride(partPath: string) {
        const ctText = await readText('[Content_Types].xml');
        if (!ctText) return;
        const re = new RegExp(`<Override PartName="/${partPath.replace(/\//g, '\\/')}"[^>]*/>`);
        zip.file('[Content_Types].xml', ctText.replace(re, ''));
    }
}

export { jsonToPptx, editPptx };
