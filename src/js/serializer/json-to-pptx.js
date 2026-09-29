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
import { toXmlDocument, escapeXml } from './xml-builder.js';
import {
    buildThemeXml, buildSlideMasterXml, buildSlideLayoutXml,
    buildPresPropsXml, buildViewPropsXml, buildTableStylesXml,
    buildPresentationXml, buildRelationshipsXml, buildContentTypesXml,
    buildCorePropsXml, buildAppPropsXml, buildRootRelsXml,
    MASTER_RELS, LAYOUT_RELS, REL_TYPES
} from './templates.js';
import { createElementContext, buildSlideRoot } from './element-builders.js';
import { PPTXXmlUtils } from '../utils/xml.js';

/**
 * 规范化演示文稿 JSON 输入
 * @param {Object|{toJSON: Function}} input - 演示文稿 JSON 或 Composer 实例
 * @returns {Object} 演示文稿 JSON 树
 */
function normalizePresentation(input) {
    let presentation = input;
    if (presentation && typeof presentation.toJSON === 'function') {
        presentation = presentation.toJSON();
    }
    if (!presentation || typeof presentation !== 'object') {
        throw new Error('jsonToPptx: 输入必须为演示文稿 JSON 对象或 PPTXComposer 实例');
    }
    if (!Array.isArray(presentation.slides) || presentation.slides.length === 0) {
        throw new Error('jsonToPptx: 演示文稿至少需要一页幻灯片（slides 数组为空）');
    }
    return presentation;
}

/**
 * 将演示文稿 JSON 序列化为 PPTX
 * @param {Object|{toJSON: Function}} presentation - 演示文稿 JSON 或 Composer 实例
 * @param {Object} [options] - 选项
 * @param {string} [options.outputType='uint8array'] - JSZip 输出类型
 *        （uint8array / arraybuffer / blob / nodebuffer / base64）
 * @returns {Promise<Uint8Array>} PPTX 文件二进制数据
 */
async function jsonToPptx(presentation, options = {}) {
    const pres = normalizePresentation(presentation);
    const slideSize = pres.slideSize || { width: 1280, height: 720 };
    const zip = new JSZip();

    // ===== 逐页构建 slide XML 与关系 =====
    const slideRefs = [];       // presentation.xml 中的引用 { relId, target }
    const allMediaExts = new Set();
    const allChartNames = [];   // 图表部件名（用于 Content-Types 覆盖）
    let presRelId = 2;         // rId1 为母版

    for (let i = 0; i < pres.slides.length; i++) {
        const slideIndex = i + 1;
        const ctx = createElementContext();
        const slideRoot = await buildSlideRoot(ctx, pres.slides[i]);

        zip.file(`ppt/slides/slide${slideIndex}.xml`, toXmlDocument(slideRoot));

        // slide 关系：rId1 版式 + 元素产生的图片/超链接/图表关系
        const slideRels = [
            { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
            ...ctx.rels
        ];
        zip.file(`ppt/slides/_rels/slide${slideIndex}.xml.rels`, buildRelationshipsXml(slideRels));

        // 媒体文件
        for (const media of ctx.media) {
            zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
            allMediaExts.add(media.name.split('.').pop().toLowerCase());
        }

        // 图表部件
        for (const chart of ctx.charts) {
            zip.file(`ppt/charts/${chart.name}`, chart.xml);
            allChartNames.push(chart.name);
        }

        slideRefs.push({ relId: `rId${presRelId++}`, target: `slides/slide${slideIndex}.xml` });
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
    zip.file('ppt/_rels/presentation.xml.rels', buildRelationshipsXml(presRels));

    // ===== 静态部件 =====
    zip.file('ppt/theme/theme1.xml', buildThemeXml());
    zip.file('ppt/slideMasters/slideMaster1.xml', buildSlideMasterXml());
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', buildRelationshipsXml(MASTER_RELS));
    zip.file('ppt/slideLayouts/slideLayout1.xml', buildSlideLayoutXml());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', buildRelationshipsXml(LAYOUT_RELS));
    zip.file('ppt/presProps.xml', buildPresPropsXml());
    zip.file('ppt/viewProps.xml', buildViewPropsXml());
    zip.file('ppt/tableStyles.xml', buildTableStylesXml());

    // ===== docProps =====
    zip.file('docProps/core.xml', buildCorePropsXml(pres.metadata));
    zip.file('docProps/app.xml', buildAppPropsXml(pres.slides.length));

    // ===== 包级文件 =====
    zip.file('_rels/.rels', buildRootRelsXml());
    let contentTypeXml = buildContentTypesXml([...allMediaExts], pres.slides.length);
    // 图表部件 Content-Types 覆盖
    for (const chartName of allChartNames) {
        const override = `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
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
function parseRelationships(relsText) {
    const rels = [];
    const re = /<Relationship\s+([^>]*?)\/>/g;
    const attrRe = /(\w+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(relsText)) !== null) {
        const attrs = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs.Id) rels.push(attrs);
    }
    return rels;
}

/** 从 presentation.xml 解析 sldIdLst 条目 */
function parseSldIdLst(presentationText) {
    const lstMatch = presentationText.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!lstMatch) return [];
    const entries = [];
    const re = /<p:sldId\s+([^>]*?)\/>/g;
    const attrRe = /([\w:.-]+)="([^"]*)"/g;
    let match;
    while ((match = re.exec(lstMatch[1])) !== null) {
        const attrs = {};
        let attrMatch;
        while ((attrMatch = attrRe.exec(match[1])) !== null) {
            attrs[attrMatch[1]] = attrMatch[2];
        }
        if (attrs['r:id']) entries.push({ id: attrs.id, relId: attrs['r:id'] });
    }
    return entries;
}

/** 用新条目列表重写 presentation.xml 的 sldIdLst */
function rewriteSldIdLst(presentationText, entries) {
    const inner = entries
        .map((e, i) => `<p:sldId id="${e.id || 256 + i}" r:id="${escapeXml(e.relId)}"/>`)
        .join('');
    return presentationText.replace(
        /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
        `<p:sldIdLst>${inner}</p:sldIdLst>`
    );
}

/** 列出 zip 中匹配 /ppt\/slides\/slide(\d+)\.xml 的幻灯片编号 */
function listSlideNumbers(zip) {
    const numbers = [];
    zip.forEach((path, entry) => {
        if (entry.dir) return;
        const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (m) numbers.push(Number(m[1]));
    });
    return numbers.sort((a, b) => a - b);
}

/** 列出 zip 中 ppt/media 下的最大媒体编号（不限 imageN 命名，取文件名中最后一个数字段） */
function maxMediaIndex(zip) {
    let max = 0;
    zip.forEach((path, entry) => {
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
async function editPptx(fileData) {
    const zip = await JSZip.loadAsync(fileData);

    const readText = async (name) => {
        const f = zip.file(name);
        return f ? await f.async('string') : null;
    };

    /** 读取 presentation.xml 解析后的信息 */
    async function getPresentationInfo() {
        const text = await readText('ppt/presentation.xml');
        if (!text) throw new Error('editPptx: 无效的 PPTX 文件（缺少 ppt/presentation.xml）');
        const relsText = await readText('ppt/_rels/presentation.xml.rels');
        const rels = parseRelationships(relsText || '');
        const relById = {};
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
    async function setAppSlideCount(n) {
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
    async function save(options = {}) {
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
        async getSlide(slideNum) {
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
        async deleteSlide(slideNum) {
            const info = await getPresentationInfo();
            if (slideNum < 1 || slideNum > info.entries.length) {
                throw new Error(`deleteSlide: 页码越界（共 ${info.entries.length} 页）`);
            }
            if (info.entries.length <= 1) {
                throw new Error('deleteSlide: 至少需保留一页幻灯片，无法删除最后一页');
            }

            const part = info.slideParts[slideNum - 1];
            const slidePath = part.target; // ppt/slides/slideN.xml

            // 1. 移除 sldIdLst 条目
            const remaining = info.entries.filter((_, i) => i !== slideNum - 1);
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, remaining));

            // 2. 移除 presentation.xml.rels 中的关系
            const relsText = await readText('ppt/_rels/presentation.xml.rels');
            const relRe = new RegExp(`<Relationship\\s+Id="${part.relId}"[^>]*/>`);
            zip.file('ppt/_rels/presentation.xml.rels', relsText.replace(relRe, ''));

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
        async moveSlide(from, to) {
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
            const mapping = {}; // 旧编号 → 新编号
            oldNums.forEach((oldNum, i) => { mapping[oldNum] = i + 1; });

            // 先读出全部内容，避免读写交叉覆盖
            const contents = [];
            const relsContents = [];
            for (const oldNum of oldNums) {
                contents.push(await readText(`ppt/slides/slide${oldNum}.xml`));
                relsContents.push(await readText(`ppt/slides/_rels/slide${oldNum}.xml.rels`));
            }

            // 按新编号写回，并重映射 rels 中的内部跳转目标（slideN.xml）
            oldNums.forEach((oldNum, i) => {
                const newNum = i + 1;
                const newRels = (relsContents[i] || '').replace(
                    /Target="slide(\d+)\.xml"/g,
                    (m, n) => `Target="slide${mapping[n] !== undefined ? mapping[n] : n}.xml"`
                );
                zip.file(`ppt/slides/slide${newNum}.xml`, contents[i]);
                zip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, newRels);
            });

            // 更新 presentation.xml.rels：各 relId 目标改为新编号文件
            let relsText = await readText('ppt/_rels/presentation.xml.rels');
            orderedRels.forEach((rel, i) => {
                const oldNum = oldNums[i];
                const newNum = i + 1;
                if (oldNum !== newNum) {
                    const re = new RegExp(`(<Relationship\\s+Id="${rel.Id}"[^>]*?Target=")slides/slide${oldNum}\\.xml(")`);
                    relsText = relsText.replace(re, `$1slides/slide${newNum}.xml$2`);
                }
            });
            zip.file('ppt/_rels/presentation.xml.rels', relsText);

            // 重写 sldIdLst（新逻辑顺序）
            zip.file('ppt/presentation.xml', rewriteSldIdLst(info.text, entries));
        },

        /**
         * 写回演示文稿元数据（docProps/core.xml）
         * @param {Object} metadata - 元数据（字段同 pptxToJson 返回的 metadata）
         */
        async setMetadata(metadata) {
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
        async addSlide(slideJson) {
            const info = await getPresentationInfo();
            const numbers = listSlideNumbers(zip);
            const nextNum = (numbers.length ? numbers[numbers.length - 1] : 0) + 1;

            // 构建新页内容
            const ctx = createElementContext({ startMediaIndex: maxMediaIndex(zip) });
            const slideRoot = await buildSlideRoot(ctx, slideJson);
            zip.file(`ppt/slides/slide${nextNum}.xml`, toXmlDocument(slideRoot));

            const slideRels = [
                { relId: 'rId1', type: REL_TYPES.slideLayout, target: '../slideLayouts/slideLayout1.xml' },
                ...ctx.rels
            ];
            zip.file(`ppt/slides/_rels/slide${nextNum}.xml.rels`, buildRelationshipsXml(slideRels));

            // 媒体文件
            const mediaExts = new Set();
            const newCtCharts = [];
            for (const media of ctx.media) {
                zip.file(`ppt/media/${media.name}`, media.base64, { base64: true });
                mediaExts.add(media.name.split('.').pop().toLowerCase());
            }

            // 图表部件
            for (const chart of ctx.charts) {
                zip.file(`ppt/charts/${chart.name}`, chart.xml);
                newCtCharts.push(chart.name);
            }

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
            zip.file('ppt/_rels/presentation.xml.rels', relsText.replace('</Relationships>', `${newRel}</Relationships>`));

            // Content-Types 追加
            const ctText = await readText('[Content_Types].xml');
            let newCt = ctText;
            const override = `<Override PartName="/ppt/slides/slide${nextNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
            newCt = newCt.replace('</Types>', `${override}</Types>`);
            const MIME_MAP = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml' };
            for (const ext of mediaExts) {
                if (!newCt.includes(`Extension="${ext}"`)) {
                    newCt = newCt.replace('</Types>', `<Default Extension="${ext}" ContentType="${MIME_MAP[ext] || 'application/octet-stream'}"/></Types>`);
                }
            }
            for (const chartName of newCtCharts) {
                if (!newCt.includes(`/ppt/charts/${chartName}`)) {
                    newCt = newCt.replace('</Types>', `<Override PartName="/ppt/charts/${chartName}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
                }
            }
            zip.file('[Content_Types].xml', newCt);
        }
    };

    /**
     * 移除 [Content_Types].xml 中指定部件的覆盖项
     * @param {string} partPath - 部件路径（如 ppt/slides/slide2.xml）
     */
    async function removeContentTypeOverride(partPath) {
        const ctText = await readText('[Content_Types].xml');
        if (!ctText) return;
        const re = new RegExp(`<Override PartName="/${partPath.replace(/\//g, '\\/')}"[^>]*/>`);
        zip.file('[Content_Types].xml', ctText.replace(re, ''));
    }
}

export { jsonToPptx, editPptx };
