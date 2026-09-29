/**
 * 生成「覆盖全部序列化能力」的 PPTX，供真机打开验证。
 * 运行：node examples/generate-test-pptx.mjs
 *
 * 同时在本脚本末尾做一轮解析自检（断言关键 OOXML 特征），
 * 任一断言失败则以非零退出码报错。
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
// 源码已转为 TypeScript，Node 无法直接运行 .ts，故改用构建产物（需先 npm run build）
import { PPTXComposer, pptxToJson } from '../dist/ppt-parser.esm.js';

// ---- 测试用图片（base64）----
// 64x64 蓝色 PNG
const BLUE_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAOklEQVR42u3OMQEAAAgDoJk/aRZw0gJ6mUm3bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2T9cCAr0H9dQAAAAASUVORK5CYII=';
// 64x64 橙色 PNG
const ORANGE_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAOklEQVR42u3OMQEAAAgDoJk/aRZw0gJ6mUm3bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2bdu2T9cCAr0H9dQAAAAASUVORK5CYII=';
// 1x1 白色 JPEG（验证 extension 覆盖）
const WHITE_JPEG = '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAA//EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AfwD/2Q==';

// 远程图片（生成时下载并内嵌；离线则回退到本地 PNG，保证文件始终可生成）
async function resolveRemoteImage() {
    const url = 'https://www.w3.org/2008/site/images/favicon.ico';
    try {
        if (typeof fetch === 'function') {
            const resp = await fetch(url);
            if (resp.ok) {
                const buf = await resp.arrayBuffer();
                const b64 = Buffer.from(new Uint8Array(buf)).toString('base64');
                return `data:image/x-icon;base64,${b64}`;
            }
        }
    } catch (_) { /* 离线回退 */ }
    console.warn('[warn] 远程图片下载失败，已回退为本地 PNG（src 能力仍走同一代码路径）');
    return `data:image/png;base64,${BLUE_PNG}`;
}

const composer = new PPTXComposer();
composer
    .title('PPTX 序列化能力全覆盖')
    .author('pptx-parser')
    .subject('真机打开验证')
    .keywords('serializer;test;all-features')
    .description('覆盖文本/形状/图片/幻灯片/演示文稿各级别全部生成能力')
    .slideSize(1280, 720);

// ============ 第 1 页：文本能力全集 ============
composer.addSlide(slide => {
    slide.background('#f8fafc');

    // 标题（value + bold + color）
    slide.addText(t => t
        .value('第1页 · 文本能力全集')
        .x(60).y(40).width(900).height(70)
        .fontSize(36).color('#0f172a').bold());

    // 普通正文 + 对齐 + 颜色
    slide.addText(t => t
        .value('普通正文：颜色 / 左对齐 / 默认样式')
        .x(60).y(130).width(700).height(36)
        .fontSize(18).color('#475569'));

    // 粗体 / 斜体 / 下划线 各一行
    slide.addText(t => t.value('粗体 bold').x(60).y(180).width(300).height(30).fontSize(18).bold());
    slide.addText(t => t.value('斜体 italic').x(360).y(180).width(300).height(30).fontSize(18).italic());
    slide.addText(t => t.value('下划线 underline').x(660).y(180).width(300).height(30).fontSize(18).underline());

    // 指定字体 fontFace + 语言 lang
    slide.addText(t => t
        .value('指定字体 Arial + lang=en-US')
        .x(60).y(225).width(700).height(30)
        .fontSize(18).fontFace('Arial').lang('en-US'));

    // 垂直居中文本框（valign=middle）
    slide.addText(t => t
        .value('垂直居中(middle)')
        .x(60).y(270).width(300).height(80)
        .fontSize(20).color('#7c3aed').valign('middle'));

    // 项目符号段落（bullet）
    slide.addText(t => t
        .value('').x(420).y(270).width(500).height(120)
        .fontSize(18)
        .paragraphs([
            { text: '项目符号 第一项', bullet: true },
            { text: '项目符号 第二项', bullet: true },
            { text: '无符号 第三项', bullet: false }
        ]));

    // runs 规范写法：外部 + 内部跳转超链接、不同颜色
    slide.addText(t => t
        .runs([
            { text: '外部链接', options: { href: 'https://github.com', color: '#2563eb', underline: true } },
            { text: '  ', options: {} },
            { text: '跳到第3页', options: { href: '#3', color: '#dc2626' } }
        ])
        .x(60).y(410).width(700).height(30));

    // runs 扁平简写 + 多行 \n 混合
    slide.addText(t => t
        .runs([
            { text: '扁平简写红色', color: '#ef4444', bold: true },
            { text: '\n换行的第二行', color: '#0891b2' }
        ])
        .x(60).y(455).width(700).height(60).fontSize(16));

    // 命名文本框（验证 cNvPr name）
    slide.addText(t => t
        .value('命名文本框(name=MyTextBox)')
        .x(60).y(525).width(700).height(30)
        .fontSize(16).name('MyTextBox'));
});

// ============ 第 2 页：形状能力全集 ============
composer.addSlide(slide => {
    slide.background('#0f172a');
    slide.addText(t => t
        .value('第2页 · 形状能力全集')
        .x(60).y(40).width(900).height(70)
        .fontSize(36).color('#f8fafc').bold());

    const palette = ['#ef4444', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#ec4899'];
    const types = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'hexagon', 'star5', 'rightArrow', 'pie', 'plus'];

    types.forEach((type, i) => {
        const col = i % 5;
        const row = Math.floor(i / 5);
        slide.addShape(s => s
            .shapeType(type)
            .x(60 + col * 230).y(130 + row * 200).width(180).height(150)
            .fill({ color: palette[i % palette.length] })
            .line({ color: '#ffffff', width: 1.5 }));
    });

    // 无填充 + 蓝色边框
    slide.addShape(s => s
        .shapeType('rect')
        .x(60).y(540).width(220).height(120)
        .fill('none')
        .line({ color: '#22d3ee', width: 3 }));

    // 无边框（line=none）
    slide.addShape(s => s
        .shapeType('ellipse')
        .x(310).y(540).width(120).height(120)
        .fill({ color: '#f43f5e' })
        .line('none'));

    // 旋转 30 度
    slide.addShape(s => s
        .shapeType('roundRect')
        .x(470).y(540).width(180).height(90)
        .fill({ color: '#84cc16' })
        .rotation(30));

    // 命名形状
    slide.addShape(s => s
        .shapeType('diamond')
        .x(700).y(540).width(120).height(120)
        .fill({ color: '#eab308' })
        .name('MyShape'));
});

// ============ 第 3 页：图片能力全集 ============
const remoteData = await resolveRemoteImage();
composer.addSlide(slide => {
    slide.background('#ffffff');
    slide.addText(t => t
        .value('第3页 · 图片能力全集')
        .x(60).y(40).width(900).height(70)
        .fontSize(36).color('#111827').bold());

    // data: base64 PNG
    slide.addImage({ data: `data:image/png;base64,${BLUE_PNG}`, x: 60, y: 130, width: 140, height: 140, name: 'Base64Png' });
    // 裸 base64 + extension 覆盖（JPEG）
    slide.addImage({ data: WHITE_JPEG, extension: 'jpeg', x: 240, y: 130, width: 140, height: 140, name: 'JpegExt' });
    // 远程 src（生成时内嵌）
    slide.addImage({ src: remoteData, x: 420, y: 130, width: 140, height: 140, name: 'RemoteImg' });
    // 图片级外部超链接
    slide.addImage({ data: `data:image/png;base64,${ORANGE_PNG}`, href: 'https://example.com', x: 600, y: 130, width: 140, height: 140, name: 'LinkedImg' });

    slide.addText(t => t
        .value('左上:base64 PNG  中左:裸base64+JPEG扩展  中右:远程src内嵌  右下:带外部链接')
        .x(60).y(300).width(1000).height(60).fontSize(16).color('#64748b'));
});

// ============ 第 4 页：幻灯片/演示级（背景 none / 尺寸 / 元数据）============
composer.addSlide(slide => {
    // 不设置 background（即 'none'，使用默认白底，验证背景缺失分支）
    slide.addText(t => t
        .value('第4页 · 背景 none（默认白底）+ 演示级元数据已在文件头写入')
        .x(60).y(60).width(1000).height(80)
        .fontSize(24).color('#111827'));
});

// ============ 第 5 页：图表能力全集（原生 OOXML 图表）============
composer.addSlide(slide => {
    slide.background('#ffffff');
    slide.addText(t => t
        .value('第5页 · 图表能力全集（PowerPoint/WPS 可编辑真图表）')
        .x(40).y(40).width(1100).height(60)
        .fontSize(28).color('#111827').bold());

    slide.addChart({
        chartType: 'barChart', title: '柱状图', x: 40, y: 120, width: 560, height: 260,
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [
            { name: '销售额', values: [120, 200, 150, 180] },
            { name: '利润', values: [30, 55, 40, 60] }
        ]
    });
    slide.addChart({
        chartType: 'lineChart', title: '折线图', x: 640, y: 120, width: 560, height: 260,
        categories: ['1月', '2月', '3月'],
        series: [{ name: '访问量', values: [300, 400, 350] }]
    });
    slide.addChart({
        chartType: 'pieChart', title: '饼图', x: 40, y: 400, width: 560, height: 260,
        categories: ['A', 'B', 'C'],
        series: [{ name: '占比', values: [40, 35, 25] }]
    });
    slide.addChart({
        chartType: 'scatterChart', title: '散点图', x: 640, y: 400, width: 560, height: 260,
        series: [{ name: '样本', x: [1, 2, 3, 4], y: [2, 4, 1, 5] }]
    });
});

// ===================== 生成并写入 =====================
const data = await composer.save(); // 默认 Uint8Array
const outPath = path.resolve('examples/test-sample.pptx');
fs.writeFileSync(outPath, Buffer.from(data));
console.log(`已生成: ${outPath} (${(data.byteLength / 1024).toFixed(1)} KB)`);

// ===================== 自检（解析断言）=====================
async function selfCheck() {
    const buffer = fs.readFileSync(outPath);
    const ab = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    const result = await pptxToJson(ab);
    const zip = await JSZip.loadAsync(buffer);

    const slideXml = [];
    const slideRels = [];
    for (let i = 1; i <= result.slides.length; i++) {
        slideXml.push(await zip.file(`ppt/slides/slide${i}.xml`).async('string'));
        const relFile = zip.file(`ppt/slides/_rels/slide${i}.xml.rels`);
        if (relFile) slideRels.push(await relFile.async('string'));
    }
    const all = slideXml.join('\n') + '\n' + slideRels.join('\n');
    const contentTypes = await zip.file('[Content_Types].xml').async('string');

    // 图表部件内容（用于断言 chart 类型与数据）
    const chartXml = [];
    for (const name of Object.keys(zip.files)) {
        if (/^ppt\/charts\/chart\d+\.xml$/.test(name)) {
            chartXml.push(await zip.file(name).async('string'));
        }
    }
    const chartAll = chartXml.join('\n');

    const checks = [];
    const assert = (name, cond) => checks.push({ name, ok: !!cond });

    // 演示级
    assert('页数=5', result.slides.length === 5);
    assert('slideSize=1280x720', result.slideSize.width === 1280 && result.slideSize.height === 720);
    assert('metadata.title', result.metadata.title === 'PPTX 序列化能力全覆盖');
    assert('metadata.keywords', result.metadata.keywords === 'serializer;test;all-features');
    assert('metadata.description', /覆盖文本/.test(result.metadata.description || ''));

    // 文本能力
    assert('粗体 b=1', /<a:rPr[^>]*\sb="1"/.test(all));
    assert('斜体 i=1', /<a:rPr[^>]*\si="1"/.test(all));
    assert('下划线 u=sng', /u="sng"/.test(all));
    assert('字体 latin', /<a:latin/.test(all));
    assert('垂直居中 anchor=ctr', /anchor="ctr"/.test(all));
    assert('项目符号 buChar', /<a:buChar/.test(all));
    assert('内部跳转 ppaction', /ppaction:\/\/hlinksldjump/.test(all));
    assert('外部超链接 rel', /TargetMode="External"/.test(all));
    assert('多行(多个 a:p)', (all.match(/<a:p>/g) || []).length >= 5);
    assert('命名文本框 name', /MyTextBox/.test(all));
    assert('lang=en-US', /lang="en-US"/.test(all));

    // 形状能力
    assert('旋转 rot', /rot="/.test(all));
    assert('无填充 noFill', /<a:noFill\/>/.test(all));
    assert('line=none 的 noFill', /<a:ln[^>]*><a:noFill\/>/.test(all));
    assert('多种 prstGeom', /prst="diamond"/.test(all) && /prst="hexagon"/.test(all) && /prst="star5"/.test(all));
    assert('命名形状 MyShape', /MyShape/.test(all));
    assert('边框线宽 ln w', /<a:ln[^>]*\bw="/.test(all));

    // 图片能力
    assert('图片 blip', /<a:blip/.test(all));
    assert('远程/内嵌媒体存在', !!zip.file('ppt/media/image1.png') || Object.keys(zip.files).some(f => f.startsWith('ppt/media/')));
    assert('JPEG 扩展媒体', Object.keys(zip.files).some(f => /ppt\/media\/image\d+\.jpeg$/.test(f)));

    // 第4页背景 none（无 p:bg）
    const slide4 = slideXml[3] || '';
    assert('第4页无背景 p:bg', !/<p:bg>/.test(slide4));
    // 第1页有背景
    assert('第1页有背景 p:bg', /<p:bg>/.test(slideXml[0]));

    // 图表能力（原生 OOXML）
    assert('graphicFrame 存在', /<p:graphicFrame/.test(all));
    assert('图表部件 chart1.xml', !!zip.file('ppt/charts/chart1.xml'));
    assert('图表部件 chart2~4', !!zip.file('ppt/charts/chart2.xml') && !!zip.file('ppt/charts/chart3.xml') && !!zip.file('ppt/charts/chart4.xml'));
    assert('c:barChart', /<c:barChart/.test(chartAll));
    assert('c:lineChart', /<c:lineChart/.test(chartAll));
    assert('c:pieChart', /<c:pieChart/.test(chartAll));
    assert('c:scatterChart', /<c:scatterChart/.test(chartAll));
    assert('图表类别/系列名 销售额', /销售额/.test(chartAll));
    assert('图表数值 <c:v>120', /<c:v>120<\/c:v>/.test(chartAll));
    assert('图表关系 rel(chart)', /relationships\/chart/.test(all));
    assert('ContentTypes chart 覆盖', /drawingml\.chart\+xml/.test(contentTypes));
    // 注：pptxToJson(JSON 路径) 目前不回解图表（仅 pptxToHtml 浏览器渲染会处理），
    // 故此处直接校验生成的图表部件 XML 结构正确（含 chartSpace/plotArea 与各类型）。
    assert('图表部件=4 且结构有效', chartXml.length === 4 && chartXml.every(x => /<c:chartSpace/.test(x) && /<c:plotArea>/.test(x)));

    let failed = 0;
    for (const c of checks) {
        console.log(`  ${c.ok ? 'PASS' : 'FAIL'}  ${c.name}`);
        if (!c.ok) failed++;
    }
    if (failed > 0) {
        throw new Error(`自检失败：${failed}/${checks.length} 项未通过`);
    }
    console.log(`\n自检全部通过：${checks.length}/${checks.length}`);
}

await selfCheck();
