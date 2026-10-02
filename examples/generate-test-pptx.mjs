/**
 * 生成「覆盖全部序列化能力」的 PPTX，供真机打开验证。
 * 运行：node examples/generate-test-pptx.mjs   （需先 npm run build，产物在 dist/）
 *
 * 按 PPTX_GEN_GAP_PLAN.md，每一页对应一个功能点（T1–T19），并附带一轮解析自检。
 * 任一断言失败则以非零退出码报错。
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { jsonToPptx, pptxToJson } from '../dist/ppt-parser.esm.js';

// ---- 测试用图片（base64）----
const BLUE_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAZUlEQVR42u3QQREAAAQAML20E9qXHM4eK7DI6vksBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAAAECBAgQIECAgPsWcJEihvVdy3EAAAAASUVORK5CYII=';
// 四象限色块图（32x32：左上红/右上绿/左下蓝/右下黄 + 中心黑块），用于平铺与裁剪的肉眼验证
const QUAD_PNG = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAVElEQVR42mO4o6b2nxIsttiLIsww6oBRB4w6YNQBow4Ycg4QlFDHi0cdMPQcoJr8+j8pmJADfp0RJQmPOmDoOQAdk2rhqANGHTDqgFEHjDpg0DkAAOrUMfjFkhTuAAAAAElFTkSuQmCC';
const ORANGE_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAZklEQVR42u3QIQ0AAAgAMDz9LS3RkINx8QKPrpzPQoAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECBAgQIAAAQIECLhvAaF00mjwgbtAAAAAAElFTkSuQmCC';

// 自定义主题 XML（覆盖默认主题，演示 T17：accent1..6 改为自定义色）
const CUSTOM_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="CustomTest">
 <a:themeElements>
  <a:clrScheme name="CustomTest">
   <a:dk1><a:sysClr val="windowText" lastClr="1F2937"/></a:dk1>
   <a:lt1><a:sysClr val="window" lastClr="F8FAFC"/></a:lt1>
   <a:dk2><a:srgbClr val="0F172A"/></a:dk2>
   <a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>
   <a:accent1><a:srgbClr val="6366F1"/></a:accent1>
   <a:accent2><a:srgbClr val="EC4899"/></a:accent2>
   <a:accent3><a:srgbClr val="10B981"/></a:accent3>
   <a:accent4><a:srgbClr val="F59E0B"/></a:accent4>
   <a:accent5><a:srgbClr val="3B82F6"/></a:accent5>
   <a:accent6><a:srgbClr val="8B5CF6"/></a:accent6>
   <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
   <a:folHlink><a:srgbClr val="9333EA"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="CustomTest">
   <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
   <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name="CustomTest">
   <a:fillStyleLst>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
   </a:fillStyleLst>
   <a:lnStyleLst>
    <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
    <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
    <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
   </a:lnStyleLst>
   <a:effectStyleLst>
    <a:effectStyle><a:effectLst/></a:effectStyle>
    <a:effectStyle><a:effectLst/></a:effectStyle>
    <a:effectStyle><a:effectLst/></a:effectStyle>
   </a:effectStyleLst>
   <a:bgFillStyleLst>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
   </a:bgFillStyleLst>
  </a:fmtScheme>
 </a:themeElements>
</a:theme>`;

// ---- 辅助：构造一页（标题 + 功能元素 + 可选 slide 级扩展）----
function page(title, elements, extra = {}) {
    return {
        background: '#ffffff',
        elements: [
            { type: 'text', x: 60, y: 30, width: 1160, height: 50, text: title, fontSize: 30, bold: true, color: '#0f172a' },
            ...elements
        ],
        ...extra
    };
}

const slides = [];

// ============ T1 表格单元格边框线 ============
slides.push(page('T1 · 表格单元格边框线', [
    { type: 'table', x: 60, y: 110, width: 1100, height: 280,
        colWidths: [280, 280, 280, 260],
        rows: [
            { cells: [
                { text: '四边粗框', border: { color: '#ef4444', width: 2 } },
                { text: '上下右分边', borders: { top: { color: '#10b981', width: 2 }, right: { color: '#10b981', width: 2 } } },
                { text: '左边粗', borders: { left: { color: '#3b82f6', width: 3 } } },
                { text: '无框', borders: { top: 'none', bottom: 'none', left: 'none', right: 'none' } }
            ] },
            { cells: [
                { text: '细框', border: { color: '#000000', width: 1 } },
                { text: '橙框', border: { color: '#f59e0b', width: 3 } },
                { text: '默认', },
                { text: '蓝框', border: { color: '#2563eb', width: 2 } }
            ] }
        ]
    }
]));

// ============ T2 形状渐变填充 + 透明度 + 阴影/发光 ============
slides.push(page('T2 · 形状渐变/透明度/阴影/发光', [
    { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 150,
        fill: { type: 'gradient', direction: 'horizontal', stops: [{ position: 0, color: '#6366f1' }, { position: 1, color: '#ec4899' }] },
        line: { color: '#111827', width: 1 } },
    { type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 150,
        fill: { type: 'solid', color: '#10b981', transparency: 40 }, line: { color: '#047857', width: 1.5 } },
    { type: 'shape', shapeType: 'ellipse', x: 620, y: 120, width: 160, height: 160,
        fill: { color: '#3b82f6' }, effects: { shadow: { blur: 10, distance: 5, angle: 90, color: '#000000', transparency: 40 } } },
    { type: 'shape', shapeType: 'diamond', x: 820, y: 120, width: 160, height: 160,
        fill: { color: '#f59e0b' }, effects: { glow: { color: '#ef4444', blur: 15 } } }
]));

// ============ T3 文本编号列表 + 行距/段间距/缩进 ============
slides.push(page('T3 · 文本编号列表/行距/缩进', [
    { type: 'text', x: 60, y: 120, width: 1040, height: 420, fontSize: 18, color: '#334155',
        paragraphs: [
            // fmt 必须是合法的 ST_TextAutonumberScheme：'ea1ChsPeriod' = 中文编号「一、二、」
            { text: '编号第一项', bullet: { type: 'number', fmt: 'ea1ChsPeriod' } },
            { text: '编号第二项', bullet: { type: 'number', fmt: 'ea1ChsPeriod' } },
            { text: '缩进 + 行距28pt + 段前/后10', bullet: true, indent: 40, lineSpacing: { type: 'pt', value: 28 }, spaceBefore: 10, spaceAfter: 10 },
            { text: '普通项（无符号）', bullet: false }
        ]
    }
]));

// ============ T4 组合 grpSp ============
slides.push(page('T4 · 组合 grpSp（childrenCoordinates=local）', [
    { type: 'group', x: 80, y: 120, width: 600, height: 320, childrenCoordinates: 'local',
        children: [
            { type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 200, height: 100, fill: { color: '#60a5fa' } },
            { type: 'text', x: 20, y: 120, width: 320, height: 40, text: '组合内文本', fontSize: 18 },
            { type: 'shape', shapeType: 'ellipse', x: 260, y: 60, width: 120, height: 120, fill: { color: '#f472b6' } }
        ]
    }
]));

// ============ T5 主题色引用 ============
slides.push(page('T5 · 主题色引用（scheme:）', [
    { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 120,
        fill: { type: 'solid', color: 'scheme:accent1' }, line: { color: 'scheme:dk1', width: 1.5 } },
    { type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 120,
        fill: { type: 'solid', color: 'scheme:accent2' } },
    { type: 'shape', shapeType: 'ellipse', x: 620, y: 120, width: 160, height: 160,
        fill: { type: 'solid', color: 'scheme:accent3' } },
    { type: 'text', x: 60, y: 300, width: 1040, height: 40, text: '引用 scheme:accent1/2/3 与 scheme:dk1（自定义主题见 T17）', color: 'scheme:dk1', fontSize: 18 }
]));

// ============ T6 形状图片填充 + 图案填充 ============
slides.push(page('T6 · 形状图片填充 / 图案填充 / 平铺裁剪', [
    { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 160,
        fill: { type: 'image', data: `data:image/png;base64,${BLUE_PNG}` } },
    { type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 160,
        fill: { type: 'pattern', prst: 'diagCross', fg: '#ef4444', bg: '#fef3c7' } },
    // 平铺：以原图尺寸（32px）为一格重复，整体不能被拉伸
    { type: 'shape', shapeType: 'rect', x: 620, y: 120, width: 240, height: 160,
        fill: {
            type: 'image', data: `data:image/png;base64,${QUAD_PNG}`,
            tile: { sx: 1, sy: 1 }
        } },
    // 平铺 + 源图裁剪叠加：只取图片居中 50% 作为一格，仍按原图尺寸重复
    { type: 'shape', shapeType: 'rect', x: 900, y: 120, width: 240, height: 160,
        fill: {
            type: 'image', data: `data:image/png;base64,${QUAD_PNG}`,
            tile: { sx: 1, sy: 1 },
            srcRect: { l: 0.25, t: 0.25, r: 0.25, b: 0.25 }
        } },
    { type: 'text', x: 60, y: 300, width: 1100, height: 60, fontSize: 14, color: '#475569',
        text: '① 图片铺满｜② 图案填充 diagCross｜③ 图片平铺（每格 = 原图 32px）｜④ 平铺 + 裁剪居中 50%（每格显示原图中心 16px）' }
]));

// ============ T7 形状几何调整值 avLst ============
slides.push(page('T7 · 形状几何调整 avLst', [
    { type: 'shape', shapeType: 'roundRect', x: 60, y: 120, width: 240, height: 120,
        fill: { color: '#6366f1' }, adjust: { adj: 30000 } },
    // adjust 的 key 必须是 OOXML 的 gd 名（rightArrow 为 adj1/adj2）：
    // adj1 = 箭身厚度（相对高度），adj2 = 箭头长度（dx1 = min(w,h) * adj2 / 100000）
    { type: 'shape', shapeType: 'rightArrow', x: 340, y: 120, width: 240, height: 120,
        fill: { color: '#10b981' }, adjust: { adj1: 50000, adj2: 40000 } },
    { type: 'shape', shapeType: 'rightArrow', x: 620, y: 120, width: 240, height: 120,
        fill: { color: '#0ea5e9' }, adjust: { adj1: 30000, adj2: 100000 } },
    { type: 'text', x: 60, y: 300, width: 900, height: 40, fontSize: 14, color: '#475569',
        text: '圆角矩形 adj=30000｜箭头 adj1=50000/adj2=40000（箭头长 20%）｜箭头 adj1=30000/adj2=100000（箭头长 50%、箭身更细）' }
]));

// ============ T8 水平/垂直翻转 ============
slides.push(page('T8 · 水平/垂直翻转', [
    { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 200, height: 140, fill: { color: '#3b82f6' }, flipH: true },
    { type: 'shape', shapeType: 'rect', x: 320, y: 120, width: 200, height: 140, fill: { color: '#ef4444' }, flipV: true },
    { type: 'shape', shapeType: 'diamond', x: 580, y: 120, width: 160, height: 160, fill: { color: '#10b981' }, flipH: true, flipV: true }
]));

// ============ T9 文本框内边距 + 竖排 ============
slides.push(page('T9 · 文本框内边距 / 竖排', [
    { type: 'text', x: 60, y: 120, width: 240, height: 200, text: '内边距 l/r/t/b=20', fontSize: 16, color: '#0f172a', inset: { l: 20, r: 20, t: 20, b: 20 } },
    // textDirection 必须是合法的 ST_TextVerticalType（eaVert=东亚竖排/竖排，wordArtVert=堆积），
    // 写成 'wordArtVertical' 这类非枚举值会被 PowerPoint/WPS 静默忽略并回退横排
    { type: 'text', x: 360, y: 120, width: 160, height: 320, text: '竖排文字', fontSize: 22, color: '#7c3aed', textDirection: 'eaVert' }
]));

// ============ T10 单元格内边距 + 对角线边框 + 表格样式 ============
// tableStyleId 用内置的 “Medium Style 2 - Accent 1”（强调色底纹 + 白色网格线）；
// 生成端会把该 ID 的等价定义写进 ppt/tableStyles.xml —— 未知 GUID 会让 WPS/PowerPoint 退化成「无样式无网格」。
slides.push(page('T10 · 单元格内边距/对角线/表格样式', [
    { type: 'table', x: 60, y: 120, width: 900, height: 260,
        tableStyleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
        rows: [
            { cells: [
                { text: '内边距 20px', inset: { l: 20, r: 20, t: 10, b: 10 } },
                { text: '对角线 tlBr', borders: { diagonal: 'tlBr' } },
                { text: '对角线 both', borders: { diagonal: 'both' } }
            ] },
            { cells: [
                { text: 'A：默认内边距' },
                { text: 'B' },
                { text: 'C' }
            ] }
        ]
    },
    { type: 'text', x: 60, y: 400, width: 900, height: 40, fontSize: 14, color: '#475569',
        text: '样式：Medium Style 2 - Accent 1（首行强调色底+白字、白色网格线）｜内边距：左上 20/10px，其余用默认 9.6/4.8px' }
]));

// ============ T11 图表类型与特性扩展 ============
slides.push(page('T11 · 图表类型扩展（3D/环/气泡/雷达/股票/曲面 + 数据标签/图例）', [
    { type: 'chart', chartType: 'bar3DChart', title: '3D柱', x: 40, y: 120, width: 320, height: 200,
        categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'A', values: [10, 20, 15] }] },
    { type: 'chart', chartType: 'doughnutChart', title: '环图', legend: 'b', x: 380, y: 120, width: 320, height: 200,
        categories: ['a', 'b', 'c'], series: [{ name: '占比', values: [40, 35, 25] }] },
    { type: 'chart', chartType: 'bubbleChart', title: '气泡', x: 720, y: 120, width: 320, height: 200,
        series: [{ name: 'b', x: [1, 2, 3], y: [2, 3, 1], values: [3, 4, 2] }] },
    { type: 'chart', chartType: 'radarChart', title: '雷达', dataLabels: { showValue: true }, x: 40, y: 340, width: 320, height: 200,
        series: [{ name: 'r', values: [1, 2, 3, 4] }] },
    { type: 'chart', chartType: 'stockChart', title: '股票', x: 380, y: 340, width: 320, height: 200,
        series: [{ name: 's', open: [1, 2, 3], high: [4, 5, 6], low: [0, 1, 2], close: [2, 3, 4] }] },
    { type: 'chart', chartType: 'surfaceChart', title: '曲面', x: 720, y: 340, width: 320, height: 200,
        series: [{ name: 'f', values: [1, 2, 3] }] }
]));

// ============ T12 图片裁剪 + 调整/透明度 ============
slides.push(page('T12 · 图片裁剪 / 调整(亮度/对比度/透明度)', [
    { type: 'image', data: `data:image/png;base64,${BLUE_PNG}`, x: 60, y: 120, width: 240, height: 240,
        crop: { l: 0.1, r: 0.1, t: 0.1, b: 0.1 }, name: 'Crop' },
    { type: 'image', data: `data:image/png;base64,${ORANGE_PNG}`, x: 360, y: 120, width: 240, height: 240,
        imageAdjust: { brightness: 20, contrast: 30, transparency: 10 }, name: 'Adjust' }
]));

// ============ T13 动画 timing ============
slides.push(page('T13 · 动画 timing（淡入 + 自动播放）', [
    { type: 'shape', shapeType: 'rect', x: 240, y: 180, width: 320, height: 160, fill: { color: '#6366f1' }, name: 'AnimShape' }
], { animations: [{ target: 1, type: 'fade', duration: 1 }] }));

// ============ T14 切换时序/自动播放 + 隐藏幻灯片 ============
slides.push(page('T14 · 切换/自动播放 + 隐藏幻灯片', [
    { type: 'shape', shapeType: 'rect', x: 240, y: 180, width: 320, height: 160, fill: { color: '#10b981' } }
], {
    transition: { type: 'fade', duration: 1000, advanceOnClick: false },
    advanceTime: 3000,
    hidden: true
}));

// ============ T15 视频/音频 media ============
slides.push(page('T15 · 视频 / 音频 media', [
    { type: 'video', data: 'data:video/mp4;base64,AAAA', extension: 'mp4', x: 60, y: 120, width: 360, height: 210,
        poster: { data: `data:image/png;base64,${BLUE_PNG}` }, name: 'Vid' },
    { type: 'audio', data: 'data:audio/m4a;base64,AAAA', extension: 'm4a', x: 480, y: 120, width: 200, height: 200, name: 'Aud' }
]));

// ============ T16 SmartArt 图示 ============
slides.push(page('T16 · SmartArt 图示（结构级四件套）', [
    { type: 'diagram', diagramType: 'hierarchy', x: 60, y: 120, width: 820, height: 400,
        nodes: [{ text: '根', children: [
            { text: '子A' },
            { text: '子B', children: [{ text: '孙' }] }
        ] }]
    }
]));

// ============ T17 自定义主题 / 母版 / 版式 ============
// 注：本文件整体已通过 options.theme 覆盖为 CUSTOM_THEME（见 jsonToPptx 调用）；此处演示主题色在页面中的引用
slides.push(page('T17 · 自定义主题覆盖（options.theme）', [
    { type: 'shape', shapeType: 'rect', x: 60, y: 120, width: 240, height: 120, fill: { type: 'solid', color: 'scheme:accent4' } },
    { type: 'shape', shapeType: 'roundRect', x: 340, y: 120, width: 240, height: 120, fill: { type: 'solid', color: 'scheme:accent5' } },
    { type: 'text', x: 60, y: 280, width: 1040, height: 40, text: 'accent4/accent5 使用自定义主题色（theme1.xml 已覆盖）', color: 'scheme:dk2', fontSize: 18 }
]));

// ============ T18 批注 comments ============
slides.push(page('T18 · 批注 comments', [
    { type: 'shape', shapeType: 'note', x: 200, y: 200, width: 300, height: 120, fill: { color: '#fde68a' } }
], {
    comments: [
        { author: 'Alice', text: '这是一条批注', dt: '2026-01-01T00:00:00Z' },
        { author: 'Bob', text: '第二条批注' }
    ]
}));

// ============ T19 自定义属性 ============
slides.push(page('T19 · 自定义文档属性（docProps/custom.xml）', [
    { type: 'text', x: 60, y: 130, width: 1040, height: 40, text: '本文档写入了自定义属性：部门 / 版本 / 项目', fontSize: 18, color: '#475569' }
]));

// ===================== 生成 =====================
const pres = {
    slideSize: { width: 1280, height: 720 },
    metadata: {
        title: 'PPTX 序列化能力全覆盖（每功能点一页）',
        author: 'pptx-parser',
        subject: '真机打开验证',
        keywords: 'serializer;test;T1-T19',
        description: '覆盖 T1-T19 全部生成能力'
    },
    customProps: { '部门': '研发', '版本': '1.0', '项目': 'pptx-parser' },
    slides
};

const data = await jsonToPptx(pres, { theme: CUSTOM_THEME }); // 默认 Uint8Array
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
    for (let i = 1; i <= result.slides.length; i++) {
        slideXml.push(await zip.file(`ppt/slides/slide${i}.xml`).async('string'));
    }
    const all = slideXml.join('\n');
    const contentTypes = await zip.file('[Content_Types].xml').async('string');
    const tableStylesXml = await zip.file('ppt/tableStyles.xml').async('string');

    const chartXml = [];
    for (const name of Object.keys(zip.files)) {
        if (/^ppt\/charts\/chart\d+\.xml$/.test(name)) chartXml.push(await zip.file(name).async('string'));
    }
    const chartAll = chartXml.join('\n');

    const checks = [];
    const assert = (name, cond) => checks.push({ name, ok: !!cond });

    assert('页数=19', result.slides.length === 19);
    assert('slideSize=1280x720', result.slideSize.width === 1280 && result.slideSize.height === 720);
    assert('metadata.title', result.metadata.title === 'PPTX 序列化能力全覆盖（每功能点一页）');
    assert('metadata.keywords', result.metadata.keywords === 'serializer;test;T1-T19');

    // T1 边框
    assert('T1 单元格边框 lnL/lnR/lnT/lnB', /<a:lnL|<a:lnR|<a:lnT|<a:lnB/.test(slideXml[0]));
    // T2 渐变/特效/透明度
    assert('T2 渐变 gradFill', /<a:gradFill/.test(slideXml[1]));
    assert('T2 特效 effectLst', /<a:effectLst/.test(slideXml[1]));
    assert('T2 透明度 alpha', /<a:alpha/.test(slideXml[1]));
    // T3 编号/行距
    assert('T3 编号 buAutoNum', /<a:buAutoNum/.test(slideXml[2]));
    // 编号 type 必须是合法的 ST_TextAutonumberScheme（否则 WPS/PowerPoint 会回退成默认编号）
    assert('T3 编号 type 合法', /<a:buAutoNum type="(ea1ChsPeriod|arabicPeriod|alphaLcPeriod|romanUcPeriod)"/.test(slideXml[2]));
    assert('T3 行距 lnSpc', /<a:lnSpc/.test(slideXml[2]));
    // T4 组合
    assert('T4 组合 grpSp', /<p:grpSp/.test(slideXml[3]));
    // T5 主题色引用
    assert('T5 schemeClr 引用', /<a:schemeClr/.test(slideXml[4]));
    // T6 图片/图案填充
    assert('T6 形状图片填充 blipFill', /<a:blipFill/.test(slideXml[5]));
    assert('T6 图片填充 r:embed 关系', /<a:blip r:embed="rId\d+"\/>/.test(slideXml[5]));
    assert('T6 图案填充 pattFill', /<a:pattFill/.test(slideXml[5]));
    assert('T6 图片平铺 a:tile', /<a:tile sx="100000" sy="100000" tx="0" ty="0"\/>/.test(slideXml[5]));
    assert('T6 源图裁剪 a:srcRect', /<a:srcRect l="25000" t="25000" r="25000" b="25000"\/>/.test(slideXml[5]));
    assert('T6 平铺+裁剪叠加顺序', /<a:srcRect[^>]*\/><a:tile[^>]*\/>/.test(slideXml[5]));
    // 素材必须是有效 PNG：此前用了一段截断的 base64，图片填充在浏览器里整块不可见
    const pngSig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const bluePngBuf = Buffer.from(BLUE_PNG, 'base64');
    assert('T6 图片素材为有效 PNG', bluePngBuf.subarray(0, 8).equals(pngSig) && bluePngBuf.includes(Buffer.from('IEND')));
    // T7 调整值 avLst
    assert('T7 几何调整 avLst', /<a:avLst/.test(slideXml[6]));
    // key 必须是 OOXML 的 gd 名（adj/adj1/adj2），否则 WPS/PowerPoint 会忽略并回退默认值
    assert('T7 圆角 adj', /<a:gd name="adj" fmla="val 30000"\/>/.test(slideXml[6]));
    assert('T7 箭头 adj1/adj2', /<a:gd name="adj1" fmla="val 50000"\/><a:gd name="adj2" fmla="val 40000"\/>/.test(slideXml[6]));
    // T8 翻转
    assert('T8 翻转 flipH/flipV', /flipH="1"|flipV="1"/.test(slideXml[7]));
    // T9 内边距/竖排
    assert('T9 内边距 lIns', /lIns=/.test(slideXml[8]));
    // 必须落在合法枚举内，否则 PowerPoint/WPS 会忽略该属性
    assert('T9 竖排 vert 合法', /vert="(horz|vert|vert270|wordArtVert|eaVert|mongolianVert|wordArtVertRtl)"/.test(slideXml[8]));
    // T10 对角线/内边距/样式
    // 对角线：a:lnTlToBr 本身就是线属性（CT_LineProperties），不能再嵌一层 a:ln
    assert('T10 对角线 lnTlToBr', /<a:lnTlToBr w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>/.test(slideXml[9]));
    assert('T10 对角线 lnBlToTr', /<a:lnBlToTr w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>/.test(slideXml[9]));
    assert('T10 对角线无双层 a:ln', !/<a:lnTlToBr><a:ln/.test(slideXml[9]));
    // 内边距必须是 a:tcPr 的属性（marL/marR/marT/marB），写成独立元素属于非法 OOXML
    assert('T10 单元格内边距 marL/marT', /<a:tcPr anchor="t" marL="190500" marR="190500" marT="95250" marB="95250"\/>/.test(slideXml[9]));
    assert('T10 无非法 tableCellInsets', !/tableCellInsets/.test(slideXml[9]));
    // 表格样式 ID 必须在 tableStyles.xml 中有对应定义，否则 WPS/PowerPoint 会退化成无网格
    assert('T10 表格样式 tableStyleId', /<a:tableStyleId>\{5C22544A-7EE6-4342-B048-85BDC9FD1C3A\}<\/a:tableStyleId>/.test(slideXml[9]));
    assert('T10 tableStyles 含样式定义', /styleId="\{5C22544A-7EE6-4342-B048-85BDC9FD1C3A\}"/.test(tableStylesXml));
    // T11 图表类型
    assert('T11 c:bar3DChart', /<c:bar3DChart/.test(chartAll));
    assert('T11 c:doughnutChart', /<c:doughnutChart/.test(chartAll));
    assert('T11 c:bubbleChart', /<c:bubbleChart/.test(chartAll));
    assert('T11 c:radarChart', /<c:radarChart/.test(chartAll));
    assert('T11 c:stockChart', /<c:stockChart/.test(chartAll));
    assert('T11 c:surfaceChart', /<c:surfaceChart/.test(chartAll));
    assert('T11 数据标签 dLbls', /<c:dLbls/.test(chartAll));
    assert('T11 图例位置 legendPos', /<c:legendPos/.test(chartAll));
    // T12 裁剪/调整
    assert('T12 裁剪 srcRect', /<a:srcRect/.test(slideXml[11]));
    assert('T12 调整 lum/alphaModFix', /<a:lum|<a:alphaModFix/.test(slideXml[11]));
    // T13 动画
    assert('T13 p:timing', /<p:timing/.test(slideXml[12]));
    assert('T13 动画目标 spTgt', /<p:spTgt/.test(slideXml[12]));
    // T14 切换/自动播放/隐藏
    assert('T14 隐藏 show="0"', /show="0"/.test(slideXml[13]));
    assert('T14 自动播放 afterTime', /afterTime/.test(slideXml[13]));
    assert('T14 切换 transition', /<p:transition/.test(slideXml[13]));
    // T15 视频/音频
    assert('T15 视频/音频节点', /videoFile|audioFile/.test(slideXml[14]));
    assert('T15 媒体部件 mp4/m4a', Object.keys(zip.files).some(f => /ppt\/media\/.+\.(mp4|m4a)$/.test(f)));
    // T16 SmartArt（部件编号按页内 diagramIndex，此处 data1.xml）
    assert('T16 diagrams/dataN.xml', Object.keys(zip.files).some(f => /^ppt\/diagrams\/data\d+\.xml$/.test(f)));
    assert('T16 graphicFrame+dgm:rel', /<p:graphicFrame/.test(slideXml[15]) && /dgm:rel/.test(slideXml[15]));
    // T17 自定义主题
    assert('T17 自定义主题色 6366F1', (await zip.file('ppt/theme/theme1.xml').async('string')).includes('6366F1'));
    // T18 批注（comments 部件按 slideIndex 命名，用正则匹配任意编号）
    const commentsFile = Object.keys(zip.files).find(f => /^ppt\/comments\/comments\d+\.xml$/.test(f));
    assert('T18 commentsN.xml', !!commentsFile);
    assert('T18 p:cm 批注', commentsFile ? /<p:cm /.test(await zip.file(commentsFile).async('string')) : false);
    assert('T18 commentAuthors.xml', !!zip.file('ppt/commentAuthors.xml'));
    // T19 自定义属性
    assert('T19 custom.xml', !!zip.file('docProps/custom.xml'));
    assert('T19 自定义属性内容', /研发/.test(await zip.file('docProps/custom.xml').async('string')));

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
