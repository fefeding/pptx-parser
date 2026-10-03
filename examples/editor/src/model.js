/**
 * 文档模型：内部编辑器模型 <-> PPTX 标准 JSON（PptxDocument）双向映射
 *
 * 内部模型与 PptxDocument 基本同构（坐标统一 px、字号 pt），
 * 额外补充编辑器需要的字段（id / locked / hidden / headerRow 等）。
 */
import { uid, clone, normalizeColor, clamp } from './util.js';

/* ======================= 常量 ======================= */
export const SLIDE_SIZES = {
  '16:9': { width: 1280, height: 720 },
  '16:10': { width: 1280, height: 800 },
  '4:3': { width: 1024, height: 768 }
};

export const FONT_LIST = [
  '微软雅黑', '黑体', '宋体', '楷体', '仿宋',
  'Arial', 'Helvetica', 'Times New Roman', 'Georgia', 'Verdana',
  'Tahoma', 'Calibri', 'Impact', 'Courier New'
];

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 56, 64, 72, 80, 96];

/** 主题：bg=页面底色 text=正文 title=标题 accent=主色 accents=图表配色 */
export const THEMES = [
  {
    id: 'blue', name: '简约蓝',
    bg: '#FFFFFF', panel: '#F8FAFC', text: '#202124', title: '#0B3D91',
    accent: '#1A73E8', accents: ['#1A73E8', '#4285F4', '#34A853', '#FBBC04', '#EA4335', '#8430CE'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  },
  {
    id: 'indigo', name: '靛紫星辰',
    bg: '#FFFFFF', panel: '#F1F0FF', text: '#1F1B2E', title: '#2B1E66',
    accent: '#5B4BDB', accents: ['#5B4BDB', '#8B7BF0', '#22D3EE', '#F472B6', '#FBBF24', '#34D399'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  },
  {
    id: 'green', name: '清新绿',
    bg: '#FFFFFF', panel: '#ECFDF5', text: '#0F2E24', title: '#065F46',
    accent: '#0F9D58', accents: ['#0F9D58', '#34D399', '#10B981', '#84CC16', '#F59E0B', '#0EA5E9'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  },
  {
    id: 'dark', name: '深夜蓝',
    bg: '#101828', panel: '#1D2939', text: '#E4E7EC', title: '#FFFFFF',
    accent: '#4F8DF7', accents: ['#4F8DF7', '#22D3EE', '#A78BFA', '#F472B6', '#FBBF24', '#34D399'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  },
  {
    id: 'warm', name: '暖阳橙',
    bg: '#FFFCF5', panel: '#FEF3E2', text: '#3B2A1A', title: '#8A4B08',
    accent: '#E8710A', accents: ['#E8710A', '#F59E0B', '#EF4444', '#8B5CF6', '#0EA5E9', '#10B981'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  },
  {
    id: 'rose', name: '玫瑰粉',
    bg: '#FFF7FA', panel: '#FCE8F0', text: '#3B1B2A', title: '#9D174D',
    accent: '#DB2777', accents: ['#DB2777', '#F472B6', '#A855F7', '#6366F1', '#F59E0B', '#14B8A6'],
    fonts: { major: '微软雅黑', minor: '微软雅黑' }
  }
];

export const getTheme = (id) => THEMES.find((t) => t.id === id) || THEMES[0];

/** 形状库：shapeType（OOXML prstGeom）+ 中文名 + 预览 SVG path */
export const SHAPES = [
  { type: 'rect', name: '矩形', d: 'M2 2h20v20H2z' },
  { type: 'roundRect', name: '圆角矩形', d: 'M6 2h12a4 4 0 0 1 4 4v12a4 4 0 0 1-4 4H6a4 4 0 0 1-4-4V6a4 4 0 0 1 4-4z' },
  { type: 'ellipse', name: '椭圆', d: 'M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20z' },
  { type: 'triangle', name: '三角形', d: 'M12 2l10 20H2z' },
  { type: 'rtTriangle', name: '直角三角形', d: 'M2 22V2h20z' },
  { type: 'diamond', name: '菱形', d: 'M12 2l10 10-10 10L2 12z' },
  { type: 'parallelogram', name: '平行四边形', d: 'M8 2h16l-8 20H0z' },
  { type: 'trapezoid', name: '梯形', d: 'M6 2h12l4 20H2z' },
  { type: 'pentagon', name: '五边形', d: 'M12 2l10 8-4 12H6L2 10z' },
  { type: 'hexagon', name: '六边形', d: 'M6 2h12l6 10-6 10H6L0 12z' },
  { type: 'octagon', name: '八边形', d: 'M8 2h8l6 6v8l-6 6H8l-6-6V8z' },
  { type: 'chevron', name: 'V 形', d: 'M2 2h9l4 10-4 10H2V12z' },
  { type: 'rightArrow', name: '右箭头', d: 'M2 6h14v-4l8 10-8 10v-4H2z' },
  { type: 'leftArrow', name: '左箭头', d: 'M22 6H8V2L0 12l8 10v-4h14z' },
  { type: 'upArrow', name: '上箭头', d: 'M12 2l10 9h-5v11H7v-11H2z' },
  { type: 'downArrow', name: '下箭头', d: 'M12 22L2 13h5V2h10v11h5z' },
  { type: 'pentagonBlock', name: '五角星块', d: 'M12 2l3 7h7l-6 5 2 8-6-4-6 4 2-8-6-5h7z' },
  { type: 'plus', name: '十字', d: 'M9 2h6v7h7v6h-7v7H9v-7H2V9h7z' },
  { type: 'heart', name: '心形', d: 'M12 22S2 15 2 8a5 5 0 0 1 10-2 5 5 0 0 1 10 2c0 7-10 14-10 14z' },
  { type: 'lightningBolt', name: '闪电', d: 'M14 2L4 14h6l-2 8 12-14h-7z' },
  { type: 'cloud', name: '云朵', d: 'M6 20a5 5 0 0 1 0-10 6 6 0 0 1 11-2 4.5 4.5 0 0 1 1 12z' },
  { type: 'moon', name: '月亮', d: 'M14 2a10 10 0 1 0 8 16A10 10 0 0 1 14 2z' },
  { type: 'sun', name: '太阳', d: 'M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10zm0-7l2 4h-4zm0 24l2-4h-4zM2 12l4-2v4zm20 0l-4 2v-4z' },
  { type: 'gear6', name: '齿轮', d: 'M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8zm0-8l1.5 3h-3zm0 24l1.5-3h-3zM2 12l3-1.5v3zm20 0l-3 1.5v-3z' },
  { type: 'donut', name: '圆环', d: 'M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm0 6a4 4 0 1 1 0 8 4 4 0 0 1 0-8z' },
  { type: 'pie', name: '扇形', d: 'M12 12V2a10 10 0 1 1-10 10z' },
  { type: 'arc', name: '弧形', d: 'M2 12a10 10 0 0 1 10-10v6a4 4 0 0 0-4 4z' },
  { type: 'cube', name: '立方体', d: 'M12 2l10 5v10l-10 5-10-5V7z' },
  { type: 'can', name: '圆柱', d: 'M6 4h12v16a5 5 0 0 1-12 0z' },
  { type: 'funnel', name: '漏斗', d: 'M2 2h20l-7 9v11h-6v-11z' },
  { type: 'frame', name: '边框', d: 'M2 2h20v20H2zm5 5v10h10V7z' },
  { type: 'foldedCorner', name: '折角', d: 'M2 2h14l6 6v14H2zm14 0v6h6' },
  { type: 'smileyFace', name: '笑脸', d: 'M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20zm-4 8h.01M16 10h.01M8 15a5 5 0 0 0 8 0' },
  { type: 'line', name: '直线', d: 'M2 12h20' }
];

/** 形状在渲染时对应的 CSS（clip-path / border-radius） */
export function shapeStyle(shapeType, w, h) {
  const r = Math.min(w, h) * 0.16;
  switch (shapeType) {
    case 'ellipse': return { borderRadius: '50%' };
    case 'roundRect': return { borderRadius: `${Math.min(r, 40)}px` };
    case 'triangle': return { clipPath: 'polygon(50% 0%,100% 100%,0% 100%)' };
    case 'rtTriangle': return { clipPath: 'polygon(0% 0%,100% 100%,0% 100%)' };
    case 'diamond': return { clipPath: 'polygon(50% 0%,100% 50%,50% 100%,0% 50%)' };
    case 'parallelogram': return { clipPath: 'polygon(20% 0%,100% 0%,80% 100%,0% 100%)' };
    case 'trapezoid': return { clipPath: 'polygon(20% 0%,80% 0%,100% 100%,0% 100%)' };
    case 'pentagon': return { clipPath: 'polygon(50% 0%,100% 38%,82% 100%,18% 100%,0% 38%)' };
    case 'hexagon': return { clipPath: 'polygon(25% 0%,75% 0%,100% 50%,75% 100%,25% 100%,0% 50%)' };
    case 'octagon': return { clipPath: 'polygon(30% 0%,70% 0%,100% 30%,100% 70%,70% 100%,30% 100%,0% 70%,0% 30%)' };
    case 'chevron': return { clipPath: 'polygon(0% 0%,60% 0%,100% 50%,60% 100%,0% 100%,40% 50%)' };
    case 'rightArrow': return { clipPath: 'polygon(0% 20%,60% 20%,60% 0%,100% 50%,60% 100%,60% 80%,0% 80%)' };
    case 'leftArrow': return { clipPath: 'polygon(100% 20%,40% 20%,40% 0%,0% 50%,40% 100%,40% 80%,100% 80%)' };
    case 'upArrow': return { clipPath: 'polygon(50% 0%,100% 100%,70% 100%,70% 60%,30% 60%,30% 100%,0% 100%)' };
    case 'downArrow': return { clipPath: 'polygon(50% 100%,0% 0%,30% 0%,30% 40%,70% 40%,70% 0%,100% 0%)' };
    case 'pentagonBlock': return { clipPath: 'polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)' };
    case 'plus': return { clipPath: 'polygon(35% 0%,65% 0%,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0% 65%,0% 35%,35% 35%)' };
    case 'heart': return { clipPath: 'polygon(50% 100%,0% 55%,0% 25%,25% 0%,50% 15%,75% 0%,100% 25%,100% 55%)' };
    case 'lightningBolt': return { clipPath: 'polygon(55% 0%,20% 55%,45% 55%,35% 100%,80% 40%,52% 40%)' };
    case 'cloud': return { borderRadius: '40% 40% 35% 35% / 45% 45% 55% 55%' };
    case 'moon': return { borderRadius: '50%' };
    case 'sun': return { borderRadius: '50%' };
    case 'gear6': return { borderRadius: '20%' };
    case 'donut': return { borderRadius: '50%' };
    case 'pie': return { clipPath: 'polygon(50% 50%,50% 0%,100% 15%,100% 50%)' };
    case 'arc': return { borderRadius: '50%' };
    case 'cube': return { clipPath: 'polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%)' };
    case 'can': return { borderRadius: '12% / 20%' };
    case 'funnel': return { clipPath: 'polygon(0% 0%,100% 0%,65% 55%,65% 100%,35% 100%,35% 55%)' };
    case 'frame': return { border: 'none', outline: '8px solid currentColor' };
    case 'foldedCorner': return { clipPath: 'polygon(0% 0%,70% 0%,100% 30%,100% 100%,0% 100%)' };
    case 'smileyFace': return { borderRadius: '50%' };
    default: return {};
  }
}

export const CHART_TYPES = [
  { value: 'barChart', name: '柱状图' },
  { value: 'barChart|bar', name: '条形图' },
  { value: 'barChart|stacked', name: '堆积柱状图' },
  { value: 'lineChart', name: '折线图' },
  { value: 'areaChart', name: '面积图' },
  { value: 'pieChart', name: '饼图' },
  { value: 'doughnutChart', name: '环形图' },
  { value: 'scatterChart', name: '散点图' },
  { value: 'radarChart', name: '雷达图' }
];

/* ======================= 元素工厂 ======================= */
function base(opts = {}) {
  return {
    id: uid(),
    x: opts.x ?? 120, y: opts.y ?? 120,
    width: opts.width ?? 400, height: opts.height ?? 120,
    rotation: opts.rotation ?? 0,
    locked: false, hidden: false,
    name: opts.name || ''
  };
}

export function createParagraph(text = '', style = {}) {
  return {
    runs: [{ text, ...style }],
    align: style.align,
    bullet: style.bullet,
    lineSpacing: style.lineSpacing
  };
}

export function createTextElement(opts = {}) {
  const el = base({ width: 460, height: 90, ...opts });
  el.type = 'text';
  el.name = opts.name || '文本框';
  const text = opts.text ?? '点击编辑文本';
  el.paragraphs = String(text).split('\n').map((t) => createParagraph(t));
  el.fontSize = opts.fontSize ?? 24;
  el.color = normalizeColor(opts.color) || '#202124';
  el.bold = !!opts.bold;
  el.italic = !!opts.italic;
  el.underline = !!opts.underline;
  el.fontFace = opts.fontFace || '微软雅黑';
  el.align = opts.align || 'left';
  el.valign = opts.valign || 'top';
  el.lineSpacing = opts.lineSpacing ?? 1.15;
  el.bullet = opts.bullet || false;
  el.indent = opts.indent ?? 0;
  return el;
}

export function createShapeElement(shapeType = 'rect', opts = {}) {
  const el = base({ width: 260, height: 180, ...opts });
  el.type = 'shape';
  el.shapeType = shapeType;
  el.name = opts.name || (SHAPES.find((s) => s.type === shapeType) || {}).name || '形状';
  el.fill = opts.fill ?? { type: 'solid', color: opts.color || '#4285F4', transparency: 0 };
  el.line = opts.line ?? 'none';
  el.shadow = opts.shadow ?? null;
  return el;
}

export function createImageElement(data, opts = {}) {
  const el = base({ width: 420, height: 300, ...opts });
  el.type = 'image';
  el.name = opts.name || '图片';
  el.data = data;
  el.imageAdjust = { brightness: 0, contrast: 0, transparency: 0 };
  return el;
}

export function createTableElement(rows = 3, cols = 3, opts = {}) {
  const el = base({ width: 720, height: 240, ...opts });
  el.type = 'table';
  el.name = opts.name || '表格';
  el.headerRow = true;
  el.border = { color: '#CBD5E1', width: 1 };
  el.headerFill = '#1A73E8';
  el.cellFill = '#FFFFFF';
  el.fontSize = 14;
  el.color = '#202124';
  el.bold = false;
  el.align = 'left';
  el.valign = 'middle';
  el.colWidths = new Array(cols).fill(Math.round(el.width / cols));
  el.rows = [];
  for (let r = 0; r < rows; r++) {
    const cells = [];
    for (let c = 0; c < cols; c++) {
      cells.push({ text: r === 0 ? `列 ${c + 1}` : '', fill: null, align: null, valign: null });
    }
    el.rows.push({ height: Math.round(el.height / rows), cells });
  }
  if (rows > 1) el.rows[1].cells[0].text = '内容';
  syncTableStyle(el);
  return el;
}

/** 表格样式同步：把表头/单元格样式写回每个 cell（导出与渲染都直接读 cell） */
export function syncTableStyle(el) {
  if (!el || el.type !== 'table') return;
  el.rows.forEach((row, ri) => {
    row.cells.forEach((cell) => {
      const isHead = el.headerRow && ri === 0;
      cell.fill = isHead ? (el.headerFill || null) : (cell.fillCustom ?? el.cellFill ?? null);
      if (isHead) {
        cell.bold = true;
        cell.color = cell.colorCustom ?? (el.headerTextColor ?? '#FFFFFF');
      } else {
        cell.bold = cell.boldCustom ?? !!el.bold;
        cell.color = cell.colorCustom ?? el.color ?? '#202124';
      }
      cell.fontSize = el.fontSize;
      cell.align = cell.align ?? el.align;
      cell.valign = cell.valign ?? el.valign;
      if (!isHead && cell.fillCustom === undefined) cell.fill = el.cellFill ?? null;
    });
  });
  return el;
}

export function createChartElement(chartType = 'barChart', opts = {}) {
  const el = base({ width: 640, height: 380, ...opts });
  el.type = 'chart';
  el.name = opts.name || '图表';
  el.chartType = chartType;
  el.title = opts.title ?? '';
  el.legend = opts.legend ?? true;
  el.dataLabels = opts.dataLabels ?? false;
  el.grouping = opts.grouping ?? 'clustered';
  el.holeSize = opts.holeSize ?? 50;
  el.smooth = opts.smooth ?? false;
  el.marker = opts.marker ?? true;
  el.categories = opts.categories ?? ['一月', '二月', '三月', '四月', '五月'];
  el.series = opts.series ?? [
    { name: '系列 1', values: [32, 45, 38, 56, 48] },
    { name: '系列 2', values: [22, 30, 42, 35, 51] }
  ];
  return el;
}

export function createGroupElement(children, opts = {}) {
  const el = base(opts);
  el.type = 'group';
  el.name = opts.name || '组合';
  el.children = children;
  return el;
}

export function createSlide(elements = [], opts = {}) {
  return {
    id: uid('s'),
    name: opts.name || '',
    background: opts.background ?? null,
    notes: opts.notes ?? '',
    hidden: !!opts.hidden,
    transition: opts.transition ?? null,
    animations: opts.animations ?? [],
    elements
  };
}

/* ======================= 过渡 / 动画常量 ======================= */
export const TRANSITIONS = [
  { value: 'none', name: '无' },
  { value: 'fade', name: '淡入淡出' },
  { value: 'wipe', name: '擦除' },
  { value: 'push', name: '推出' },
  { value: 'cover', name: '覆盖' },
  { value: 'blinds', name: '百叶窗' },
  { value: 'split', name: '分割' },
  { value: 'reveal', name: '显示' },
  { value: 'randomBar', name: '随机条' },
  { value: 'zoom', name: '缩放' },
  { value: 'fly', name: '飞入' },
];

export const TRANSITION_SPEEDS = [
  { value: 500, name: '快' },
  { value: 800, name: '中' },
  { value: 1500, name: '慢' },
];

export const ANIM_CLASSES = [
  { value: 'entr', name: '进入' },
  { value: 'exit', name: '退出' },
  { value: 'emph', name: '强调' },
];

export const ANIM_TYPES = {
  entr: [
    { value: 'flyIn', name: '飞入' },
    { value: 'fadeIn', name: '淡入' },
    { value: 'wipeIn', name: '擦除' },
    { value: 'zoomIn', name: '缩放' },
    { value: 'riseUp', name: '升起' },
    { value: 'bounceIn', name: '弹跳' },
  ],
  exit: [
    { value: 'flyOut', name: '飞出' },
    { value: 'fadeOut', name: '淡出' },
    { value: 'wipeOut', name: '擦除退出' },
    { value: 'zoomOut', name: '缩小退出' },
  ],
  emph: [
    { value: 'pulse', name: '脉冲' },
    { value: 'shake', name: '抖动' },
    { value: 'flash', name: '闪烁' },
    { value: 'grow', name: '放大' },
  ],
};

export const ANIM_DIRECTIONS = [
  { value: 'l', name: '← 左' },
  { value: 'r', name: '→ 右' },
  { value: 't', name: '↑ 上' },
  { value: 'b', name: '↓ 下' },
];

export const ANIM_TRIGGERS = [
  { value: 'onClick', name: '单击时' },
  { value: 'withPrev', name: '与上一动画同时' },
  { value: 'afterPrev', name: '上一动画之后' },
];

export function createDoc(themeId = 'blue', sizeKey = '16:9') {
  return {
    title: '未命名演示文稿',
    theme: themeId,
    slideSize: { ...SLIDE_SIZES[sizeKey] },
    slides: []
  };
}

/* ======================= 版式模板 ======================= */
export const LAYOUTS = [
  { id: 'title', name: '标题页', build: buildTitleLayout },
  { id: 'titleBody', name: '标题 + 正文', build: buildTitleBodyLayout },
  { id: 'titleOnly', name: '仅标题', build: buildTitleOnlyLayout },
  { id: 'section', name: '章节标题', build: buildSectionLayout },
  { id: 'twoCol', name: '两栏内容', build: buildTwoColLayout },
  { id: 'comparison', name: '对比', build: buildComparisonLayout },
  { id: 'quote', name: '引言', build: buildQuoteLayout },
  { id: 'imageText', name: '图文混排', build: buildImageTextLayout },
  { id: 'blank', name: '空白', build: () => [] }
];

function L(theme, size) {
  const W = size.width, H = size.height;
  const pad = Math.round(W * 0.06);
  return { W, H, pad, theme };
}

function buildTitleLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  const bar = createShapeElement('rect', {
    x: pad, y: Math.round(H * 0.52), width: Math.round(W * 0.18), height: 8,
    fill: { type: 'solid', color: theme.accent, transparency: 0 }
  });
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.3), width: W - pad * 2, height: Math.round(H * 0.2),
      text: '演示文稿标题', fontSize: Math.round(H * 0.075), bold: true, color: theme.title,
      valign: 'bottom', name: '标题'
    }),
    bar,
    createTextElement({
      x: pad, y: Math.round(H * 0.58), width: W - pad * 2, height: Math.round(H * 0.12),
      text: '副标题 · 演讲者与日期', fontSize: Math.round(H * 0.032), color: theme.text,
      name: '副标题'
    })
  ];
}

function buildTitleBodyLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.08), width: W - pad * 2, height: Math.round(H * 0.14),
      text: '页面标题', fontSize: Math.round(H * 0.055), bold: true, color: theme.title, name: '标题'
    }),
    createTextElement({
      x: pad, y: Math.round(H * 0.28), width: W - pad * 2, height: Math.round(H * 0.6),
      text: '· 要点一\n· 要点二\n· 要点三', fontSize: Math.round(H * 0.035),
      color: theme.text, bullet: true, lineSpacing: 1.5, name: '正文'
    })
  ];
}

function buildTitleOnlyLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.4), width: W - pad * 2, height: Math.round(H * 0.18),
      text: '页面标题', fontSize: Math.round(H * 0.06), bold: true, color: theme.title, name: '标题'
    })
  ];
}

function buildSectionLayout(theme, size) {
  const { W, H } = L(theme, size);
  return [
    createShapeElement('rect', {
      x: 0, y: 0, width: W, height: H,
      fill: { type: 'solid', color: theme.accent, transparency: 0 }
    }),
    createTextElement({
      x: Math.round(W * 0.08), y: Math.round(H * 0.38), width: Math.round(W * 0.84), height: Math.round(H * 0.2),
      text: '章节标题', fontSize: Math.round(H * 0.085), bold: true, color: '#FFFFFF', valign: 'middle', name: '章节标题'
    })
  ];
}

function buildTwoColLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  const colW = Math.round((W - pad * 3) / 2);
  const top = Math.round(H * 0.28);
  const colH = Math.round(H * 0.58);
  const mk = (x, title, body) => ([
    createTextElement({ x, y: top, width: colW, height: Math.round(H * 0.09), text: title, fontSize: Math.round(H * 0.038), bold: true, color: theme.accent }),
    createTextElement({ x, y: top + Math.round(H * 0.11), width: colW, height: colH - Math.round(H * 0.11), text: body, fontSize: Math.round(H * 0.03), color: theme.text, lineSpacing: 1.5 })
  ]);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.08), width: W - pad * 2, height: Math.round(H * 0.14),
      text: '页面标题', fontSize: Math.round(H * 0.055), bold: true, color: theme.title, name: '标题'
    }),
    ...mk(pad, '小标题一', '在此输入内容…'),
    ...mk(pad * 2 + colW, '小标题二', '在此输入内容…')
  ];
}

function buildComparisonLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  const colW = Math.round((W - pad * 3) / 2);
  const top = Math.round(H * 0.28);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.08), width: W - pad * 2, height: Math.round(H * 0.14),
      text: '对比分析', fontSize: Math.round(H * 0.055), bold: true, color: theme.title, name: '标题'
    }),
    createShapeElement('roundRect', { x: pad, y: top, width: colW, height: Math.round(H * 0.58), fill: { type: 'solid', color: theme.accents[0], transparency: 0 } }),
    createTextElement({ x: pad, y: top + 20, width: colW, height: 60, text: '方案 A', fontSize: Math.round(H * 0.04), bold: true, color: '#FFFFFF', align: 'center' }),
    createTextElement({ x: pad + 20, y: top + 100, width: colW - 40, height: Math.round(H * 0.4), text: '优势与说明…', fontSize: Math.round(H * 0.03), color: '#FFFFFF' }),
    createShapeElement('roundRect', { x: pad * 2 + colW, y: top, width: colW, height: Math.round(H * 0.58), fill: { type: 'solid', color: theme.accents[3], transparency: 0 } }),
    createTextElement({ x: pad * 2 + colW, y: top + 20, width: colW, height: 60, text: '方案 B', fontSize: Math.round(H * 0.04), bold: true, color: '#FFFFFF', align: 'center' }),
    createTextElement({ x: pad * 2 + colW + 20, y: top + 100, width: colW - 40, height: Math.round(H * 0.4), text: '优势与说明…', fontSize: Math.round(H * 0.03), color: '#FFFFFF' })
  ];
}

function buildQuoteLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.2), width: W - pad * 2, height: Math.round(H * 0.3),
      text: '"一句值得记住的话"', fontSize: Math.round(H * 0.06), color: theme.title, align: 'center', valign: 'middle', italic: true, name: '引言'
    }),
    createTextElement({
      x: pad, y: Math.round(H * 0.55), width: W - pad * 2, height: Math.round(H * 0.1),
      text: '—— 作者', fontSize: Math.round(H * 0.03), color: theme.text, align: 'center'
    })
  ];
}

function buildImageTextLayout(theme, size) {
  const { W, H, pad } = L(theme, size);
  const imgW = Math.round((W - pad * 3) * 0.45);
  return [
    createTextElement({
      x: pad, y: Math.round(H * 0.08), width: W - pad * 2, height: Math.round(H * 0.13),
      text: '页面标题', fontSize: Math.round(H * 0.055), bold: true, color: theme.title, name: '标题'
    }),
    createShapeElement('roundRect', {
      x: pad, y: Math.round(H * 0.28), width: imgW, height: Math.round(H * 0.58),
      fill: { type: 'solid', color: theme.panel, transparency: 0 },
      line: { color: theme.accent, width: 1, dashType: 'solid' }
    }),
    createTextElement({
      x: pad, y: Math.round(H * 0.28), width: imgW, height: Math.round(H * 0.58),
      text: '图片占位', fontSize: Math.round(H * 0.03), color: theme.text, align: 'center', valign: 'middle'
    }),
    createTextElement({
      x: pad * 2 + imgW, y: Math.round(H * 0.28), width: W - pad * 3 - imgW, height: Math.round(H * 0.58),
      text: '· 说明一\n· 说明二\n· 说明三', fontSize: Math.round(H * 0.035), color: theme.text, bullet: true, lineSpacing: 1.5
    })
  ];
}

export function layoutElements(layoutId, theme, slideSize) {
  const layout = LAYOUTS.find((l) => l.id === layoutId) || LAYOUTS[1];
  return layout.build(theme, slideSize);
}

/** 用版式+主题生成一整页 */
export function buildSlideFromLayout(layoutId, themeId, slideSize, extra = {}) {
  const theme = getTheme(themeId);
  const slide = createSlide(layoutElements(layoutId, theme, slideSize), {
    background: { type: 'solid', color: theme.bg }
  });
  Object.assign(slide, extra);
  return slide;
}

/* ======================= 内部模型 → PptxDocument ======================= */
function cleanRuns(paragraph, el) {
  const runs = (paragraph.runs || []).map((r) => {
    const run = { text: r.text == null ? '' : String(r.text) };
    if (r.fontSize != null && r.fontSize !== el.fontSize) run.fontSize = r.fontSize;
    if (r.color && normalizeColor(r.color) && normalizeColor(r.color) !== normalizeColor(el.color)) run.color = normalizeColor(r.color);
    if (r.bold && !el.bold) run.bold = true;
    if (r.italic && !el.italic) run.italic = true;
    if (r.underline && !el.underline) run.underline = true;
    if (r.fontFace && r.fontFace !== el.fontFace) run.fontFace = r.fontFace;
    return run;
  });
  return runs.length ? runs : [{ text: '' }];
}

function fillToPptx(fill) {
  if (!fill || fill === 'none') return 'none';
  if (typeof fill === 'string') return fill;
  if (fill.type === 'gradient') {
    return { type: 'gradient', direction: fill.direction || 'horizontal', stops: (fill.stops || []).map((s) => ({ color: s.color, position: s.position })) };
  }
  const out = { type: 'solid', color: normalizeColor(fill.color) || 'FFFFFF' };
  if (fill.transparency) out.transparency = fill.transparency;
  return out;
}

export function elementToPptx(el) {
  const out = {
    x: Math.round(el.x), y: Math.round(el.y),
    width: Math.round(el.width), height: Math.round(el.height)
  };
  if (el.rotation) out.rotation = Math.round(el.rotation * 100) / 100;
  if (el.name) out.name = el.name;

  switch (el.type) {
    case 'text': {
      out.type = 'text';
      if (el.valign) out.valign = el.valign;
      if (el.align) out.align = el.align;
      if (el.fontSize) out.fontSize = el.fontSize;
      if (el.color) out.color = normalizeColor(el.color) || undefined;
      if (el.bold) out.bold = true;
      if (el.italic) out.italic = true;
      if (el.underline) out.underline = true;
      if (el.fontFace) out.fontFace = el.fontFace;
      if (el.lineSpacing) out.lineSpacing = el.lineSpacing;
      if (el.bullet) out.bullet = el.bullet === 'number' ? 'number' : true;
      if (el.indent) out.indentLeft = el.indent;
      out.paragraphs = (el.paragraphs || []).map((p) => {
        const para = { runs: cleanRuns(p, el) };
        if (p.align) para.align = p.align;
        if (p.bullet) para.bullet = p.bullet === 'number' ? 'number' : { type: 'bullet' };
        if (p.lineSpacing) para.lineSpacing = p.lineSpacing;
        return para;
      });
      if (!out.paragraphs.length) out.paragraphs = [{ runs: [{ text: '' }] }];
      return out;
    }
    case 'shape': {
      out.type = 'shape';
      out.shapeType = el.shapeType || 'rect';
      out.fill = fillToPptx(el.fill);
      out.line = el.line && el.line !== 'none'
        ? { color: normalizeColor(el.line.color) || '000000', width: el.line.width || 1, dashType: el.line.dashType || 'solid' }
        : 'none';
      if (el.shadow) out.effects = { shadow: el.shadow };
      return out;
    }
    case 'image': {
      out.type = 'image';
      if (el.data) out.data = el.data;
      if (el.src) out.src = el.src;
      const adj = el.imageAdjust || {};
      if (adj.brightness || adj.contrast || adj.transparency) {
        out.imageAdjust = {
          brightness: adj.brightness || 0,
          contrast: adj.contrast || 0,
          transparency: adj.transparency || 0
        };
      }
      return out;
    }
    case 'table': {
      out.type = 'table';
      out.colWidths = (el.colWidths || []).map((w) => Math.round(w));
      if (el.border) out.border = { color: normalizeColor(el.border.color) || 'CBD5E1', width: el.border.width ?? 1 };
      out.rows = el.rows.map((row) => ({
        height: Math.round(row.height || 40),
        cells: row.cells.map((c) => {
          const cell = { text: c.text || '' };
          if (c.fill) cell.fill = normalizeColor(c.fill) || undefined;
          if (c.colSpan && c.colSpan > 1) cell.colSpan = c.colSpan;
          if (c.rowSpan && c.rowSpan > 1) cell.rowSpan = c.rowSpan;
          if (c.align) cell.align = c.align;
          if (c.valign) cell.valign = c.valign;
          if (c.fontSize) cell.fontSize = c.fontSize;
          if (c.color) cell.color = normalizeColor(c.color) || undefined;
          if (c.bold) cell.bold = true;
          return cell;
        })
      }));
      return out;
    }
    case 'chart': {
      out.type = 'chart';
      const [baseType, variant] = String(el.chartType).split('|');
      out.chartType = baseType;
      if (baseType === 'barChart') {
        out.barDir = variant === 'bar' ? 'bar' : 'col';
        if (variant === 'stacked') out.grouping = 'stacked';
      }
      if (el.title) out.title = el.title;
      out.legend = !!el.legend;
      out.dataLabels = !!el.dataLabels;
      out.categories = el.categories || [];
      out.series = (el.series || []).map((s) => {
        const ser = { name: s.name || '系列' };
        if (el.chartType.startsWith('scatterChart')) {
          ser.x = (s.values || []).map((_, i) => i + 1);
          ser.y = (s.values || []).map((v) => Number(v) || 0);
        } else {
          ser.values = (s.values || []).map((v) => Number(v) || 0);
        }
        if (s.color) ser.color = normalizeColor(s.color) || undefined;
        return ser;
      });
      if (/pie|doughnut/i.test(baseType)) out.varyColors = true;
      if (el.grouping && el.grouping !== 'clustered') out.grouping = el.grouping;
      if (baseType === 'doughnutChart') out.holeSize = el.holeSize ?? 50;
      if (el.smooth) out.smooth = true;
      if (el.marker) out.marker = true;
      return out;
    }
    case 'group': {
      out.type = 'group';
      out.childrenCoordinates = 'page';
      out.children = (el.children || []).filter((c) => !c.hidden).map(elementToPptx);
      return out;
    }
    default: {
      out.type = 'raw';
      if (el.__raw) { out.__raw = el.__raw; out.rawFallback = true; }
      return out;
    }
  }
}

export function slideToPptx(slide) {
  const visibleEls = (slide.elements || []).filter((e) => !e.hidden);
  const idToIndex = new Map();
  visibleEls.forEach((e, i) => idToIndex.set(e.id, i));
  const out = { elements: visibleEls.map(elementToPptx) };
  if (slide.background) out.background = slide.background;
  if (slide.notes) out.notes = slide.notes;
  if (slide.hidden) out.hidden = true;
  if (slide.transition && slide.transition.type && slide.transition.type !== 'none') {
    out.transition = {
      type: slide.transition.type,
      duration: slide.transition.duration || 800,
      advanceOnClick: slide.transition.advanceOnClick !== false
    };
  }
  if (slide.animations && slide.animations.length) {
    out.animations = slide.animations.map((a) => {
      const target = idToIndex.has(a.target) ? idToIndex.get(a.target) : (typeof a.target === 'number' ? a.target : 0);
      const o = { target, type: a.type, duration: a.duration || 0.5 };
      if (a.presetClass) o.presetClass = a.presetClass;
      if (a.direction) o.direction = a.direction;
      if (a.trigger) o.trigger = a.trigger;
      if (a.delay != null) o.delay = a.delay;
      if (a.repeat != null) o.repeat = a.repeat;
      return o;
    });
  }
  return out;
}

/** 主题对象 → 完整 theme1.xml（jsonToPptx 的 theme 只接受整串 XML） */
export function buildThemeXml(theme) {
  const t = theme || getTheme('blue');
  const accents = t.accents || [];
  const a = (i, fb) => (accents[i] || fb).replace('#', '');
  const bg = (t.bg || '#FFFFFF').replace('#', '');
  const tx = (t.text || '#202124').replace('#', '');
  const panel = (t.panel || '#F1F3F4').replace('#', '');
  const major = t.fonts?.major || '微软雅黑';
  const minor = t.fonts?.minor || '微软雅黑';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="EditorTheme"><a:themeElements><a:clrScheme name="Editor"><a:dk1><a:srgbClr val="${tx}"/></a:dk1><a:lt1><a:srgbClr val="${bg}"/></a:lt1><a:dk2><a:srgbClr val="${tx}"/></a:dk2><a:lt2><a:srgbClr val="${panel}"/></a:lt2><a:accent1><a:srgbClr val="${a(0, '1A73E8')}"/></a:accent1><a:accent2><a:srgbClr val="${a(1, '4285F4')}"/></a:accent2><a:accent3><a:srgbClr val="${a(2, '34A853')}"/></a:accent3><a:accent4><a:srgbClr val="${a(3, 'FBBC04')}"/></a:accent4><a:accent5><a:srgbClr val="${a(4, 'EA4335')}"/></a:accent5><a:accent6><a:srgbClr val="${a(5, '8430CE')}"/></a:accent6><a:hlink><a:srgbClr val="${a(0, '1A73E8')}"/></a:hlink><a:folHlink><a:srgbClr val="${a(5, '8430CE')}"/></a:folHlink></a:clrScheme><a:fontScheme name="Editor"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface="${major}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/><a:ea typeface="${minor}"/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Editor"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

export function docToPptx(doc) {
  const theme = getTheme(doc.theme);
  return {
    version: '1.0',
    slideSize: { width: doc.slideSize.width, height: doc.slideSize.height },
    metadata: {
      title: doc.title || '演示文稿',
      author: 'Slides 在线编辑器',
      lastModifiedBy: 'Slides 在线编辑器',
      modified: new Date().toISOString()
    },
    theme: buildThemeXml(theme),
    slides: doc.slides.map(slideToPptx)
  };
}

/* ======================= PptxDocument → 内部模型 ======================= */
function pptxRunsToRuns(runs, fallback) {
  if (!runs || !runs.length) return [{ text: '' }];
  return runs.map((r) => ({
    text: String(r.text == null ? '' : r.text),
    fontSize: r.fontSize ?? fallback.fontSize,
    color: normalizeColor(r.color) || fallback.color,
    bold: !!r.bold,
    italic: !!r.italic,
    underline: !!r.underline,
    fontFace: r.fontFace || fallback.fontFace
  }));
}

function fillFromPptx(fill) {
  if (!fill) return null;
  if (fill === 'none') return 'none';
  if (typeof fill === 'string') return { type: 'solid', color: normalizeColor(fill) || '#FFFFFF', transparency: 0 };
  if (fill.type === 'gradient') {
    return {
      type: 'gradient',
      direction: fill.direction || 'horizontal',
      stops: (fill.stops || []).map((s) => ({ color: normalizeColor(s.color) || '#FFFFFF', position: s.position ?? 0 }))
    };
  }
  return { type: 'solid', color: normalizeColor(fill.color) || '#FFFFFF', transparency: fill.transparency || 0 };
}

function elementFromPptx(pe) {
  const el = base({
    x: Number(pe.x) || 0, y: Number(pe.y) || 0,
    width: Number(pe.width) || 100, height: Number(pe.height) || 100,
    rotation: Number(pe.rotation) || 0,
    name: pe.name
  });
  switch (pe.type) {
    case 'text': {
      el.type = 'text';
      el.fontSize = pe.fontSize ?? 18;
      el.color = normalizeColor(pe.color) || '#202124';
      el.bold = !!pe.bold; el.italic = !!pe.italic; el.underline = !!pe.underline;
      el.fontFace = pe.fontFace || '微软雅黑';
      el.align = pe.align || 'left';
      el.valign = pe.valign || 'top';
      el.lineSpacing = typeof pe.lineSpacing === 'number' ? pe.lineSpacing : 1.15;
      el.bullet = pe.bullet === 'number' ? 'number' : !!pe.bullet;
      el.indent = pe.indentLeft || 0;
      const fallback = { fontSize: el.fontSize, color: el.color, fontFace: el.fontFace };
      if (pe.paragraphs && pe.paragraphs.length) {
        el.paragraphs = pe.paragraphs.map((p) => ({
          runs: p.runs ? pptxRunsToRuns(p.runs, fallback) : [{ text: String(p.text || '') }],
          align: p.align || el.align,
          bullet: p.bullet ? (p.bullet === 'number' ? 'number' : true) : undefined,
          lineSpacing: p.lineSpacing || el.lineSpacing
        }));
      } else if (pe.runs && pe.runs.length) {
        el.paragraphs = [{ runs: pptxRunsToRuns(pe.runs, fallback), align: el.align }];
      } else {
        el.paragraphs = String(pe.text || '').split('\n').map((t) => ({ runs: [{ text: t }], align: el.align }));
      }
      return el;
    }
    case 'shape': {
      el.type = 'shape';
      el.shapeType = pe.shapeType || 'rect';
      el.fill = fillFromPptx(pe.fill);
      el.line = pe.line && pe.line !== 'none'
        ? { color: normalizeColor(pe.line.color) || '#000000', width: pe.line.width ?? 1, dashType: pe.line.dashType || 'solid' }
        : 'none';
      el.shadow = (pe.effects && pe.effects.shadow && typeof pe.effects.shadow === 'object') ? pe.effects.shadow : null;
      return el;
    }
    case 'image': {
      el.type = 'image';
      el.data = pe.data || '';
      el.src = pe.src || '';
      el.imageAdjust = Object.assign({ brightness: 0, contrast: 0, transparency: 0 }, pe.imageAdjust || {});
      return el;
    }
    case 'table': {
      el.type = 'table';
      el.colWidths = (pe.colWidths && pe.colWidths.length) ? pe.colWidths.map(Number) : null;
      el.border = pe.border ? { color: normalizeColor(pe.border.color) || '#CBD5E1', width: pe.border.width ?? 1 } : { color: '#CBD5E1', width: 1 };
      el.rows = (pe.rows || []).map((r) => ({
        height: Number(r.height) || 40,
        cells: (r.cells || []).map((c) => ({
          text: c.text || (Array.isArray(c.paragraphs) ? c.paragraphs.map((p) => p.text || '').join('\n') : ''),
          fill: c.fill ? normalizeColor(c.fill) : null,
          align: c.align || 'left',
          valign: c.valign || 'middle',
          fontSize: c.fontSize || 14,
          color: normalizeColor(c.color) || '#202124',
          bold: !!c.bold,
          colSpan: c.colSpan, rowSpan: c.rowSpan
        }))
      }));
      if (!el.colWidths) {
        const cols = el.rows[0] ? el.rows[0].cells.length : 1;
        el.colWidths = new Array(cols).fill(Math.round(el.width / cols));
      }
      el.headerRow = false;
      el.headerFill = '#1A73E8';
      el.cellFill = '#FFFFFF';
      el.fontSize = el.rows[0]?.cells[0]?.fontSize || 14;
      el.color = '#202124';
      return el;
    }
    case 'chart': {
      el.type = 'chart';
      el.chartType = pe.chartType || 'barChart';
      el.title = pe.title || '';
      el.legend = pe.legend !== false;
      el.dataLabels = !!pe.dataLabels;
      el.grouping = pe.grouping || 'clustered';
      el.holeSize = pe.holeSize ?? 50;
      el.smooth = !!pe.smooth;
      el.marker = !!pe.marker;
      el.categories = pe.categories || [];
      el.series = (pe.series || []).map((s) => ({
        name: s.name || '系列',
        values: (s.values || s.y || []).map((v) => Number(v) || 0),
        color: s.color || ''
      }));
      return el;
    }
    case 'group': {
      el.type = 'group';
      const offsetX = Number(pe.x) || 0, offsetY = Number(pe.y) || 0;
      el.children = (pe.children || []).map((c) => {
        if (pe.childrenCoordinates === 'local') {
          c = clone(c);
          c.x = (Number(c.x) || 0) + offsetX;
          c.y = (Number(c.y) || 0) + offsetY;
        }
        return elementFromPptx(c);
      });
      return el;
    }
    default: {
      el.type = 'raw';
      el.label = pe.type || '未知元素';
      el.__raw = pe.__raw;
      return el;
    }
  }
}

export function docFromPptx(pptxDoc) {
  const doc = createDoc('blue');
  const size = pptxDoc.slideSize || {};
  doc.slideSize = { width: Number(size.width) || 1280, height: Number(size.height) || 720 };
  const key = Object.keys(SLIDE_SIZES).find((k) =>
    Math.abs(SLIDE_SIZES[k].width - doc.slideSize.width) < 4 && Math.abs(SLIDE_SIZES[k].height - doc.slideSize.height) < 4);
  doc.sizeKey = key || 'custom';
  doc.title = (pptxDoc.metadata && pptxDoc.metadata.title) || '导入的演示文稿';
  doc.slides = (pptxDoc.slides || []).map((s) => {
    const slide = createSlide((s.elements || []).map(elementFromPptx), {
      background: s.background || null,
      notes: s.notes || '',
      hidden: !!s.hidden
    });
    if (s.transition && s.transition.type) {
      slide.transition = { type: s.transition.type, duration: s.transition.duration || 800 };
    }
    if (s.animations && s.animations.length) {
      slide.animations = s.animations.map((a) => {
        const idx = typeof a.target === 'number' ? a.target : 0;
        const el = slide.elements[idx];
        return {
          target: el ? el.id : (slide.elements[0] ? slide.elements[0].id : ''),
          type: a.type, duration: a.duration || 0.5,
          presetClass: a.presetClass, direction: a.direction,
          trigger: a.trigger, delay: a.delay, repeat: a.repeat
        };
      });
    }
    return slide;
  });
  if (!doc.slides.length) doc.slides.push(buildSlideFromLayout('titleBody', doc.theme, doc.slideSize));
  return doc;
}

/* ======================= 文本样式批量应用 ======================= */
/** 把样式写入元素默认样式，并同步到所有 run（保证选中态读取一致） */
export function applyTextStyle(el, patch) {
  if (el.type !== 'text') return el;
  for (const [k, v] of Object.entries(patch)) {
    if (v === undefined) continue;
    el[k] = v;
  }
  const keys = ['fontSize', 'color', 'bold', 'italic', 'underline', 'fontFace'];
  for (const p of el.paragraphs || []) {
    for (const r of p.runs || []) {
      for (const k of keys) if (patch[k] !== undefined) r[k] = patch[k];
    }
    if (patch.align !== undefined) p.align = patch.align;
    if (patch.bullet !== undefined) p.bullet = patch.bullet === false ? undefined : patch.bullet;
    if (patch.lineSpacing !== undefined) p.lineSpacing = patch.lineSpacing;
  }
  return el;
}

/** 生成默认演示文稿（首次进入） */
export function createStarterDoc() {
  const doc = createDoc('blue', '16:9');
  const theme = getTheme('blue');
  doc.slides = [
    buildSlideFromLayout('title', 'blue', doc.slideSize),
    buildSlideFromLayout('titleBody', 'blue', doc.slideSize),
    buildSlideFromLayout('twoCol', 'blue', doc.slideSize)
  ];
  doc.slides[0].elements[0].paragraphs[0].runs[0].text = '欢迎使用 Slides 编辑器';
  doc.slides[0].elements[2].paragraphs[0].runs[0].text = '在线编辑 · 一键导出 PPTX';
  doc.title = '未命名演示文稿';
  return doc;
}
