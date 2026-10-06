/**
 * 文档模型：内部编辑器模型 <-> PPTX 标准 JSON（PptxDocument）双向映射
 *
 * 内部模型与 PptxDocument 基本同构（坐标统一 px、字号 pt），
 * 额外补充编辑器需要的字段（id / locked / hidden / headerRow 等）。
 */
import { uid, clone, normalizeColor, clamp, ptToPx, pxToPt } from './util.js';

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

export const getTheme = (id) => {
  if (id && typeof id === 'object') return id; // 已解析的主题对象（如导入 PPTX 的自定义主题）
  return THEMES.find((t) => t.id === id) || THEMES[0];
};

/** 当前解析 scheme 主题色引用时使用的主题对象（docFromPptx 在导入含自定义主题的 PPTX 时设置） */
let _activeTheme = 'blue';
export function setActiveTheme(t) { _activeTheme = t; }

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

/**
 * 音频元素（p:pic + p:nvPr/a:audioFile）。
 * poster 为占位封面（OOXML 里 a:blip 指向的图），缺省给一个喇叭图标，
 * 否则回写时 a:blip 会指向媒体关系（见 serializer 的 buildMediaElement）。
 */
export function createAudioElement(data, opts = {}) {
  const el = base({ width: 32, height: 32, ...opts });
  el.type = 'audio';
  el.name = opts.name || '音频';
  el.data = data || '';
  el.extension = opts.extension || (String(data || '').match(/data:audio\/([a-z0-9.+-]+)/i) || [])[1] || 'mp3';
  el.poster = opts.poster || { data: defaultMediaPoster('audio'), extension: 'svg' };
  return el;
}

/** 视频元素（p:pic + p:nvPr/a:videoFile），默认 16:9 */
export function createVideoElement(data, opts = {}) {
  const el = base({ width: 480, height: 270, ...opts });
  el.type = 'video';
  el.name = opts.name || '视频';
  el.data = data || '';
  el.extension = opts.extension || (String(data || '').match(/data:video\/([a-z0-9.+-]+)/i) || [])[1] || 'mp4';
  el.poster = opts.poster || { data: defaultMediaPoster('video'), extension: 'svg' };
  return el;
}

/** 内置占位封面（SVG data URL），避免把二进制图标塞进文档数据 */
export function defaultMediaPoster(kind) {
  const isVideo = kind === 'video';
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
    + `<rect width="64" height="64" rx="8" fill="${isVideo ? '#1F2937' : '#334155'}"/>`
    + `<text x="32" y="41" font-family="sans-serif" font-size="26" fill="#fff" text-anchor="middle">${isVideo ? '▶' : '♪'}</text>`
    + '</svg>';
  return 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
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
      // fontSizeSet 为假表示源文件未显式指定字号（靠主题/表格样式继承）；
      // 只有用户在属性面板改过字号（面板会置 fontSizeSet=true）才回写到每个单元格
      if (el.fontSizeSet) { cell.fontSize = el.fontSize; cell.fontSizeSet = true; }
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
/** 自动编号前缀（单元格纯文本用）：支持 hebrew 与常见数字/字母编号 */
const HEBREW_LETTERS = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ', 'ק', 'ר', 'ש', 'ת'];
function autoNumPrefix(fmt, n) {
  if (fmt === 'hebrew1Minus' || fmt === 'hebrew2Minus') {
    let out = '', rest = Math.max(0, n - 1);
    do { out = HEBREW_LETTERS[rest % HEBREW_LETTERS.length] + out; rest = Math.floor(rest / HEBREW_LETTERS.length) - 1; } while (rest >= 0);
    return (out || HEBREW_LETTERS[0]) + '-';
  }
  return `${n}.`;
}

/**
 * 表格单元格原始段落 → 生成端可识别的段落结构。
 * 保留 runs 样式与 bullet（含自动编号 fmt/start），供未被编辑的单元格原样回写。
 */
function cleanCellParagraphs(paragraphs) {
  return (paragraphs || []).map((p) => {
    const out = { runs: (p.runs || []).map((r) => {
      const run = { text: r.text == null ? '' : String(r.text) };
      if (r.fontSize != null) run.fontSize = r.fontSize;
      if (r.color) run.color = normalizeColor(r.color) || undefined;
      if (r.bold) run.bold = true;
      if (r.italic) run.italic = true;
      if (r.underline) run.underline = true;
      if (r.fontFace) run.fontFace = r.fontFace;
      if (r.href) run.href = r.href;
      if (r.break) run.break = true;
      return run;
    }) };
    if (!out.runs.length) out.runs = [{ text: '' }];
    if (p.align) out.align = p.align;
    if (p.rtl) out.rtl = true;
    // 段落级缩进/边距：表格单元格段落也需保留 marL/marR/indent，否则列表缩进丢失
    if (p.indentLeft != null) out.indentLeft = p.indentLeft;
    if (p.indentRight != null) out.indentRight = p.indentRight;
    if (p.indent != null) out.indent = p.indent;
    // 自动编号：fmt（hebrew2Minus 等）与 startAt 必须原样保留，否则导出后退化为阿拉伯数字
    const b = p.bullet;
    if (b && typeof b === 'object' && b.type === 'number') {
      out.bullet = { type: 'number', fmt: b.fmt || 'arabicPeriod', start: b.start ?? 1 };
    } else if (b && typeof b === 'object' && b.type === 'bullet') {
      out.bullet = { type: 'bullet', char: b.char };
      if (b.font) out.bullet.font = b.font;
      if (b.sizePct) out.bullet.sizePct = b.sizePct;
    } else if (b === 'number') {
      out.bullet = 'number';
    } else if (b === true || b === 'bullet') {
      out.bullet = true;
    }
    return out;
  });
}

function cleanRuns(paragraph, el) {
  const runs = (paragraph.runs || []).map((r) => {
    const run = { text: r.text == null ? '' : String(r.text) };
    if (r.fontSize != null && r.fontSize !== el.fontSize) run.fontSize = r.fontSize;
    if (r.color && normalizeColor(r.color) && normalizeColor(r.color) !== normalizeColor(el.color)) run.color = normalizeColor(r.color);
    if (r.bold && !el.bold) run.bold = true;
    if (r.italic && !el.italic) run.italic = true;
    if (r.underline && !el.underline) run.underline = true;
    if (r.fontFace && r.fontFace !== el.fontFace) run.fontFace = r.fontFace;
    if (r.outline) run.outline = r.outline;
    if (r.shadow) run.shadow = r.shadow;
    if (r.href) { run.href = r.href; if (r.hrefTooltip) run.hrefTooltip = r.hrefTooltip; }
    if (r.break) run.break = true;
    return run;
  });
  return runs.length ? runs : [{ text: '' }];
}

/**
 * 逐点填充项归一化：纯色字符串加 '#'，渐变对象原样透传。
 * 3D 饼/环的 c:dPt 多为径向渐变，压成纯色会让渲染端退回默认调色板。
 */
function normPointColor(c) {
  if (!c) return undefined;
  if (typeof c === 'object') return c;
  return normalizeColor(c) || undefined;
}

function fillToPptx(fill) {
  // null/undefined = 源文件未指定填充（继承主题/版式），导出时不能写成 'none'（会变成显式无填充，阻断继承）
  if (fill === 'none') return 'none';
  if (!fill) return undefined;
  if (typeof fill === 'string') return fill;
  if (fill.type === 'gradient') {
    const g = { type: 'gradient', direction: fill.direction || 'horizontal', stops: (fill.stops || []).map((s) => ({ color: s.color, position: s.position })) };
    // 径向渐变（a:path）：缺了会退化成水平线性
    if (fill.gradientType === 'radial') { g.gradientType = 'radial'; g.gradientPath = fill.gradientPath || 'circle'; }
    return g;
  }
  if (fill.type === 'pattern') {
    const out = { type: 'pattern', prst: fill.prst || 'pct10', fg: normalizeColor(fill.fg) || '1A73E8', bg: normalizeColor(fill.bg) || 'FFFFFF' };
    return out;
  }
  if (fill.type === 'image') {
    const out = { type: 'image', extension: fill.extension || 'png' };
    if (fill.data) out.data = fill.data;
    if (fill.src) out.src = fill.src;
    if (fill.tile) out.tile = fill.tile;
    if (fill.srcRect) out.srcRect = fill.srcRect;
    return out;
  }
  const out = { type: 'solid', color: normalizeColor(fill.color) || 'FFFFFF' };
  if (fill.transparency) out.transparency = fill.transparency;
  return out;
}

export function elementToPptx(el) {
  // 保留 2 位小数：整数 px 取整会让每页所有元素位移最多 0.5px（≈0.2mm），
  // 与原始 PPTX 的 EMU 坐标产生可见偏差（导出后排版整体微移）
  const r2 = (v) => Math.round((Number(v) || 0) * 100) / 100;
  const out = {
    x: r2(el.x), y: r2(el.y),
    width: r2(el.width), height: r2(el.height)
  };
  if (el.rotation) out.rotation = Math.round(el.rotation * 100) / 100;
  if (el.name) out.name = el.name;

  switch (el.type) {
    case 'text': {
      out.type = 'text';
      // 只回写「源文件显式指定过」或「用户在属性面板改过」的样式，
      // 其余保持继承（propsSet 缺失时按旧行为全量回写，兼容新建元素与旧 JSON）
      const ps = el.propsSet;
      const ok = (k) => !ps || ps[k];
      if (ok('valign') && el.valign) out.valign = el.valign;
      if (ok('align') && el.align) out.align = el.align;
      if (ok('fontSize') && el.fontSize) out.fontSize = el.fontSize;
      if (el.color) out.color = normalizeColor(el.color) || undefined;
      if (el.bold) out.bold = true;
      if (el.italic) out.italic = true;
      if (el.underline) out.underline = true;
      if (ok('fontFace') && el.fontFace) out.fontFace = el.fontFace;
      if (ok('lineSpacing') && el.lineSpacing) out.lineSpacing = el.lineSpacing;
      if (el.bullet) out.bullet = el.bullet === 'number' ? 'number' : true;
      if (el.indent) out.indentLeft = el.indent;
      if (el.textDirection) out.textDirection = el.textDirection;
      if (el.inset) out.inset = el.inset;
      if (el.rtlCol) out.rtlCol = true;
      // 写回几何与外观：带文字的形状底，或带填充/边框的纯文本框
      if (el.shapeType) {
        out.shapeType = el.shapeType;
        if (el.adjust && typeof el.adjust === 'object') out.adjust = el.adjust;
      }
      if (el.custGeom && Array.isArray(el.custGeom.paths) && el.custGeom.paths.length) {
        out.custGeom = el.custGeom;
      }
      const f = fillToPptx(el.fill);
      if (f !== undefined) out.fill = f;
      if (el.line) {
        out.line = el.line !== 'none'
          ? { color: normalizeColor(el.line.color) || '000000', width: el.line.width || 1, dashType: el.line.dashType || 'solid' }
          : 'none';
      }
      if (el.shadow || el.glow) {
        out.effects = {};
        if (el.shadow) out.effects.shadow = el.shadow;
        if (el.glow) out.effects.glow = el.glow;
      }
      if (el.noWrap) out.noWrap = true;
      if (el.rtlCol) out.rtlCol = true;
      out.paragraphs = (el.paragraphs || []).map((p) => {
        const para = { runs: cleanRuns(p, el) };
        // 同理于元素级：只回写源段落显式写过、或用户改过的样式（继承值不固化）
        if ((p.alignSet !== undefined ? p.alignSet : !!p.align) && p.align) para.align = p.align;
        if (p.rtl) para.rtl = true;
        const pb = normalizeBulletOut(p.bullet);
        if (pb) para.bullet = pb;
        // 行距同上：源段落未写且用户未改时保持继承，避免把面板默认 1.15 固化进 XML
        if ((p.lineSpacingSet !== undefined ? p.lineSpacingSet : !!p.lineSpacing) && p.lineSpacing) para.lineSpacing = p.lineSpacing;
        if (p.indent != null) para.indent = pxToPt(p.indent);
        if (p.indentLeft != null) para.indentLeft = pxToPt(p.indentLeft);
        if (p.indentRight != null) para.indentRight = pxToPt(p.indentRight);
        if (p.spaceBefore != null) para.spaceBefore = pxToPt(p.spaceBefore);
        if (p.spaceAfter != null) para.spaceAfter = pxToPt(p.spaceAfter);
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
      if (el.shadow || el.glow) {
        out.effects = {};
        if (el.shadow) out.effects.shadow = el.shadow;
        if (el.glow) out.effects.glow = el.glow;
      }
      if (el.flipH) out.flipH = true;
      if (el.flipV) out.flipV = true;
      if (el.adjust && typeof el.adjust === 'object' && Object.keys(el.adjust).length) out.adjust = el.adjust;
      if (el.custGeom && Array.isArray(el.custGeom.paths) && el.custGeom.paths.length) out.custGeom = el.custGeom;
      return out;
    }
    case 'image': {
      out.type = 'image';
      if (el.data) out.data = el.data;
      if (el.src) out.src = el.src;
      if (el.extension) out.extension = el.extension;
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
      // 与坐标同理：列宽取整会让各列累计偏差，表格总宽与 gridCol 不再吻合
      out.colWidths = (el.colWidths || []).map((w) => Math.round((Number(w) || 0) * 100) / 100);
      if (el.border) out.border = { color: normalizeColor(el.border.color) || 'CBD5E1', width: el.border.width ?? 1 };
      if (el.inset) out.inset = { l: el.inset.l, r: el.inset.r, t: el.inset.t, b: el.inset.b };
      if (el.tableStyleId) out.tableStyleId = el.tableStyleId;
      out.rows = el.rows.map((row) => ({
        // 与坐标/列宽同理：取整 px 会引入 EMU 舍入误差（37.76px→38px→361950EMU vs 原始 359759EMU）
        height: Math.round((Number(row.height) || 40) * 100) / 100,
        cells: row.cells.map((c) => {
          const cell = {};
          // 未编辑的单元格：回写原始段落结构（保留 RTL 对齐与自动编号格式 hebrew2Minus 等），
          // 否则展平后的纯文本会丢掉编号语义，导出后编号退化成 arabicPeriod
          if (c.paragraphs && c.paragraphs.length && c.text === c.paragraphsText) {
            cell.paragraphs = cleanCellParagraphs(c.paragraphs);
          } else {
            cell.text = c.text || '';
          }
          if (c.fill) cell.fill = normalizeColor(c.fill) || undefined;
          if (c.colSpan && c.colSpan > 1) cell.colSpan = c.colSpan;
          if (c.rowSpan && c.rowSpan > 1) cell.rowSpan = c.rowSpan;
          // 仅当源文件显式写过、或用户在属性面板改过时才回写，避免把面板默认值当成显式值导出
          if (c.rtl) cell.rtl = true;
          if (c.alignSet) cell.align = c.align;
          if (c.valignSet) cell.valign = c.valign;
          if (c.fontSizeSet) cell.fontSize = c.fontSize;
          if (c.color) cell.color = normalizeColor(c.color) || undefined;
          if (c.bold) cell.bold = true;
          if (c.inset) cell.inset = { l: c.inset.l, r: c.inset.r, t: c.inset.t, b: c.inset.b };
          if (c.borders) {
            // 遍历全部边（含 insideH/insideV 内部网格线）：只处理四边会让内部网格线丢色，
            // 生成端回退成默认黑色，白底表格会变成黑网格
            const sides = {};
            for (const [k, b] of Object.entries(c.borders)) {
              if (b === 'none') sides[k] = 'none';
              else if (typeof b === 'string') sides[k] = b; // diagonal 是方向串（tlBr/blTr/both）
              else if (b && typeof b === 'object') sides[k] = { color: normalizeColor(b.color) || '#000000', width: b.width ?? 1 };
            }
            if (Object.keys(sides).length) cell.borders = sides;
          }
          return cell;
        })
      }));
      return out;
    }
    case 'chart': {
      out.type = 'chart';
      const [baseType, variant] = String(el.chartType).split('|');
      // 曾为 3D 类型的图表，导出时还原原始类型
      out.chartType = el.chartType3D || baseType;
      if (baseType === 'barChart') {
        out.barDir = variant === 'bar' ? 'bar' : 'col';
        if (variant === 'stacked') out.grouping = 'stacked';
      }
      if (el.title) out.title = el.title;
      out.legend = !!el.legend;
      if (el.legendPosition) out.legendPosition = el.legendPosition;
      out.dataLabels = !!el.dataLabels;
      out.categories = el.categories || [];
      const isBubble = el.chartType3D === 'bubbleChart';
      const isScatter = !isBubble && el.chartType === 'scatterChart';
      const isStock = el.chartType3D === 'stockChart';
      out.series = (el.series || []).map((s) => {
        const ser = { name: s.name || '系列' };
        if (isBubble) {
          ser.x = (s.values || []).map((v) => v && v.x);
          ser.y = (s.values || []).map((v) => v && v.y);
          ser.values = (s.values || []).map((v) => v && v.size);
        } else if (isScatter) {
          ser.x = (s.values || []).map((v) => v && v.x);
          ser.y = (s.values || []).map((v) => v && v.y);
        } else if (isStock) {
          ser.open = (s.values || []).map((v) => v && v[0]);
          ser.close = (s.values || []).map((v) => v && v[1]);
          ser.low = (s.values || []).map((v) => v && v[2]);
          ser.high = (s.values || []).map((v) => v && v[3]);
        } else {
          ser.values = (s.values || []).map((v) => Number(v) || 0);
        }
        if (s.color) ser.color = normalizeColor(s.color) || undefined;
        // 逐点填充（c:dPt 回写）：纯色走 normalizeColor，渐变对象原样透传
        if (s.pointColors && s.pointColors.some((c) => c)) {
          ser.pointColors = (s.pointColors || []).map(normPointColor);
        }
        return ser;
      });
      if (el.spaceFill) out.spaceFill = el.spaceFill;
      if (/pie|doughnut/i.test(baseType)) out.varyColors = true;
      if (el.grouping && el.grouping !== 'clustered') out.grouping = el.grouping;
      if (baseType === 'doughnutChart') out.holeSize = el.holeSize ?? 50;
      if (el.smooth) out.smooth = true;
      if (el.marker) out.marker = true;
      return out;
    }
    case 'group': {
      out.type = 'group';
      // 导入的组合子元素为相对坐标，保持 'relative'；编辑器创建的组合为页面绝对坐标 'page'
      out.childrenCoordinates = el.childrenCoordinates === 'relative' ? 'relative' : 'page';
      out.children = (el.children || []).filter((c) => !c.hidden).map(elementToPptx);
      return out;
    }
    case 'video':
    case 'audio': {
      out.type = el.type;
      if (el.data) out.data = el.data;
      if (el.src) out.src = el.src;
      if (el.extension) out.extension = el.extension;
      if (el.poster) out.poster = el.poster;
      return out;
    }
    case 'diagram': {
      out.type = 'diagram';
      out.diagramType = el.diagramType || 'list';
      out.nodes = (el.texts || []).map((t) => ({ text: t }));
      // 缓存绘图形状（树形布局），保留以便导出 JSON 后仍能还原树形展示
      if (Array.isArray(el.shapes) && el.shapes.length) out.shapes = clone(el.shapes);
      // 透传原始 SmartArt 部件，序列化器据此重建 data1/colors1/quickStyle1/layout1/drawing1
      if (el.__raw) out.__raw = el.__raw;
      if (el.dataPath) out.dataPath = el.dataPath;
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
  // 批注：pos 由 px 转回 EMU（标准 JSON 的批注锚点为 EMU）
  if (Array.isArray(slide.comments) && slide.comments.length) {
    out.comments = slide.comments.map((c) => ({
      author: c.author || 'Author',
      text: c.text || '',
      dt: c.dt || undefined,
      pos: c.pos ? { x: Math.round((Number(c.pos.x) || 0) * 9525), y: Math.round((Number(c.pos.y) || 0) * 9525) } : undefined
    }));
  }
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

/** 主题对象 → 完整 theme1.xml（jsonToPptx 的 theme 只接受整串 XML）；字符串则原样透传（解析端回读的原始 theme1.xml） */
export function buildThemeXml(theme) {
  if (typeof theme === 'string') return theme;
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
  const out = {
    version: '1.0',
    slideSize: { width: doc.slideSize.width, height: doc.slideSize.height },
    // 保留原始元数据（title/author/subject/keywords 等），仅追加修改者标记
    metadata: doc.metadata || {
      title: doc.title || '演示文稿',
      author: 'Slides 在线编辑器',
      lastModifiedBy: 'Slides 在线编辑器',
      modified: new Date().toISOString()
    },
    theme: doc.themeXml || buildThemeXml(theme),
    slides: doc.slides.map(slideToPptx)
  };
  // 自定义属性透传
  if (doc.customProps) out.customProps = doc.customProps;
  // 缩略图透传
  if (doc.thumbnail) out.thumbnail = doc.thumbnail;
  // 多主题/母版/版式无损回退：与 themeXml 同理，导入时保留的原结构直接写回
  if (doc.tableStylesXml) out.tableStylesXml = doc.tableStylesXml;
  if (doc.themeXmls && doc.themeXmls.length) out.themeXmls = doc.themeXmls;
  if (doc.masters && doc.masters.length) {
    out.masters = doc.masters;
    out.slides.forEach((sl, i) => { if (doc.slides[i] && doc.slides[i].layout != null) sl.layout = doc.slides[i].layout; });
  }
  return out;
}

/* ======================= PptxDocument → 内部模型 ======================= */
function pptxRunsToRuns(runs, fallback) {
  if (!runs || !runs.length) return [{ text: '' }];
  return runs.map((r) => ({
    text: String(r.text == null ? '' : r.text),
    fontSize: r.fontSize ?? fallback.fontSize,
    color: colorFromPptx(r.color) || fallback.color,
    bold: !!r.bold,
    italic: !!r.italic,
    underline: !!r.underline,
    fontFace: r.fontFace || fallback.fontFace,
    outline: r.outline,
    shadow: r.shadow,
    href: r.href || undefined,
    hrefTooltip: r.hrefTooltip || undefined,
    // 软换行（a:br）：空文本 run，渲染为 <br>
    break: !!r.break
  }));
}

/** 解析 'scheme:<name>' 主题色引用 → 编辑器当前主题（导入 PPTX 的自定义主题或默认蓝主题）下的具体颜色 */
function resolveSchemeColor(c) {
  if (typeof c !== 'string' || !c.startsWith('scheme:')) return c;
  const t = getTheme(_activeTheme);
  const map = {
    accent1: t.accents[0], accent2: t.accents[1], accent3: t.accents[2],
    accent4: t.accents[3], accent5: t.accents[4], accent6: t.accents[5],
    dk1: t.text, lt1: t.bg, dk2: t.title, lt2: t.panel,
    tx1: t.text, bg1: t.bg, tx2: t.title, bg2: t.panel,
    hlink: t.accents[0], folHlink: t.accents[5]
  };
  return map[c.slice(7)] || '#202124';
}
/** PPTX 颜色 → 编辑器颜色（兼容 scheme 引用 / 裸 hex / #hex） */
function colorFromPptx(c) {
  return normalizeColor(resolveSchemeColor(c));
}

/**
 * 标准 JSON bullet → 编辑器内部 bullet 表示。
 * 保留自动编号格式（fmt：chineseCounting 等）与项目符号字符（char），
 * 布尔值用于常见无格式场景。
 */
export function normalizeBulletIn(b) {
  if (b == null || b === false) return undefined;
  if (b === 'number' || (typeof b === 'object' && b.type === 'number' && !b.fmt)) return 'number';
  if (b === true || b === 'bullet') return true;
  if (typeof b === 'object') {
    if (b.type === 'number') return { type: 'number', fmt: b.fmt || 'arabicPeriod', start: b.start || 1 };
    // 图片项目符号（a:buBlp）：保留 data 与字号比例，否则图片符号整段丢失
    if (b.type === 'picture') {
      const pb = { type: 'picture' };
      if (b.data) pb.data = b.data;
      if (b.rid) pb.rid = b.rid;
      if (b.sizePct) pb.sizePct = b.sizePct;
      return pb;
    }
    // 字符符号：保留符号字体（a:buFont）与字号比例（a:buSzPct），
    // 否则 Wingdings 之类会退化成普通字形
    const out = { type: 'bullet', char: b.char || '•' };
    if (b.font) out.font = b.font;
    if (b.sizePct) out.sizePct = b.sizePct;
    return out;
  }
  return true;
}

/** 编辑器内部 bullet → 标准 JSON bullet（保留 fmt/char/font/sizePct/图片） */
export function normalizeBulletOut(b) {
  if (b == null || b === false) return undefined;
  if (b === 'number') return 'number';
  if (b === true || b === 'bullet') return { type: 'bullet' };
  if (typeof b === 'object' && b.type === 'number') return { type: 'number', fmt: b.fmt, start: b.start };
  if (typeof b === 'object' && b.type === 'picture') {
    const ob = { type: 'picture' };
    if (b.data) ob.data = b.data;
    if (b.rid) ob.rid = b.rid;
    if (b.sizePct) ob.sizePct = b.sizePct;
    return ob;
  }
  if (typeof b === 'object' && b.char) {
    const ob = { type: 'bullet', char: b.char };
    if (b.font) ob.font = b.font;
    if (b.sizePct) ob.sizePct = b.sizePct;
    return ob;
  }
  return { type: 'bullet' };
}

function fillFromPptx(fill) {
  if (!fill) return null;
  if (fill === 'none') return 'none';
  if (typeof fill === 'string') return { type: 'solid', color: colorFromPptx(fill) || '#FFFFFF', transparency: 0 };
  if (fill.type === 'gradient') {
    const g = {
      type: 'gradient',
      direction: fill.direction || 'horizontal',
      stops: (fill.stops || []).map((s) => ({ color: colorFromPptx(s.color) || '#FFFFFF', position: s.position ?? 0 }))
    };
    if (fill.gradientType === 'radial') { g.gradientType = 'radial'; g.gradientPath = fill.gradientPath || 'circle'; }
    return g;
  }
  if (fill.type === 'pattern') {
    return { type: 'pattern', prst: fill.prst || 'pct10', fg: colorFromPptx(fill.fg) || '#1A73E8', bg: colorFromPptx(fill.bg) || '#FFFFFF' };
  }
  if (fill.type === 'image') {
    return { type: 'image', extension: fill.extension || 'png', data: fill.data || '', src: fill.src || '', tile: fill.tile || null, srcRect: fill.srcRect || null };
  }
  return { type: 'solid', color: colorFromPptx(fill.color) || '#FFFFFF', transparency: fill.transparency || 0 };
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
      el.color = colorFromPptx(pe.color) || '#202124';
      el.bold = !!pe.bold; el.italic = !!pe.italic; el.underline = !!pe.underline;
      el.fontFace = pe.fontFace || '微软雅黑';
      el.align = pe.align || 'left';
      el.valign = pe.valign || 'top';
      el.textDirection = pe.textDirection || '';
      el.noWrap = !!pe.noWrap;
      el.rtlCol = !!pe.rtlCol;
      el.inset = (pe.inset && typeof pe.inset === 'object') ? {
        l: pe.inset.l ?? 7.2, r: pe.inset.r ?? 7.2, t: pe.inset.t ?? 3.6, b: pe.inset.b ?? 3.6
      } : null;
      // 带文字的形状底（椭圆/饼图/弧线等）：保留几何与外观
      el.shapeType = pe.shapeType || null;
      if (pe.adjust && typeof pe.adjust === 'object') el.adjust = pe.adjust;
      if (pe.custGeom && Array.isArray(pe.custGeom.paths) && pe.custGeom.paths.length) el.custGeom = pe.custGeom;
      el.fill = fillFromPptx(pe.fill);
      el.line = pe.line && pe.line !== 'none'
        ? { color: colorFromPptx(pe.line.color) || '#000000', width: pe.line.width ?? 1, dashType: pe.line.dashType || 'solid' }
        : (pe.line === 'none' ? 'none' : null);
      // 带文字形状的阴影/发光（来自 spPr/a:effectLst 或 p:style/a:effectRef 主题样式）
      el.shadow = (pe.effects && pe.effects.shadow && typeof pe.effects.shadow === 'object') ? pe.effects.shadow : null;
      if (el.shadow && el.shadow.color) el.shadow.color = colorFromPptx(el.shadow.color) || '#000000';
      el.glow = (pe.effects && pe.effects.glow && typeof pe.effects.glow === 'object') ? pe.effects.glow : null;
      if (el.glow && el.glow.color) el.glow.color = colorFromPptx(el.glow.color) || '#FFFF00';
      // 标准 JSON 的行距是对象 { type:'percent', value }（也可能直接是数字）
      el.lineSpacing = typeof pe.lineSpacing === 'number'
        ? pe.lineSpacing
        : (pe.lineSpacing && typeof pe.lineSpacing === 'object' && typeof pe.lineSpacing.value === 'number')
          ? pe.lineSpacing.value
          : 1.15;
      // 「源是否显式指定」标记：属性面板必须显示非空的默认值，
      // 但导出时若把 18pt/微软雅黑/1.15 这些面板默认值当成显式值写回，
      // 会让原本靠母版/占位符继承字号与字体的文字被固化成默认值（整页排版随之改变）。
      el.propsSet = {
        fontSize: pe.fontSize != null,
        fontFace: !!pe.fontFace,
        align: !!pe.align,
        valign: !!pe.valign,
        lineSpacing: pe.lineSpacing != null && pe.lineSpacing !== ''
      };
      // 标准 JSON 的 bullet 是对象 {type:'number'|'bullet',...} 或 'number'/true
      const eb = pe.bullet;
      el.bullet = eb ? ((eb === 'number' || (eb.type === 'number')) ? 'number' : true) : false;
      // indentLeft 单位是 pt，编辑器内部用 px
      el.indent = pe.indentLeft ? Math.round(ptToPx(pe.indentLeft)) : 0;
      const fallback = { fontSize: el.fontSize, color: el.color, fontFace: el.fontFace };
      const lineSpacingNum = (ls) => {
        if (ls == null) return el.lineSpacing;
        if (typeof ls === 'number') return ls;
        if (ls.type === 'percent') return ls.value;
        if (ls.type === 'pt') return Math.max(0.8, Math.round((ls.value / (el.fontSize || 18)) * 100) / 100);
        return el.lineSpacing;
      };
      if (pe.paragraphs && pe.paragraphs.length) {
        el.paragraphs = pe.paragraphs.map((p) => {
          const pb = p.bullet;
          return {
            runs: p.runs ? pptxRunsToRuns(p.runs, fallback) : [{ text: String(p.text || '') }],
            align: p.align || el.align,
            // 保留编号格式（fmt：chineseCounting 等）与项目符号字符（char），渲染/导出都需要
            bullet: normalizeBulletIn(pb),
            lineSpacing: lineSpacingNum(p.lineSpacing),
            // 源段落是否显式写过：面板需要非空默认值，但导出时不能把继承来的值固化
            alignSet: !!p.align,
            lineSpacingSet: p.lineSpacing != null,
            // 从右到左段落（a:pPr@rtl）：RTL 段落缺省右对齐
            rtl: !!p.rtl,
            // 段落级缩进（pt→px）与段前/段后间距（pt→px）
            indent: p.indent != null ? Math.round(ptToPx(p.indent)) : undefined,
            // marL/marR：列表左/右边界（pt→px），缺失时导出端不写、由版式继承
            indentLeft: p.indentLeft != null ? Math.round(ptToPx(p.indentLeft)) : undefined,
            indentRight: p.indentRight != null ? Math.round(ptToPx(p.indentRight)) : undefined,
            spaceBefore: p.spaceBefore != null ? Math.round(ptToPx(p.spaceBefore)) : undefined,
            spaceAfter: p.spaceAfter != null ? Math.round(ptToPx(p.spaceAfter)) : undefined
          };
        });
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
        ? { color: colorFromPptx(pe.line.color) || '#000000', width: pe.line.width ?? 1, dashType: pe.line.dashType || 'solid' }
        : 'none';
      el.shadow = (pe.effects && pe.effects.shadow && typeof pe.effects.shadow === 'object') ? pe.effects.shadow : null;
      if (el.shadow && el.shadow.color) el.shadow.color = colorFromPptx(el.shadow.color) || '#000000';
      el.glow = (pe.effects && pe.effects.glow && typeof pe.effects.glow === 'object') ? pe.effects.glow : null;
      if (el.glow && el.glow.color) el.glow.color = colorFromPptx(el.glow.color) || '#FFFF00';
      if (pe.flipH) el.flipH = true;
      if (pe.flipV) el.flipV = true;
      if (pe.adjust && typeof pe.adjust === 'object') el.adjust = pe.adjust;
      if (pe.custGeom && Array.isArray(pe.custGeom.paths) && pe.custGeom.paths.length) el.custGeom = pe.custGeom;
      return el;
    }
    case 'image': {
      el.type = 'image';
      el.data = pe.data || '';
      el.src = pe.src || '';
      el.extension = pe.extension || '';
      el.imageAdjust = Object.assign({ brightness: 0, contrast: 0, transparency: 0 }, pe.imageAdjust || {});
      return el;
    }
    case 'table': {
      el.type = 'table';
      el.colWidths = (pe.colWidths && pe.colWidths.length) ? pe.colWidths.map(Number) : null;
      el.border = pe.border ? { color: normalizeColor(pe.border.color) || '#CBD5E1', width: pe.border.width ?? 1 } : { color: '#CBD5E1', width: 1 };
      if (pe.inset) el.inset = { l: pe.inset.l, r: pe.inset.r, t: pe.inset.t, b: pe.inset.b };
      if (pe.tableStyleId) el.tableStyleId = pe.tableStyleId;
      el.rows = (pe.rows || []).map((r) => ({
        height: Number(r.height) || 40,
        cells: (r.cells || []).map((c) => {
          const hasFill = c.fill && c.fill !== 'none';
          const hasColor = c.color && c.color !== 'none';
          const cellInset = c.inset ? { l: c.inset.l, r: c.inset.r, t: c.inset.t, b: c.inset.b } : undefined;
          // 段落形态（runs）拼为纯文本；带自动编号的段落前置编号文本（hebrew2Minus 等）
          const flatText = c.text || (Array.isArray(c.paragraphs)
            ? c.paragraphs.map((p, pi) => {
                const t = (p.runs || []).map((r) => (r.text == null ? '' : String(r.text))).join('');
                const b = p.bullet;
                if (b && typeof b === 'object' && b.type === 'number') {
                  return autoNumPrefix(b.fmt || 'arabicPeriod', (b.start || 1) + pi) + ' ' + t;
                }
                return t;
              }).join('\n')
            : '');
          const cell = {
            text: flatText,
            // 原始段落结构（含 RTL 对齐、自动编号格式）：文本未被改动时原样回写
            paragraphs: Array.isArray(c.paragraphs) && c.paragraphs.length ? c.paragraphs : undefined,
            paragraphsText: flatText,
            fill: hasFill ? colorFromPptx(c.fill) : null,
            fillCustom: hasFill || undefined,
            align: c.align || 'left',
            valign: c.valign || 'middle',
            fontSize: c.fontSize || 14,
            color: colorFromPptx(c.color) || '#202124',
            colorCustom: hasColor || undefined,
            bold: !!c.bold,
            // 「源是否显式写过」标记：属性面板需要非空的默认值才能显示，
            // 但导出时必须能区分「源文件本来就有」与「面板补的默认值」，
            // 否则会把 14pt/middle/left 强加到原本靠主题继承的单元格上，字号与位置随之改变。
            rtl: !!c.rtl,
            alignSet: c.align != null,
            valignSet: c.valign != null,
            fontSizeSet: c.fontSize != null,
            colSpan: c.colSpan, rowSpan: c.rowSpan,
            inset: cellInset,
            // 同上：保留全部边（含 insideH/insideV 与 diagonal 方向串）
            borders: c.borders
              ? Object.fromEntries(Object.entries(c.borders).map(([k, b]) => [k,
                  b === 'none' ? 'none'
                    : typeof b === 'string' ? b
                      : (b && typeof b === 'object') ? { color: colorFromPptx(b.color) || '#000000', width: Number(b.width) || 1 } : undefined])
                .filter(([, v]) => v !== undefined))
              : undefined
          };
          return cell;
        })
      }));
      // 根据已有单元格反推表头/单元格默认色，让属性面板与渲染一致
      const firstRow = el.rows[0];
      const secondRow = el.rows[1];
      if (firstRow && firstRow.cells[0] && firstRow.cells[0].fillCustom) {
        el.headerFill = firstRow.cells[0].fill;
      }
      if (secondRow && secondRow.cells[0] && secondRow.cells[0].fillCustom) {
        el.cellFill = secondRow.cells[0].fill;
      }
      // 对角线：目前只保留方向；后续可扩展颜色/线宽
      const anyDiag = el.rows.find((r) => r.cells.find((c) => c.borders && c.borders.diagonal));
      if (anyDiag) el.hasDiagonal = true;
      if (!el.colWidths) {
        const cols = el.rows[0] ? el.rows[0].cells.length : 1;
        el.colWidths = new Array(cols).fill(Math.round(el.width / cols));
      }
      el.headerRow = false;
      el.headerFill = '#1A73E8';
      el.cellFill = '#FFFFFF';
      el.fontSize = el.rows[0]?.cells[0]?.fontSize || 14;
      // 源表格是否显式指定过字号：决定导出时是否回写（未指定则保留继承）
      el.fontSizeSet = (pe.rows || []).some((r) => (r.cells || []).some((c) => c.fontSize != null));
      el.color = '#202124';
      return el;
    }
    case 'chart': {
      el.type = 'chart';
      // 3D/衍生图表类型映射为编辑器支持的 2D 等价类型（保留原始类型供导出还原）
      const CHART_3D_MAP = {
        bar3DChart: 'barChart', line3DChart: 'lineChart', area3DChart: 'areaChart',
        pie3DChart: 'pieChart', surface3DChart: 'areaChart', ofPieChart: 'pieChart',
        bubbleChart: 'scatterChart', stockChart: 'lineChart'
      };
      el.chartType = CHART_3D_MAP[pe.chartType] || pe.chartType || 'barChart';
      if (CHART_3D_MAP[pe.chartType]) el.chartType3D = pe.chartType;
      el.title = pe.title || '';
      el.legend = pe.legend !== false;
      el.legendPosition = pe.legendPosition || '';
      el.dataLabels = !!pe.dataLabels;
      el.grouping = pe.grouping || 'clustered';
      el.holeSize = pe.holeSize ?? 50;
      el.smooth = !!pe.smooth;
      el.marker = !!pe.marker;
      el.view3D = pe.view3D || null;
      // 图表区填充（c:chartSpace/c:spPr）：'none' 或 #RRGGBB
      el.spaceFill = pe.spaceFill === 'none' ? 'none' : (normalizeColor(pe.spaceFill) || '');
      el.categories = pe.categories || [];
      const isBubble = el.chartType3D === 'bubbleChart';
      const isScatter = !isBubble && el.chartType === 'scatterChart';
      const isStock = el.chartType3D === 'stockChart';
      el.series = (pe.series || []).map((s) => {
        const ser = { name: s.name || '系列', color: normalizeColor(s.color) || '' };
        if (s.pointColors && s.pointColors.some((c) => c)) {
          ser.pointColors = (s.pointColors || []).map(normPointColor);
        }
        if (isBubble) {
          ser.values = (s.x || []).map((x, i) => ({
            x: Number(x) || 0,
            y: Number((s.y || [])[i]) || 0,
            size: Number((s.values || [])[i]) || 0
          }));
        } else if (isScatter) {
          ser.values = (s.x || []).map((x, i) => ({
            x: Number(x) || 0,
            y: Number((s.y || [])[i]) || 0
          }));
        } else if (isStock) {
          const len = Math.max(
            (s.open || []).length, (s.close || []).length,
            (s.low || []).length, (s.high || []).length
          );
          ser.values = Array.from({ length: len }, (_, i) => [
            Number((s.open || [])[i]) || 0,
            Number((s.close || [])[i]) || 0,
            Number((s.low || [])[i]) || 0,
            Number((s.high || [])[i]) || 0
          ]);
        } else {
          ser.values = (s.values || s.y || []).map((v) => Number(v) || 0);
        }
        return ser;
      });
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
      // 记录子元素坐标约定：relative=相对组合原点；local 已折算为页面绝对坐标
      el.childrenCoordinates = pe.childrenCoordinates === 'relative' ? 'relative' : 'page';
      return el;
    }
    case 'video':
    case 'audio': {
      el.type = pe.type;
      el.data = pe.data || '';
      el.src = pe.src || '';
      el.extension = pe.extension || '';
      el.poster = pe.poster || null;
      return el;
    }
    case 'diagram': {
      el.type = 'diagram';
      el.texts = (pe.texts || []).map(String);
      el.shapes = Array.isArray(pe.shapes) ? pe.shapes : [];
      el.dataPath = pe.dataPath || '';
      // 保留原始 SmartArt 部件（data1/colors1/quickStyle1/layout1/drawing1 的整串 XML），
      // 否则序列化器只能从 shapes/texts 生成最小化 diagram 文件，丢失原始配色与布局
      if (pe.__raw) el.__raw = pe.__raw;
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
  // 保留原始元数据与自定义属性，导出时原样写回
  if (pptxDoc.metadata) doc.metadata = pptxDoc.metadata;
  if (pptxDoc.customProps) doc.customProps = pptxDoc.customProps;
  if (pptxDoc.thumbnail) doc.thumbnail = pptxDoc.thumbnail;
  // 导入的 PPTX 若携带主题配色（a:theme/a:clrScheme），提前设置为当前主题，确保下方 elementFromPptx
  // 解析 'scheme:<name>' 引用时使用该主题，而非默认蓝主题
  const th = pptxDoc.theme;
  if (th && th.colors) {
    const c = th.colors;
    doc.theme = {
      id: 'imported', name: th.name || '导入主题',
      bg: c.lt1 || '#FFFFFF', panel: c.lt2 || c.lt1 || '#FFFFFF',
      text: c.dk1 || '#202124', title: c.dk2 || c.dk1 || '#202124',
      accent: c.accent1 || '#1A73E8',
      accents: [c.accent1, c.accent2, c.accent3, c.accent4, c.accent5, c.accent6]
        .map((x) => x || '#1A73E8')
    };
    setActiveTheme(doc.theme);
  }
  // 保留解析端回读的原始 theme1.xml（无损回退，确保导出时 fmtScheme/fontScheme 等细节与源文件一致）
  if (pptxDoc.themeXml) doc.themeXml = pptxDoc.themeXml;
  // 保留原始 tableStyles.xml（表格网格线颜色/底纹/条带的样式定义，重新生成会退化为黑网格）
  if (pptxDoc.tableStylesXml) doc.tableStylesXml = pptxDoc.tableStylesXml;
  // 保留原始主题部件列表与母版/版式（多主题文件：不同页绑定不同母版→不同主题，
  // 丢弃后所有页会退化为单一 theme1.xml，SmartArt/图表的 schemeClr 配色随之错位）
  if (pptxDoc.themeXmls && pptxDoc.themeXmls.length) doc.themeXmls = pptxDoc.themeXmls;
  if (pptxDoc.masters && pptxDoc.masters.length) doc.masters = pptxDoc.masters;
  doc.slides = (pptxDoc.slides || []).map((s) => {
    const slide = createSlide((s.elements || []).map(elementFromPptx), {
      background: s.background || null,
      notes: s.notes || '',
      hidden: !!s.hidden
    });
    if (s.transition && s.transition.type) {
      slide.transition = { type: s.transition.type, duration: s.transition.duration || 800 };
    }
    // 版式下标（展平版式序列）：配合 doc.masters 写回原母版/版式/主题绑定
    if (typeof s.layout === 'number') slide.layout = s.layout;
    // 批注：pos 由 EMU 转 px（与元素坐标一致），导出时再转回
    if (Array.isArray(s.comments) && s.comments.length) {
      slide.comments = s.comments.map((c) => ({
        author: c.author || 'Author',
        text: c.text || '',
        dt: c.dt || '',
        pos: c.pos ? { x: (Number(c.pos.x) || 0) / 9525, y: (Number(c.pos.y) || 0) / 9525 } : null
      }));
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
    // 用户显式改过 → 标记为该属性已指定，导出时才写回（否则会被当成面板默认值丢弃）
    if (el.propsSet && k in el.propsSet) el.propsSet[k] = true;
  }
  const keys = ['fontSize', 'color', 'bold', 'italic', 'underline', 'fontFace'];
  for (const p of el.paragraphs || []) {
    for (const r of p.runs || []) {
      for (const k of keys) if (patch[k] !== undefined) r[k] = patch[k];
    }
    // 用户改过段落级样式 → 标记为已指定，导出时才写回（否则视为继承值丢弃）
    if (patch.align !== undefined) { p.align = patch.align; p.alignSet = true; }
    if (patch.bullet !== undefined) p.bullet = patch.bullet === false ? undefined : patch.bullet;
    if (patch.lineSpacing !== undefined) { p.lineSpacing = patch.lineSpacing; p.lineSpacingSet = true; }
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
