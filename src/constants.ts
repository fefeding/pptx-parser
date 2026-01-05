/**
 * PPTX解析核心常量定义
 * 遵循 ECMA-376 OpenXML 标准
 */

/** PPTX 命名空间定义（官方标准） */
export const NS = {
  /** PresentationML 命名空间 - 幻灯片核心标签 */
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  /** DrawingML 命名空间 - 绘图、样式、效果标签 */
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  /** Relationships 命名空间 - 关联关系标签 */
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  /** Markup Compatibility 命名空间 - 兼容性扩展标签 */
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  /** Chart 命名空间 - 图表标签 */
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart'
} as const;

/** 单位转换常量 */
export const EMU_PER_INCH = 914400; // 1 英寸 = 914400 EMU
export const PIXELS_PER_INCH = 96; // 1 英寸 = 96 像素 (默认DPI)
export const EMU_TO_PIXEL_RATIO = PIXELS_PER_INCH / EMU_PER_INCH; // EMU转像素比例
export const PIXEL_TO_EMU_RATIO = EMU_PER_INCH / PIXELS_PER_INCH; // 像素转EMU比例

/** 默认值常量 */
export const DEFAULTS = {
  SLIDE_WIDTH: 1280,
  SLIDE_HEIGHT: 720,
  FONT_SIZE: 14,
  FONT_COLOR: '#333333',
  BACKGROUND_COLOR: '#ffffff',
  BORDER_COLOR: '#000000',
  BORDER_WIDTH: 1
} as const;

/** PPTX 目录结构常量 */
export const PATHS = {
  SLIDES: 'ppt/slides/',
  SLIDE_RELS: 'ppt/slides/_rels/',
  MEDIA: 'ppt/media/',
  DOCPROPS: 'docProps/',
  THEMES: 'ppt/theme/',
  LAYOUTS: 'ppt/slideLayouts/'
} as const;

/** 元素类型常量（对应XML标签） */
export const ELEMENT_TYPES = {
  SHAPE: 'shape',
  IMAGE: 'image',
  OLE: 'ole',
  GROUP: 'group',
  CHART: 'chart',
  TABLE: 'table',
  VIDEO: 'video',
  AUDIO: 'audio'
} as const;

/** 文本对齐方式常量 */
export const TEXT_ALIGN = {
  LEFT: 'left',
  CENTER: 'center',
  RIGHT: 'right',
  JUSTIFY: 'justify'
} as const;

/** 文本垂直对齐方式常量 */
export const VERTICAL_ALIGN = {
  TOP: 'top',
  MIDDLE: 'middle',
  BOTTOM: 'bottom'
} as const;
