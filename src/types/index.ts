/** 核心配置项 - 对应原库$("#id").pptxToHtml(options)的入参 */
export interface PptxToHtmlOptions {
  /** PPTX文件路径（浏览器：URL；Node.js：本地路径/Buffer） */
  pptxFileUrl: string | Buffer;
  /** 上传文件input的ID（仅浏览器） */
  fileInputId?: string;
  /** 幻灯片缩放比例（百分比） */
  slidesScale?: string;
  /** 是否开启幻灯片模式 */
  slideMode: boolean;
  /** 是否开启键盘快捷键 */
  keyBoardShortCut: boolean;
  /** 是否处理音视频媒体 */
  mediaProcess: boolean;
  /** JSZip v2路径（false则使用内置） */
  jsZipV2: string | false;
  /** 主题处理规则 */
  themeProcess: boolean | "colorsAndImageOnly";
  /** 幻灯片尺寸增量（px） */
  incSlide?: { height: number; width: number };
  /** 幻灯片渲染模式 */
  slideType: "divs2slidesjs" | "revealjs";
  /** divs2slidesjs模式配置 */
  slideModeConfig: SlideModeConfig;
  /** revealjs模式配置 */
  revealjsConfig?: RevealJsConfig;
}

/** 幻灯片模式配置（divs2slidesjs） */
export interface SlideModeConfig {
  /** 起始幻灯片 */
  first: number;
  /** 是否显示导航按钮 */
  nav: boolean;
  /** 导航文本颜色 */
  navTxtColor: string;
  /** 下一页导航文本（HTML实体） */
  navNextTxt?: string;
  /** 上一页导航文本（HTML实体） */
  navPrevTxt?: string;
  /** 是否显示播放/暂停按钮 */
  showPlayPauseBtn: boolean;
  /** 是否显示当前页码 */
  showSlideNum: boolean;
  /** 是否显示总页数 */
  showTotalSlideNum: boolean;
  /** 自动轮播（false/秒数） */
  autoSlide: false | number;
  /** 随机自动轮播（需autoSlide=true） */
  randomAutoSlide: boolean;
  /** 是否循环播放 */
  loop: boolean;
  /** 背景色（false/颜色值） */
  background: false | string;
  /** 切换动画类型 */
  transition: "slid" | "fade" | "default" | "random";
  /** 切换动画时长（秒） */
  transitionTime: number;
}

/** Reveal.js配置 */
export interface RevealJsConfig {
  transition?: string;
  backgroundTransition?: string;
  autoSlide?: number;
  loop?: boolean;
  slideNumber?: boolean;
  // 可扩展Reveal.js原生配置
}

/** PPTX元素类型枚举 */
export enum PptxElementType {
  TEXT = "text",
  TEXT_BLOCK = "text_block",
  SHAPE = "shape",
  MEDIA = "media",
  GRAPH = "graph",
  TABLE = "table",
  SMART_ART = "smart_art",
  EQUATION = "equation",
}

/** 媒体类型（图片/音频/视频） */
export interface MediaInfo {
  type: "image" | "audio" | "video";
  src: string;
  format: string; // jpg/mp4/mp3等
  /** 各浏览器兼容配置（原库的浏览器适配规则） */
  browserSupport: Record<string, string[]>;
}

/** 环境类型 */
export type EnvType = "browser" | "node";