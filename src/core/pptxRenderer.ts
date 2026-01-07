import { PptxToHtmlOptions, EnvType } from "../types";
import { PptxParser } from "../core/pptxParser";
import { DomAdapter } from "../adapter/domAdapter";
import { detectEnv } from "../adapter/fileReader";

/** PPTX转HTML渲染器 */
export class PptxRenderer {
  private env: EnvType;
  private domAdapter: DomAdapter;
  private parser: PptxParser;

  constructor(private options: PptxToHtmlOptions, private buffer: Buffer) {
    this.env = detectEnv();
    this.domAdapter = new DomAdapter(this.env);
    this.parser = new PptxParser(buffer);
  }

  /** 核心渲染方法 */
  async render(targetId: string): Promise<void | string> {
    // 1. 初始化解析器
    await this.parser.init();
    const slides = this.parser.getSlides();
    const totalSlides = this.parser.getTotalSlides();

    // 2. 渲染幻灯片容器
    const container = this.domAdapter.createElement(
      "div",
      { class: "pptxjs-container" },
      ""
    );

    // 3. 逐页渲染幻灯片
    for (let i = 1; i <= totalSlides; i++) {
      const slideData = slides[i];
      const slideEl = this.renderSlide(slideData, i);
      if (this.env === "browser") {
        (container as HTMLElement).appendChild(slideEl as HTMLElement);
      } else {
        (container as string) += slideEl as string;
      }
    }

    // 4. 初始化交互（全屏、快捷键、轮播等）
    if (this.env === "browser") {
      this.initInteractions(container as HTMLElement);
    }

    // 5. 挂载容器
    return this.domAdapter.mount(container, targetId);
  }

  /** 渲染单张幻灯片 */
  private renderSlide(slideData: any, slideId: number): HTMLElement | string {
    // 根据slideType选择渲染模式（divs2slidesjs/revealjs）
    if (this.options.slideType === "revealjs") {
      return this.renderRevealJsSlide(slideData, slideId);
    }
    return this.renderDivs2SlidesJsSlide(slideData, slideId);
  }

  /** 渲染divs2slidesjs模式幻灯片 */
  private renderDivs2SlidesJsSlide(slideData: any, slideId: number): HTMLElement | string {
    // 复刻原库的div布局逻辑：
    // 1. 渲染文本块（对齐、背景、边框）
    // 2. 渲染形状（背景色/渐变、旋转、边框）
    // 3. 渲染媒体（图片/音视频播放器）
    // 4. 渲染图表（基于d3/nv.d3）
    // 5. 渲染表格/公式/SmartArt
    const slideAttrs = {
      class: "slide",
      "data-slide-id": slideId.toString(),
      style: `width: ${100 - (this.options.incSlide?.width || 0)}%; height: ${100 - (this.options.incSlide?.height || 0)}%`,
    };
    return this.domAdapter.createElement("div", slideAttrs, "");
  }

  /** 渲染revealjs模式幻灯片 */
  private renderRevealJsSlide(slideData: any, slideId: number): HTMLElement | string {
    // 适配Reveal.js的HTML结构
    const slideAttrs = {
      class: "reveal-slide",
      "data-slide-id": slideId.toString(),
    };
    return this.domAdapter.createElement("section", slideAttrs, "");
  }

  /** 初始化交互（全屏、快捷键、轮播等，仅浏览器） */
  private initInteractions(container: HTMLElement): void {
    // 复刻原库的交互逻辑：
    // 1. 全屏功能（基于jquery.fullscreen）
    // 2. 幻灯片轮播（autoSlide/loop）
    // 3. 键盘快捷键（上一页/下一页/播放/暂停）
    // 4. 导航按钮（上一页/下一页）
    // 5. 切换动画（slid/fade等）
  }
}