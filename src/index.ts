import { PptxToHtmlOptions } from "./types/index";
import { readPptxFile, detectEnv } from "./adapter/fileReader";
import { PptxRenderer } from "./render/pptxRenderer";

/**
 * PPTX转HTML核心方法
 * @param targetId 目标容器ID（浏览器）/ 仅用于Node.js生成HTML结构
 * @param options 配置项
 * @returns Node.js返回HTML字符串；浏览器无返回（直接挂载DOM）
 */
export async function pptxToHtml(
  targetId: string,
  options: PptxToHtmlOptions
): Promise<void | string> {
  // 1. 读取PPTX文件
  const buffer = await readPptxFile(options.pptxFileUrl);

  // 2. 初始化渲染器
  const renderer = new PptxRenderer(options, buffer);

  // 3. 渲染并返回结果
  return renderer.render(targetId);
}

// 浏览器环境：挂载到jQuery/全局
if (detectEnv() === "browser") {
  // 适配jQuery（复刻原库的$("#id").pptxToHtml()）
  if (typeof window !== "undefined" && (window as any).jQuery) {
    (window as any).jQuery.fn.pptxToHtml = function (options: PptxToHtmlOptions) {
      const targetId = this.attr("id");
      if (!targetId) throw new Error("元素必须有ID");
      pptxToHtml(targetId, options);
      return this;
    };
  }

  // 挂载到全局
  (window as any).pptxToHtml = pptxToHtml;
}

// 导出所有类型
export * from "./types";