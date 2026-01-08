import { PptxToHtmlOptions } from "./types/index";
import { readPptxFile, detectEnv } from "./adapter/fileReader";
import { PptxRenderer } from "./render/pptxRenderer";

/**
 * PPTX转HTML核心方法
 * @param options 配置项（必须包含 pptxFileUrl）
 * @returns 返回完整的HTML字符串
 */
export async function pptxToHtml(
  options: PptxToHtmlOptions
): Promise<string> {
  // 1. 读取PPTX文件
  const buffer = await readPptxFile(options.pptxFileUrl);

  // 2. 初始化渲染器
  const renderer = new PptxRenderer(options, buffer);

  // 3. 渲染并返回结果
  return renderer.render();
}

// 浏览器环境：挂载到jQuery/全局
if (detectEnv() === "browser") {
  // 挂载到全局
  (window as any).pptxToHtml = pptxToHtml;
}

export { 
  detectEnv,
  readPptxFile
 };
// 导出所有类型
export * from "./types";