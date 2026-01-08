import { detectEnv } from "./fileReader";
import { EnvType } from "../types/index";

/** DOM操作适配器 */
export class DomAdapter {
  private env: EnvType;

  constructor(env?: EnvType) {
    this.env = env || detectEnv();
  }

  /** 创建元素（浏览器：真实DOM；Node：HTML字符串） */
  createElement(tag: string, attrs?: Record<string, string>, content?: string): HTMLElement | string {
    if (this.env === "browser") {
      const el = document.createElement(tag);
      if (attrs) {
        Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
      }
      if (content) el.innerHTML = content;
      return el;
    }

    // Node.js：拼接HTML字符串
    const attrStr = attrs ? Object.entries(attrs).map(([k, v]) => `${k}="${v}"`).join(" ") : "";
    return `<${tag} ${attrStr}>${content || ""}</${tag}>`;
  }

  /** 挂载元素（浏览器：挂载到DOM；Node：返回根HTML） */
  mount(el: HTMLElement | string, targetId: string): void | string {
    if (this.env === "browser") {
      const target = document.getElementById(targetId);
      if (!target) throw new Error(`目标元素#${targetId}不存在`);
      target.appendChild(el as HTMLElement);
      return;
    }

    // Node.js：返回完整HTML结构
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/pptxjs-ts/dist/css/pptxjs.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/nvd3@1.8.6/build/nv.d3.min.css">
</head>
<body>
  <div id="${targetId}">${el}</div>
</body>
</html>
    `;
  }
}