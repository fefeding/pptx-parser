/// <reference types="vite/client" />

declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}

// 根目录 dist/convert.browser.js 为构建产物，类型定义见 dist/convert.d.ts
declare module '*convert.browser.js' {
  export interface ConvertOptions {
    title?: string
    fonts?: { regular?: Uint8Array | ArrayBuffer | string; bold?: Uint8Array | ArrayBuffer | string }
    autoFont?: boolean
  }
  export function convert(
    input: ArrayBuffer | Uint8Array | Blob | string,
    target: string,
    options?: ConvertOptions
  ): Promise<Uint8Array | string>
}

declare module '*cjk-font.js' {
  export function collectCjkChars(text: string): string
  export function loadCjkFont(chars: string): Promise<Uint8Array | null>
}
