import { EnvType } from "../types/index";

// 浏览器兼容的Buffer polyfill
class BrowserBuffer {
  private data: Uint8Array;

  constructor(input: ArrayBuffer | Uint8Array | number) {
    if (typeof input === 'number') {
      this.data = new Uint8Array(input);
    } else if (input instanceof ArrayBuffer) {
      this.data = new Uint8Array(input);
    } else {
      this.data = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    }
  }

  static isBuffer(obj: any): boolean {
    return obj instanceof BrowserBuffer;
  }

  static from(arrayBuffer: ArrayBuffer): BrowserBuffer {
    return new BrowserBuffer(arrayBuffer);
  }

  toString(): string {
    return String.fromCharCode.apply(null, Array.from(this.data));
  }
}

// 全局Buffer兼容处理
if (typeof window !== 'undefined' && typeof Buffer === 'undefined') {
  (window as any).Buffer = BrowserBuffer;
}

// 使用全局Buffer或polyfill
const BufferCompat = typeof Buffer !== 'undefined' ? Buffer : (window as any).Buffer;

/** 检测运行环境 */
export const detectEnv = (): EnvType => {
  if (typeof window !== "undefined" && typeof document !== "undefined") {
    return "browser";
  }
  return "node";
};

/** 读取PPTX文件（适配双环境） */
export const readPptxFile = async (
  source: string | ArrayBuffer | Uint8Array,
  env: EnvType = detectEnv()
): Promise<Uint8Array> => {
  if (env === "node") {
    // Node.js环境：需要动态导入fs
    const fs = await import("fs");
    // Node.js环境：读取本地文件/直接使用Buffer
    if (BufferCompat.isBuffer && BufferCompat.isBuffer(source)) {
      return new Uint8Array((source as any).buffer || source);
    }
    if (typeof source === 'string') {
      const buffer = await fs.promises.readFile(source);
      return new Uint8Array(buffer.buffer || buffer);
    }
    return new Uint8Array(source);
  }

  // 浏览器环境：FileReader读取（支持URL/File对象）
  return new Promise((resolve, reject) => {
    if (typeof source === "string") {
      // 远程URL：先fetch再转Uint8Array
      fetch(source)
        .then((res) => res.arrayBuffer())
        .then((ab) => resolve(new Uint8Array(ab)))
        .catch(reject);
    } else if (source instanceof File) {
      // 本地文件（input上传）
      const reader = new FileReader();
      reader.onload = (e) => resolve(new Uint8Array(e.target!.result as ArrayBuffer));
      reader.onerror = reject;
      reader.readAsArrayBuffer(source);
    } else if (source instanceof ArrayBuffer || source instanceof Uint8Array) {
      // 直接使用ArrayBuffer或Uint8Array
      resolve(new Uint8Array(source));
    } else {
      reject(new Error("浏览器环境仅支持URL/File对象/ArrayBuffer作为PPTX源"));
    }
  });
};