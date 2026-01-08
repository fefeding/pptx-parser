import fs from "fs";
import { EnvType } from "../types/index";

/** 检测运行环境 */
export const detectEnv = (): EnvType => {
  if (typeof window !== "undefined" && typeof document !== "undefined") {
    return "browser";
  }
  return "node";
};

/** 读取PPTX文件（适配双环境） */
export const readPptxFile = async (
  source: string | Buffer,
  env: EnvType = detectEnv()
): Promise<Buffer> => {
  if (env === "node") {
    // Node.js环境：读取本地文件/直接使用Buffer
    if (Buffer.isBuffer(source)) {
      return source;
    }
    return fs.promises.readFile(source);
  }

  // 浏览器环境：FileReader读取（支持URL/File对象）
  return new Promise((resolve, reject) => {
    if (typeof source === "string") {
      // 远程URL：先fetch再转Buffer
      fetch(source)
        .then((res) => res.arrayBuffer())
        .then((ab) => resolve(Buffer.from(ab)))
        .catch(reject);
    } else if (source instanceof File) {
      // 本地文件（input上传）
      const reader = new FileReader();
      reader.onload = (e) => resolve(Buffer.from(e.target!.result as ArrayBuffer));
      reader.onerror = reject;
      reader.readAsArrayBuffer(source);
    } else {
      reject(new Error("浏览器环境仅支持URL/File对象作为PPTX源"));
    }
  });
};