import { PptxToHtmlOptions } from "./types/index";
import { readPptxFile, detectEnv } from "./adapter/fileReader";
export declare function pptxToHtml(options: PptxToHtmlOptions): Promise<string>;
export { detectEnv, readPptxFile };
export * from "./types";
