/**
 * PPTX Parser TypeScript Entry Point
 * @module @fefeding/ppt-parser
 * @description PPTX文件解析与序列化核心库
 */

// Import the JavaScript modules (they expose globals)
import './js/pptxjs.js';
import './js/utils.js';
import './js/pptx-parser.js';

// The JavaScript modules expose their APIs via global window object
// For TypeScript users, these are accessible via the global window object

// Global type declarations for PPTX global objects
declare global {
  interface Window {
    PPTXUtils: {
      getTextByPathList: (node: any, path: string[]) => string | undefined;
      getTextByPathListStr: (node: any, path: string[], defaultVal?: string) => string;
      getVal: (node: any, path: string[], defaultVal?: string) => string;
      getBool: (node: any, path: string[], defaultVal?: boolean) => boolean;
      getInt: (node: any, path: string[], defaultVal?: number) => number;
      getUnit: (node: any, path: string[], defaultVal?: string) => string;
      getColor: (node: any, path: string[]) => string;
      spPr2ShapeStr: (spNode: any, slideLayoutSpNode: any, isSlideModeBg: boolean) => string;
      archaicNumbers: (num: number) => string;
      resolveRelationshipTarget: (relId: string, relationships: any) => string;
      [key: string]: any;
    };
    PPTXParser: {
      configure: (options: any) => void;
      parse: (fileData: any) => any;
      indexNodes: (node: any) => { [key: string]: any };
      [key: string]: any;
    };
    PPTXHtml: {
      [key: string]: any;
    };
  }
}

export {};
