import type { SlideLayoutResult, MasterSlideResult } from './types';
export interface StyleContext {
    slide?: any;
    layout?: SlideLayoutResult;
    master?: MasterSlideResult;
    theme?: any;
}
export declare function getPlaceholderStyle(element: any, context: StyleContext): any;
export declare function mergeStyles(baseStyle: any, newStyle: any): any;
export declare function resolveColor(colorNode: any, themeColors?: Record<string, string>): string | undefined;
export declare function createStyleContext(slide?: any, layout?: SlideLayoutResult, master?: MasterSlideResult, theme?: any): StyleContext;
export declare function applyStyleInheritance(slide: any, layout?: SlideLayoutResult, master?: MasterSlideResult, theme?: any): void;
