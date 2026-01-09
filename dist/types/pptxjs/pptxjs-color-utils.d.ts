export declare enum ColorType {
    SOLID = "solid",
    GRADIENT = "gradient",
    PATTERN = "pattern",
    NONE = "none"
}
export interface ColorValue {
    type: ColorType;
    color?: string;
    alpha?: number;
    stops?: Array<{
        position: number;
        color: string;
        alpha?: number;
    }>;
}
export interface ColorMap {
    bg1?: string;
    tx1?: string;
    bg2?: string;
    tx2?: string;
    accent1?: string;
    accent2?: string;
    accent3?: string;
    accent4?: string;
    accent5?: string;
    accent6?: string;
    hlink?: string;
    folHlink?: string;
}
export declare const THEME_COLORS: Record<string, string>;
export declare function getColorValue(colorNode: any): string | null;
export declare function getThemeColor(schemeColor: string): string;
export declare function getPresetColor(presetColor: string): string;
export declare function getAlphaValue(colorNode: any): number;
export declare function applyColorMap(color: string, colorMapOvr?: ColorMap): string;
export declare function parseColorFill(fillNode: any): ColorValue | null;
export declare function generateCssColor(colorValue: ColorValue): string;
export declare function hexToRgba(hex: string, alpha: number): string;
export declare function parseColorMapOverride(slideContent: any, slideLayoutContent: any, slideMasterContent: any): {
    slide?: ColorMap;
    layout?: ColorMap;
    master?: ColorMap;
};
export declare function getTextByPathList(obj: any, pathList: string[]): any;
