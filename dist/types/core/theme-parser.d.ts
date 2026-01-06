import JSZip from 'jszip';
export interface ThemeColors {
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
export interface ThemeResult {
    colors: ThemeColors;
}
export declare function parseTheme(zip: JSZip, themePath?: string): Promise<ThemeResult | null>;
export declare function parseColorMap(root: Element): Record<string, string>;
export declare function resolveSchemeColor(schemeColor: string, themeColors: ThemeColors, colorMap?: Record<string, string>): string;
