import { BaseElement } from './BaseElement';
import type { ThemeResult, ThemeColors } from '../core/types';
export interface FontScheme {
    majorFont?: {
        latin?: string;
        ea?: string;
        cs?: string;
        [script: string]: string | undefined;
    };
    minorFont?: {
        latin?: string;
        ea?: string;
        cs?: string;
        [script: string]: string | undefined;
    };
}
export interface EffectScheme {
    fillStyles?: any[];
    lineStyles?: any[];
    effectStyles?: any[];
    bgFillStyles?: any[];
}
export declare class ThemeElement extends BaseElement {
    type: 'theme';
    name: string;
    colors: ThemeColors;
    fonts?: FontScheme;
    effects?: EffectScheme;
    themeId: string;
    constructor(id: string, name: string, colors: ThemeColors, themeId?: string, fonts?: FontScheme, effects?: EffectScheme);
    static fromResult(result: ThemeResult, themeName?: string): ThemeElement;
    getThemeClassPrefix(): string;
    getThemeClass(suffix: string): string;
    private sanitizeThemeName;
    generateThemeCSS(): string;
    getColor(colorKey: keyof ThemeColors): string | undefined;
    getMajorFont(): string | undefined;
    getMinorFont(): string | undefined;
    toHTML(): string;
}
