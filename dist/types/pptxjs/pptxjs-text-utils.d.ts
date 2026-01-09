export declare enum TextAlign {
    LEFT = "left",
    CENTER = "center",
    RIGHT = "right",
    JUSTIFY = "justify",
    DISTRIBUTED = "distributed"
}
export declare enum VerticalAlign {
    TOP = "top",
    MIDDLE = "middle",
    BOTTOM = "bottom",
    JUSTIFY = "justify",
    DISTRIBUTED = "distributed"
}
export interface TextStyle {
    fontFace?: string;
    fontSize?: number;
    color?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    baseline?: number;
    textAlign?: TextAlign;
    textVerticalAlign?: VerticalAlign;
    lineSpacing?: number;
    spacingBefore?: number;
    spacingAfter?: number;
    indent?: number;
    marginLeft?: number;
    marginRight?: number;
    textHighlight?: string;
    textShadow?: boolean;
}
export interface TextParagraph {
    text: string;
    styles?: TextStyle[];
    textAlign?: TextAlign;
    textVerticalAlign?: VerticalAlign;
    lineSpacing?: number;
    spacingBefore?: number;
    spacingAfter?: number;
    indent?: number;
    marginLeft?: number;
    marginRight?: number;
}
export interface TextRun {
    text: string;
    style?: TextStyle;
}
export declare function parseTextProps(textPropsNode: any): TextStyle;
export declare function parseParagraphProps(paraPropsNode: any): Partial<TextStyle>;
export declare function parseTextBoxContent(txBodyNode: any): TextParagraph[];
export declare function mergeTextStyles(baseStyle: TextStyle, ...additionalStyles: (TextStyle | undefined)[]): TextStyle;
export declare function generateTextStyleCss(style: TextStyle): string;
export declare function generateTextParagraphHtml(paragraph: TextParagraph): string;
export declare function generateTextBoxHtml(paragraphs: TextParagraph[]): string;
export declare function processTextLineBreaks(text: string): string;
export declare function getDefaultTextStyle(): TextStyle;
export declare function getTextByPathList(obj: any, pathList: string[]): any;
