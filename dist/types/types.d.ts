export type PptNodeType = 'text' | 'image' | 'shape' | 'table' | 'chart' | 'container' | 'media' | 'video' | 'audio' | 'line' | 'connector' | 'group' | 'smartart' | 'equation';
export interface PptRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
export type Position = PptRect;
export interface PptTransform {
    rotate?: number;
    flipH?: boolean;
    flipV?: boolean;
}
export interface PptTextStyle {
    fontSize?: number;
    fontFamily?: string;
    fontStyle?: 'normal' | 'italic';
    fontWeight?: 'normal' | 'bold';
    textDecoration?: 'none' | 'underline' | 'line-through';
    color?: string;
    textAlign?: 'left' | 'center' | 'right' | 'justify';
    textVerticalAlign?: 'top' | 'middle' | 'bottom';
    lineHeight?: number;
    letterSpacing?: number;
    textShadow?: string;
}
export interface PptFill {
    type?: 'solid' | 'gradient' | 'pattern' | 'picture' | 'none';
    color?: string;
    gradientStops?: Array<{
        position: number;
        color: string;
    }>;
    gradientDirection?: number;
    image?: string;
    opacity?: number;
}
export interface PptBorder {
    color?: string;
    width?: number;
    style?: 'solid' | 'dashed' | 'dotted' | 'double';
    dashStyle?: string;
}
export interface PptShadow {
    color?: string;
    blur?: number;
    offsetX?: number;
    offsetY?: number;
    opacity?: number;
}
export interface PptReflection {
    blur?: number;
    opacity?: number;
    offset?: number;
}
export interface PptGlow {
    color?: string;
    radius?: number;
    opacity?: number;
}
export interface PptEffect3D {
    material?: 'matte' | 'plastic' | 'metal' | 'wireframe';
    lightRig?: 'harsh' | 'flat' | 'normal' | 'soft';
    bevel?: {
        type?: 'relaxed' | 'slope' | 'angle' | 'circle' | 'convex';
        width?: number;
        height?: number;
    };
    contour?: {
        color?: string;
        width?: number;
    };
}
export interface PptStyle extends PptTextStyle {
    backgroundColor?: string | PptFill;
    borderColor?: string;
    borderWidth?: number;
    borderStyle?: 'solid' | 'dashed' | 'dotted' | 'double';
    border?: PptBorder;
    fill?: PptFill;
    shadow?: PptShadow;
    reflection?: PptReflection;
    glow?: PptGlow;
    effect3d?: PptEffect3D;
    opacity?: number;
    zIndex?: number;
}
export interface PptTextParagraph {
    text: string;
    style?: Partial<PptTextStyle>;
    bullet?: {
        type?: 'none' | 'bullet' | 'numbered';
        char?: string;
        level?: number;
    };
    hyperlink?: {
        url: string;
        tooltip?: string;
    };
}
export type PptTextContent = PptTextParagraph[];
export interface PptImageContent {
    src: string;
    alt?: string;
    mimeType?: string;
}
export type PptShapeType = 'rectangle' | 'roundRectangle' | 'ellipse' | 'circle' | 'triangle' | 'diamond' | 'star' | 'arrow' | 'line' | 'curve' | 'polygon' | 'custom';
export interface PptShapeContent {
    shapeType: PptShapeType;
    text?: string | PptTextContent;
    path?: string;
    roundedCorners?: number;
}
export interface PptTableCellStyle {
    backgroundColor?: string | PptFill;
    borderColor?: string;
    borderWidth?: number;
    verticalAlign?: 'top' | 'middle' | 'bottom';
    padding?: {
        top?: number;
        bottom?: number;
        left?: number;
        right?: number;
    };
}
export interface PptTableCell {
    text: string;
    style?: PptTableCellStyle;
    colspan?: number;
    rowspan?: number;
}
export type PptTableContent = PptTableCell[][];
export type PptChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area' | 'radar' | 'bubble';
export interface PptChartSeries {
    name: string;
    data: number[];
    color?: string;
}
export interface PptChartContent {
    chartType: PptChartType;
    title?: string;
    categories: string[];
    series: PptChartSeries[];
    showLegend?: boolean;
    showDataLabels?: boolean;
    showGrid?: boolean;
}
export interface PptVideoContent {
    src: string;
    poster?: string;
    mimeType?: string;
    autoplay?: boolean;
    loop?: boolean;
    muted?: boolean;
    controls?: boolean;
}
export interface PptAudioContent {
    src: string;
    mimeType?: string;
    autoplay?: boolean;
    loop?: boolean;
    volume?: number;
}
export type PptMediaContent = PptVideoContent | PptAudioContent;
export interface PptConnectorStyle {
    startArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
    endArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
    lineType?: 'straight' | 'elbow' | 'curved';
}
export interface PptConnectorContent {
    startElementId?: string;
    endElementId?: string;
    startPoint?: {
        x: number;
        y: number;
    };
    endPoint?: {
        x: number;
        y: number;
    };
    style?: PptConnectorStyle;
}
export type PptSmartArtType = 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid' | 'timeline';
export interface PptSmartArtContent {
    smartArtType: PptSmartArtType;
    nodes: Array<{
        text: string;
        children?: PptSmartArtContent['nodes'];
        level?: number;
    }>;
}
export interface PptEquationContent {
    latex?: string;
    mathML?: string;
    image?: string;
}
export type PptElementContent = string | PptTextContent | PptImageContent | PptShapeContent | PptTableContent | PptChartContent | PptMediaContent | PptConnectorContent | PptSmartArtContent | PptEquationContent;
export interface PptElement {
    id: string;
    type: PptNodeType;
    rect: PptRect;
    transform?: PptTransform;
    style: PptStyle;
    content: PptElementContent;
    props: Record<string, unknown>;
    children?: PptElement[];
    parentId?: string;
}
export interface PptSlideTransition {
    type?: 'none' | 'fade' | 'slide' | 'push' | 'wipe' | 'zoom';
    duration?: number;
    direction?: 'left' | 'right' | 'up' | 'down';
}
export type PptSlideLayout = 'blank' | 'title' | 'titleOnly' | 'titleAndContent' | 'sectionHeader' | 'twoContent' | 'comparison' | 'verticalText' | 'contentWithCaption';
export interface PptSlide {
    id: string;
    title: string;
    bgColor: string | PptFill;
    backgroundImage?: string;
    elements: PptElement[];
    props: {
        width: number;
        height: number;
        slideLayout: PptSlideLayout;
        transition?: PptSlideTransition;
        notes?: string;
        slideNumber?: number;
    };
}
export interface PptTheme {
    name?: string;
    colors?: {
        background?: string;
        text?: string;
        accent1?: string;
        accent2?: string;
        accent3?: string;
        accent4?: string;
        accent5?: string;
        accent6?: string;
    };
    fonts?: {
        heading?: string;
        body?: string;
    };
}
export interface PptDocument {
    id: string;
    title: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    created?: string;
    modified?: string;
    slides: PptSlide[];
    theme?: PptTheme;
    props: {
        width: number;
        height: number;
        ratio: number;
        pageSize?: '4:3' | '16:9' | '16:10' | 'custom';
    };
}
export interface XmlSlide {
    xml: string;
    slideId: string;
    rId: string;
    layout: string;
}
export interface XmlElement {
    tag: string;
    attrs: Record<string, string>;
    children: XmlElement[];
    text: string;
}
export interface Relationship {
    id: string;
    type: string;
    target: string;
}
export interface RelationshipMap {
    [key: string]: Relationship;
}
export interface RelsMap {
    [key: string]: {
        id: string;
        type: string;
        target: string;
    };
}
export interface SlideParseResult {
    slideId: string;
    slideLayout?: string;
    elements: any[];
    width?: number;
    height?: number;
}
export interface ParsedShapeElement {
    id: string;
    type: 'shape';
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    shapeType?: string;
    placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    text?: string;
    attrs?: Record<string, string>;
    rawNode?: Element;
}
export interface ParsedImageElement {
    id: string;
    type: 'image';
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    src?: string;
    mediaType?: 'image' | 'video' | 'audio';
    relId?: string;
    attrs?: Record<string, string>;
    rawNode?: Element;
}
export interface ParsedChartElement {
    id: string;
    type: 'chart';
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    relId?: string;
    chartType?: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
    attrs?: Record<string, string>;
    rawNode?: Element;
}
export interface ParsedOleElement {
    id: string;
    type: 'ole';
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    src?: string;
    relId?: string;
    progId?: string;
    hasFallback?: boolean;
    attrs?: Record<string, string>;
    rawNode?: Element;
}
export interface ParsedGroupElement {
    id: string;
    type: 'group';
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    children?: any[];
    childOffset?: {
        x: number;
        y: number;
    };
    attrs?: Record<string, string>;
    rawNode?: Element;
}
