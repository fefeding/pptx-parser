import type { BaseElement } from '../elements/BaseElement';
export interface RelsMap {
    [relId: string]: {
        id: string;
        type: string;
        target: string;
    };
}
export interface ParseOptions {
    parseImages?: boolean;
    keepRawXml?: boolean;
    verbose?: boolean;
}
export interface Metadata {
    title?: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    created?: string;
    modified?: string;
}
export interface SlideSize {
    width: number;
    height: number;
}
export interface SlideProps {
    width: number;
    height: number;
    ratio: number;
    pageSize: '4:3' | '16:9' | '16:10' | 'custom';
}
export interface PptxParseResult {
    id: string;
    title: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    created?: string;
    modified?: string;
    slides: SlideParseResult[];
    props: SlideProps;
    globalRelsMap: RelsMap;
    theme?: ThemeResult;
    masterSlides?: MasterSlideResult[];
    slideLayouts?: Record<string, SlideLayoutResult>;
    notesMasters?: NotesMasterResult[];
    notesSlides?: NotesSlideResult[];
    charts?: ChartResult[];
    diagrams?: DiagramResult[];
    tags?: TagsResult[];
    mediaMap?: Map<string, string>;
}
export type PptDocument = PptxParseResult;
export interface SlideLayoutResult {
    id: string;
    name?: string;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
        schemeRef?: string;
    };
    elements: any[];
    placeholders?: Placeholder[];
    relsMap: RelsMap;
    colorMap?: Record<string, string>;
    textStyles?: TextStyles;
    masterRef?: string;
    master?: MasterSlideResult;
}
export interface Placeholder {
    id: string;
    type: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    name?: string;
    rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    hAlign?: 'left' | 'center' | 'right';
    vAlign?: 'top' | 'middle' | 'bottom';
    idx?: number;
    rawNode?: Element;
}
export interface ThemeResult {
    colors: ThemeColors;
}
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
export interface MasterSlideResult {
    id: string;
    masterId?: string;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
        schemeRef?: string;
    };
    elements: any[];
    placeholders?: any[];
    colorMap: Record<string, string>;
    textStyles?: TextStyles;
    themeRef?: string;
    relsMap?: any;
}
export interface TextStyles {
    titleParaPr?: any;
    bodyPr?: any;
    otherPr?: any;
}
export interface Background {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
}
export interface SlideParseResult {
    id: string;
    title: string;
    background?: string | Background;
    elements: BaseElement[];
    relsMap: RelsMap;
    rawXml?: string;
    index?: number;
    layoutId?: string;
    layout?: SlideLayoutResult;
    master?: MasterSlideResult;
    styleApplied?: boolean;
}
export interface Relationship {
    id: string;
    type: string;
    target: string;
}
export interface ChartSeries {
    name?: string;
    idx?: number;
    order?: number;
    points?: ChartDataPoint[];
    color?: string;
}
export type ChartSeriesData = ChartSeries;
export interface ChartDataPoint {
    idx?: number;
    value?: number;
    category?: string;
}
export interface ChartResult {
    id: string;
    chartType: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
    title?: string;
    series?: ChartSeries[];
    categories?: string[];
    xTitle?: string;
    yTitle?: string;
    showLegend?: boolean;
    showDataLabels?: boolean;
    relsMap: RelsMap;
}
export interface DiagramShape {
    id: string;
    type: string;
    position?: {
        x: number;
        y: number;
    };
    size?: {
        width: number;
        height: number;
    };
    text?: string;
}
export type DiagramShapeData = DiagramShape;
export interface DiagramResult {
    id: string;
    diagramType?: string;
    layout?: string;
    colors?: Record<string, string>;
    data?: Record<string, any>;
    shapes?: DiagramShape[];
    relsMap: RelsMap;
}
export interface NotesPlaceholder {
    id: string;
    type: 'header' | 'body' | 'dateTime' | 'slideImage' | 'footer' | 'other';
    name?: string;
    rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    };
}
export interface NotesMasterResult {
    id: string;
    elements: any[];
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    placeholders?: NotesPlaceholder[];
    relsMap: RelsMap;
}
export interface NotesSlideResult {
    id: string;
    slideId?: string;
    text?: string;
    elements: any[];
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    relsMap: RelsMap;
    masterRef?: string;
    master?: NotesMasterResult;
}
export interface SlideTag {
    name: string;
    value: string;
}
export interface ExtensionData {
    uri?: string;
    data?: any;
}
export interface CustomProperty {
    name: string;
    value: any;
    type?: 'string' | 'number' | 'boolean' | 'date';
}
export interface TagsResult {
    id: string;
    slideId?: string;
    tags: SlideTag[];
    extensions: ExtensionData[];
    customProperties: CustomProperty[];
    relsMap: RelsMap;
}
