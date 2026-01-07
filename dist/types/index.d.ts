declare function emu2px(emu: string | number): number;
declare function px2emu(px: number): number;

declare function getAttrs(node: Element): Record<string, string>;

type PptNodeType = 'text' | 'image' | 'shape' | 'table' | 'chart' | 'container' | 'media' | 'video' | 'audio' | 'line' | 'connector' | 'group' | 'smartart' | 'equation';
interface PptRect {
    x: number;
    y: number;
    width: number;
    height: number;
}
type Position = PptRect;
interface PptTransform {
    rotate?: number;
    flipH?: boolean;
    flipV?: boolean;
}
interface PptTextStyle {
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
interface PptFill {
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
interface PptBorder {
    color?: string;
    width?: number;
    style?: 'solid' | 'dashed' | 'dotted' | 'double';
    dashStyle?: string;
}
interface PptShadow {
    color?: string;
    blur?: number;
    offsetX?: number;
    offsetY?: number;
    opacity?: number;
}
interface PptReflection {
    blur?: number;
    opacity?: number;
    offset?: number;
}
interface PptGlow {
    color?: string;
    radius?: number;
    opacity?: number;
}
interface PptEffect3D {
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
interface PptStyle extends PptTextStyle {
    backgroundColor?: string | PptFill;
    background?: string;
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
    padding?: string;
    paddingTop?: number;
    paddingBottom?: number;
    paddingLeft?: number;
    paddingRight?: number;
    marginLeft?: number;
    marginRight?: number;
    spaceBefore?: number;
    spaceAfter?: number;
}
interface PptTextParagraph {
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
type PptTextContent = PptTextParagraph[];
interface PptImageContent {
    src: string;
    alt?: string;
    mimeType?: string;
}
type PptShapeType = 'rectangle' | 'roundRectangle' | 'ellipse' | 'circle' | 'triangle' | 'diamond' | 'star' | 'arrow' | 'line' | 'curve' | 'polygon' | 'custom';
interface PptShapeContent {
    shapeType: PptShapeType;
    text?: string | PptTextContent;
    path?: string;
    roundedCorners?: number;
}
interface PptTableCellStyle {
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
interface PptTableCell {
    text: string;
    style?: PptTableCellStyle;
    colspan?: number;
    rowspan?: number;
}
type PptTableContent = PptTableCell[][];
type PptChartType = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'scatter' | 'area' | 'radar' | 'bubble';
interface PptChartSeries {
    name: string;
    data: number[];
    color?: string;
}
interface PptChartContent {
    chartType: PptChartType;
    title?: string;
    categories: string[];
    series: PptChartSeries[];
    showLegend?: boolean;
    showDataLabels?: boolean;
    showGrid?: boolean;
}
interface PptVideoContent {
    src: string;
    poster?: string;
    mimeType?: string;
    autoplay?: boolean;
    loop?: boolean;
    muted?: boolean;
    controls?: boolean;
}
interface PptAudioContent {
    src: string;
    mimeType?: string;
    autoplay?: boolean;
    loop?: boolean;
    volume?: number;
}
type PptMediaContent = PptVideoContent | PptAudioContent;
interface PptConnectorStyle {
    startArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
    endArrow?: 'none' | 'arrow' | 'stealth' | 'diamond' | 'oval';
    lineType?: 'straight' | 'elbow' | 'curved';
}
interface PptConnectorContent {
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
type PptSmartArtType = 'process' | 'cycle' | 'hierarchy' | 'relationship' | 'matrix' | 'pyramid' | 'timeline';
interface PptSmartArtContent {
    smartArtType: PptSmartArtType;
    nodes: Array<{
        text: string;
        children?: PptSmartArtContent['nodes'];
        level?: number;
    }>;
}
interface PptEquationContent {
    latex?: string;
    mathML?: string;
    image?: string;
}
type PptElementContent = string | PptTextContent | PptImageContent | PptShapeContent | PptTableContent | PptChartContent | PptMediaContent | PptConnectorContent | PptSmartArtContent | PptEquationContent;
interface PptElement {
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
interface PptSlideTransition {
    type?: 'none' | 'fade' | 'slide' | 'push' | 'wipe' | 'zoom';
    duration?: number;
    direction?: 'left' | 'right' | 'up' | 'down';
}
type PptSlideLayout = 'blank' | 'title' | 'titleOnly' | 'titleAndContent' | 'sectionHeader' | 'twoContent' | 'comparison' | 'verticalText' | 'contentWithCaption';
interface PptSlide {
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
interface PptTheme {
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
interface PptDocument {
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
interface RelsMap$1 {
    [key: string]: {
        id: string;
        type: string;
        target: string;
    };
}
interface ParsedShapeElement {
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
interface ParsedImageElement {
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
interface ParsedChartElement {
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
interface ParsedOleElement {
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
interface ParsedGroupElement {
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

declare abstract class BaseElement {
    id: string;
    abstract type: string;
    rect: PptRect;
    style: PptStyle;
    content: any;
    props: any;
    name?: string;
    hidden?: boolean;
    rawNode?: Element;
    isPlaceholder?: boolean;
    protected relsMap: Record<string, any>;
    constructor(id: string, type: string, rect: PptRect, content?: any, props?: any, relsMap?: Record<string, any>);
    abstract toHTML(): string;
    protected getContainerStyle(): string;
    protected getDataAttributes(): Record<string, string>;
    protected formatDataAttributes(): string;
    protected parsePosition(node: Element, tag?: string, namespace?: "http://schemas.openxmlformats.org/presentationml/2006/main"): Position;
    protected parseIdAndName(node: Element, nonVisualTag: string, namespace?: "http://schemas.openxmlformats.org/presentationml/2006/main"): {
        id: string;
        name: string;
        hidden: boolean;
    };
    protected generateId(): string;
    protected getAttributes(node: Element): Record<string, string>;
}

interface RelsMap {
    [relId: string]: {
        id: string;
        type: string;
        target: string;
    };
}
interface ParseOptions {
    parseImages?: boolean;
    keepRawXml?: boolean;
    verbose?: boolean;
}
interface SlideProps {
    width: number;
    height: number;
    ratio: number;
    pageSize: '4:3' | '16:9' | '16:10' | 'custom';
}
interface PptxParseResult {
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
interface SlideLayoutResult {
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
interface Placeholder {
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
interface ThemeResult {
    colors: ThemeColors;
    name?: string;
}
interface ThemeColors {
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
interface MasterSlideResult {
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
interface TextStyles {
    titleParaPr?: any;
    bodyPr?: any;
    otherPr?: any;
}
interface Background {
    type: 'color' | 'image' | 'none';
    value?: string;
    relId?: string;
    schemeRef?: string;
}
interface SlideParseResult {
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
interface ChartSeries$1 {
    name?: string;
    idx?: number;
    order?: number;
    points?: ChartDataPoint$1[];
    color?: string;
}
interface ChartDataPoint$1 {
    idx?: number;
    value?: number;
    category?: string;
}
interface ChartResult {
    id: string;
    chartType: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
    title?: string;
    series?: ChartSeries$1[];
    categories?: string[];
    xTitle?: string;
    yTitle?: string;
    showLegend?: boolean;
    showDataLabels?: boolean;
    relsMap: RelsMap;
}
interface DiagramShape$1 {
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
interface DiagramResult {
    id: string;
    diagramType?: string;
    layout?: string;
    colors?: Record<string, string>;
    data?: Record<string, any>;
    shapes?: DiagramShape$1[];
    relsMap: RelsMap;
}
interface NotesPlaceholder {
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
interface NotesMasterResult {
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
interface NotesSlideResult {
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
interface SlideTag {
    name: string;
    value: string;
}
interface ExtensionData {
    uri?: string;
    data?: any;
}
interface CustomProperty {
    name: string;
    value: any;
    type?: 'string' | 'number' | 'boolean' | 'date';
}
interface TagsResult {
    id: string;
    slideId?: string;
    tags: SlideTag[];
    extensions: ExtensionData[];
    customProperties: CustomProperty[];
    relsMap: RelsMap;
}

declare function parseRels(relsXml: string): RelsMap;
declare function parseMetadata(coreXml: string): {
    title?: string;
    author?: string;
    created?: string;
    modified?: string;
    subject?: string;
    keywords?: string;
};

declare function generateId(prefix?: string): string;

declare const PptParseUtils: {
    generateId: typeof generateId;
    parseXmlText: (text: string) => string;
    px2emu: typeof px2emu;
    emu2px: typeof emu2px;
    parseXmlToTree: (xmlStr: string) => any;
    parseXmlAttrs: (attrs: NamedNodeMap) => Record<string, string>;
    parseXmlRect: (attrs: Record<string, string>) => {
        x: number;
        y: number;
        width: number;
        height: number;
    };
    parseXmlStyle: (attrs: Record<string, string>) => any;
    hexToRgb: (hex: string) => {
        r: number;
        g: number;
        b: number;
    };
    rgbToHex: (r: number, g: number, b: number) => string;
    parseColor: (color: string) => string;
};
declare function parsePptx(file: File | Blob | ArrayBuffer, options?: ParseOptions & {
    returnFormat?: 'enhanced' | 'simple';
}): Promise<PptxParseResult | PptDocument>;
declare function serializePptx(pptDoc: PptDocument): Promise<Blob>;

interface TextRun {
    text: string;
    fontSize?: number;
    fontFamily?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    strike?: boolean;
    color?: string;
    backgroundColor?: string;
    highlight?: string;
    letterSpacing?: number;
}
interface BulletStyle {
    type?: 'none' | 'char' | 'blip' | 'autoNum';
    char?: string;
    imageSrc?: string;
    autoNumType?: string;
    level?: number;
    color?: string;
    size?: number;
    font?: string;
}
declare class ShapeElement extends BaseElement {
    type: 'shape' | 'text';
    shapeType?: string;
    text?: string;
    textStyle?: TextRun[];
    paragraphStyle?: {
        align?: 'left' | 'center' | 'right' | 'justify';
        indent?: number;
        lineSpacing?: number;
        spaceBefore?: number;
        spaceAfter?: number;
        marginLeft?: number;
        marginRight?: number;
        paddingTop?: number;
        paddingBottom?: number;
        rtl?: boolean;
    };
    bulletStyle?: BulletStyle;
    isPlaceholder?: boolean;
    placeholderType?: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    hyperlink?: {
        id?: string;
        url?: string;
        tooltip?: string;
    };
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    static fromNode(node: Element, relsMap: RelsMap$1): ShapeElement | null;
    private parseShapeProperties;
    private parseFill;
    private parseGradientFill;
    private parseGradientStopColor;
    private generateGradientCSS;
    private parseTextBody;
    private parseParagraph;
    private parseBulletStyle;
    private parseTextRun;
    private parseRunProperties;
    private parseColor;
    private detectShapeType;
    toHTML(): string;
    private getTextStyle;
    private getRotationStyle;
    private getShapeStyle;
    private textStyleFromAlign;
    private textStyleFromFontSize;
    private escapeHtml;
    private parseStyleString;
    toParsedElement(): ParsedShapeElement;
}

type MediaType = 'image' | 'video' | 'audio';
interface VideoInfo {
    src: string;
    type: 'blob' | 'link';
    format?: string;
    autoplay?: boolean;
    muted?: boolean;
    controls?: boolean;
    loop?: boolean;
}
interface AudioInfo {
    src: string;
    type: 'blob' | 'link';
    format?: string;
    autoplay?: boolean;
    muted?: boolean;
    controls?: boolean;
    loop?: boolean;
}
declare class ImageElement extends BaseElement {
    type: "image";
    mediaType: MediaType;
    src: string;
    relId: string;
    mimeType?: string;
    altText?: string;
    videoInfo?: VideoInfo;
    audioInfo?: AudioInfo;
    static fromNode(node: Element, relsMap: RelsMap$1): ImageElement | null;
    private parseVideoInfo;
    private parseAudioInfo;
    private parseImageSrc;
    private detectVideoFormat;
    private detectAudioFormat;
    getFilePath(): string;
    getDataAttributes(): Record<string, string>;
    toHTML(): string;
    private toImageHTML;
    private toVideoHTML;
    private getVideoSourceTag;
    private toAudioHTML;
    private getAudioSourceTag;
    toParsedElement(): ParsedImageElement;
}

declare class OleElement extends BaseElement {
    type: "ole";
    progId?: string;
    relId: string;
    oleName?: string;
    hasFallback?: boolean;
    static fromNode(node: Element, relsMap: RelsMap$1): OleElement | null;
    constructor(id: string, rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, progId: string, relId: string, props?: any, relsMap?: Record<string, any>);
    toHTML(): string;
    toParsedElement(): ParsedOleElement;
}

interface ChartDataPoint {
    value?: number;
    category?: string;
}
interface ChartSeries {
    name?: string;
    points?: ChartDataPoint[];
    color?: string;
}
interface ChartData {
    type: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart';
    title?: string;
    xTitle?: string;
    yTitle?: string;
    series?: ChartSeries[];
}
declare class ChartElement extends BaseElement {
    type: "chart";
    chartType?: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
    chartData?: ChartData;
    relId: string;
    static fromNode(node: Element, relsMap: RelsMap$1): ChartElement | null;
    private detectChartType;
    toHTML(): string;
    private getChartLabel;
    private renderChartData;
    toParsedElement(): ParsedChartElement;
}

interface TableCell {
    text?: string;
    rowSpan?: number;
    colSpan?: number;
    style?: {
        backgroundColor?: string;
        borderColor?: string;
        borderWidth?: string;
        padding?: string;
        verticalAlign?: string;
    };
}
interface TableRow {
    cells: TableCell[];
    height?: number;
    style?: {
        backgroundColor?: string;
        borderColor?: string;
        borderWidth?: string;
        isHeader?: boolean;
        isFooter?: boolean;
    };
}
interface TableStyle {
    firstRow?: {
        backgroundColor?: string;
        fontWeight?: string;
        fontSize?: number;
    };
    lastRow?: {
        backgroundColor?: string;
    };
    firstCol?: {
        backgroundColor?: string;
        fontWeight?: string;
    };
    lastCol?: {
        backgroundColor?: string;
    };
    bandRow?: {
        odd?: {
            backgroundColor?: string;
        };
        even?: {
            backgroundColor?: string;
        };
    };
    bandCol?: {
        odd?: {
            backgroundColor?: string;
        };
        even?: {
            backgroundColor?: string;
        };
    };
}
declare class TableElement extends BaseElement {
    type: "table";
    rows: TableRow[];
    colWidths: number[];
    tableStyle?: TableStyle;
    rtl?: boolean;
    static fromNode(node: Element, relsMap: RelsMap$1): TableElement | null;
    private parseTable;
    private parseTableRow;
    private parseTableCell;
    private parseCellStyle;
    private parseBorder;
    private parseTableStyle;
    private parseRowStyle;
    private parseRowBackgroundColor;
    private parseTableWholeStyle;
    private generateBoxShadow;
    toHTML(): string;
    private getTableStyle;
    private rowToHTML;
    private getRowStyle;
    private cellToHTML;
    private getCellStyle;
    toParsedElement(): any;
}

interface DiagramData {
    colors?: Record<string, string>;
    data?: Record<string, any>;
    layout?: string;
    shapes?: DiagramShape[];
}
interface DiagramShape {
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
declare class DiagramElement extends BaseElement {
    type: "diagram";
    diagramData?: DiagramData;
    relId: string;
    constructor(id: string, rect: PptRect, content?: any, props?: any, relsMap?: Record<string, any>);
    static fromNode(node: Element, relsMap: RelsMap$1): DiagramElement | null;
    private parseDiagramData;
    private parseRelIds;
    private fetchColorData;
    private fetchLayoutData;
    private fetchColorDataSync;
    private fetchLayoutDataSync;
    private parseShapes;
    toHTML(): string;
    private getDiagramInfo;
    toParsedElement(): any;
}

declare class GroupElement extends BaseElement {
    type: "group";
    children: BaseElement[];
    rotation?: number;
    flipH?: boolean;
    flipV?: boolean;
    childOffset?: {
        x: number;
        y: number;
    };
    static fromNode(node: Element, relsMap: RelsMap$1): GroupElement | null;
    private parseGraphicFrame;
    private parseGroupProperties;
    toHTML(): string;
    private getGroupStyle;
    toParsedElement(): ParsedGroupElement;
}

declare class PlaceholderElement extends BaseElement {
    type: 'placeholder';
    placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other';
    idx?: number;
    name?: string;
    text?: string;
    textStyle?: TextRun[];
    constructor(id: string, placeholderType: 'title' | 'body' | 'dateTime' | 'slideNumber' | 'footer' | 'other', rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, props?: any);
    toHTML(): string;
    private getPlaceholderContent;
    private renderTextContent;
    private get isContentSet();
    private getTextRunStyle;
}
declare class LayoutElement extends BaseElement {
    type: 'layout';
    name?: string;
    placeholders: PlaceholderElement[];
    elements: BaseElement[];
    textStyles?: any;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    relsMap: Record<string, any>;
    mediaMap?: Map<string, string>;
    constructor(id: string, name?: string, placeholders?: PlaceholderElement[], elements?: BaseElement[], props?: any);
    static fromResult(result: SlideLayoutResult, mediaMap?: Map<string, string>): LayoutElement;
    toHTML(): string;
    private getBackgroundStyle;
}

declare class MasterElement extends BaseElement {
    type: 'master';
    masterId?: string;
    elements: BaseElement[];
    placeholders: PlaceholderElement[];
    textStyles?: any;
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    colorMap: Record<string, string>;
    mediaMap?: Map<string, string>;
    constructor(id: string, elements?: BaseElement[], placeholders?: PlaceholderElement[], props?: any);
    static fromResult(result: MasterSlideResult, mediaMap?: Map<string, string>): MasterElement;
    toHTML(): string;
    private getBackgroundStyle;
    getPlaceholderStyle(placeholderType: 'title' | 'body' | 'other'): any;
    private parseParagraphProperties;
}

declare class SlideElement {
    id: string;
    title: string;
    background: string;
    elements: BaseElement[];
    rawResult: SlideParseResult;
    layoutElement?: LayoutElement;
    masterElement?: MasterElement;
    layout?: SlideLayoutResult;
    master?: MasterSlideResult;
    mediaMap?: Map<string, string>;
    constructor(result: SlideParseResult, elements: BaseElement[], layoutElement?: LayoutElement, masterElement?: MasterElement, layoutResult?: SlideLayoutResult, masterResult?: MasterSlideResult, mediaMap?: Map<string, string>);
    toHTML(): string;
    private renderMasterElements;
    private renderLayoutElements;
    private renderSlideElementsWithLayout;
    private getTextContentStyle;
    private getSlideBackground;
    toHTMLString(): string;
    private escapeHtml;
}
declare class PptxDocument {
    id: string;
    title: string;
    author?: string;
    slides: SlideElement[];
    width: number;
    height: number;
    ratio: number;
    constructor(id: string, title: string, slides: SlideElement[], width?: number, height?: number, author?: string);
    static fromParseResult(result: any): PptxDocument;
    toHTML(): string;
    toHTMLWithNavigation(): string;
    private escapeHtml;
}

interface FontScheme {
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
interface EffectScheme {
    fillStyles?: any[];
    lineStyles?: any[];
    effectStyles?: any[];
    bgFillStyles?: any[];
}
declare class ThemeElement extends BaseElement {
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

declare class TagsElement extends BaseElement {
    type: 'tags';
    tags: SlideTag[];
    extensions: ExtensionData[];
    customProperties: CustomProperty[];
    slideId?: string;
    constructor(id: string, tags?: SlideTag[], customProperties?: CustomProperty[], extensions?: ExtensionData[], props?: any);
    static fromResult(result: TagsResult): TagsElement;
    toHTML(): string;
    toHTMLDebug(): string;
    getTag(name: string): string | undefined;
    getProperty(name: string): any | undefined;
    setTag(name: string, value: string): void;
    setProperty(name: string, value: any, type?: 'string' | 'number' | 'boolean' | 'date'): void;
}

declare class NotesMasterElement extends BaseElement {
    type: 'notesMaster';
    elements: BaseElement[];
    placeholders: any[];
    background?: {
        type: 'color' | 'image' | 'none';
        value?: string;
        relId?: string;
    };
    mediaMap?: Map<string, string>;
    constructor(id: string, elements?: BaseElement[], placeholders?: any[], props?: any);
    static fromResult(result: NotesMasterResult, mediaMap?: Map<string, string>): NotesMasterElement;
    toHTML(): string;
    private getBackgroundStyle;
}
declare class NotesSlideElement extends BaseElement {
    type: 'notesSlide';
    text?: string;
    masterRef?: string;
    master?: NotesMasterElement;
    slideId?: string;
    constructor(id: string, rect: {
        x: number;
        y: number;
        width: number;
        height: number;
    }, props?: any);
    static fromResult(result: NotesSlideResult): NotesSlideElement;
    toHTML(): string;
    setMaster(master: NotesMasterElement): void;
}

interface HtmlRenderOptions {
    includeStyles?: boolean;
    includeScripts?: boolean;
    includeLayoutElements?: boolean;
    withNavigation?: boolean;
    customCss?: string;
}
declare class DocumentElement extends BaseElement {
    type: 'document';
    title: string;
    author?: string;
    subject?: string;
    keywords?: string;
    description?: string;
    created?: string;
    modified?: string;
    slides: SlideElement[];
    layouts: Record<string, LayoutElement>;
    masters: MasterElement[];
    tags: TagsElement[];
    notesMasters: NotesMasterElement[];
    notesSlides: NotesSlideElement[];
    width: number;
    height: number;
    ratio: number;
    pageSize: '4:3' | '16:9' | '16:10' | 'custom';
    globalRelsMap: Record<string, any>;
    mediaMap?: Map<string, string>;
    theme?: ThemeElement;
    constructor(id: string, title: string, width?: number, height?: number, props?: any);
    static fromParseResult(result: PptxParseResult): DocumentElement;
    toHTML(options?: HtmlRenderOptions): string;
    toHTMLDocument(options?: HtmlRenderOptions): string;
    toHTMLWithNavigation(options?: HtmlRenderOptions): string;
    private generateStyles;
    private generateNavigationStyles;
    getSlide(index: number): SlideElement | undefined;
    getLayout(layoutId: string): LayoutElement | undefined;
    getMaster(masterId: string): MasterElement | undefined;
    private escapeHtml;
}
declare function createDocument(result: PptxParseResult): DocumentElement;

declare function createElementFromData(data: any, relsMap?: Record<string, any>, mediaMap?: Map<string, string>): BaseElement | null;

declare function createElementFromNode(node: Element, relsMap: Record<string, any>): BaseElement | null;

declare const NS: {
    readonly p: "http://schemas.openxmlformats.org/presentationml/2006/main";
    readonly a: "http://schemas.openxmlformats.org/drawingml/2006/main";
    readonly r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    readonly mc: "http://schemas.openxmlformats.org/markup-compatibility/2006";
    readonly c: "http://schemas.openxmlformats.org/drawingml/2006/chart";
    readonly d: "http://schemas.openxmlformats.org/drawingml/2006/diagram";
};
declare const EMU_PER_INCH = 914400;
declare const PIXELS_PER_INCH = 96;

declare const slide2HTML: (slide: any, options?: HtmlRenderOptions) => any;
declare const ppt2HTML: (result: PptxParseResult, options?: HtmlRenderOptions) => any;
declare const ppt2HTMLDocument: (result: PptxParseResult, options?: HtmlRenderOptions) => any;

declare const PptParserCore: {
    utils: {
        generateId: typeof generateId;
        parseXmlText: (text: string) => string;
        px2emu: typeof px2emu;
        emu2px: typeof emu2px;
        parseXmlToTree: (xmlStr: string) => any;
        parseXmlAttrs: (attrs: NamedNodeMap) => Record<string, string>;
        parseXmlRect: (attrs: Record<string, string>) => {
            x: number;
            y: number;
            width: number;
            height: number;
        };
        parseXmlStyle: (attrs: Record<string, string>) => any;
        hexToRgb: (hex: string) => {
            r: number;
            g: number;
            b: number;
        };
        rgbToHex: (r: number, g: number, b: number) => string;
        parseColor: (color: string) => string;
    };
    parse: typeof parsePptx;
    serialize: typeof serializePptx;
};

export { BaseElement, ChartElement, DiagramElement, DocumentElement, EMU_PER_INCH, GroupElement, ImageElement, LayoutElement, MasterElement, NS, NotesMasterElement, NotesSlideElement, OleElement, PIXELS_PER_INCH, PlaceholderElement, PptParseUtils, PptxDocument, ShapeElement, SlideElement, TableElement, TagsElement, createDocument, createElementFromData, createElementFromNode, PptParserCore as default, emu2px, getAttrs, parseMetadata, parsePptx, parseRels, ppt2HTML, ppt2HTMLDocument, px2emu, serializePptx, slide2HTML };
export type { HtmlRenderOptions, ParseOptions, PptDocument, PptElement, PptNodeType, PptRect, PptSlide, PptStyle, PptxParseResult, SlideParseResult };
