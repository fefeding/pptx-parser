type DocInput = ArrayBuffer | ArrayBufferView | Uint8Array | Blob | string;

type AssetScope = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'bundle';
type AssetAnchorKind = 'block' | 'cell' | 'slide' | 'page' | 'other';
interface AssetAnchor {
    scope: AssetScope;
    kind: AssetAnchorKind;
    index: number;
    path?: string;
    label?: string;
}
interface DocAsset {
    id: string;
    name: string;
    mime: string;
    data: Uint8Array;
    part?: string;
    anchor: AssetAnchor;
    width?: number;
    height?: number;
    alt?: string;
    meta?: Record<string, unknown>;
}

interface PdfFontSet {
    regular?: string | ArrayBuffer | Uint8Array;
    bold?: string | ArrayBuffer | Uint8Array;
    italic?: string | ArrayBuffer | Uint8Array;
    boldItalic?: string | ArrayBuffer | Uint8Array;
}

interface ConvertOptions {
    title?: string;
    fonts?: PdfFontSet;
    autoFont?: boolean;
    assets?: DocAsset[];
}
interface Converter<Id extends string = string> {
    readonly id: Id;
    readonly ext: string;
    readonly label?: string;
    convert(input: DocInput, options?: ConvertOptions): Promise<Uint8Array | string>;
}
declare class ConverterRegistry {
    private readonly map;
    register(converter: Converter): this;
    get(id: string): Converter | undefined;
    has(id: string): boolean;
    list(): Converter[];
    remove(id: string): boolean;
}
declare const converters: ConverterRegistry;
type ConvertTarget = 'docx' | 'pdf' | 'md' | 'html';
type ConvertPptxOptions = ConvertOptions;
declare function registerConverter(converter: Converter): void;
declare function convert(input: DocInput, target: string, options?: ConvertOptions): Promise<Uint8Array | string>;
declare function convertPptx(input: DocInput, target: ConvertTarget, options?: ConvertOptions): Promise<Uint8Array | string>;

export { convert, convertPptx, converters, registerConverter };
export type { ConvertOptions, ConvertPptxOptions, ConvertTarget, Converter, DocAsset, DocInput, PdfFontSet };
