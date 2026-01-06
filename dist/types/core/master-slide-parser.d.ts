import JSZip from 'jszip';
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
    themeRef?: string;
    relsMap?: any;
}
export declare function parseAllMasterSlides(zip: JSZip): Promise<MasterSlideResult[]>;
