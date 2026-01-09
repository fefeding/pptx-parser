import JSZip from 'jszip';
import type { ChartResult, DiagramResult } from './types';
export declare function parseAllCharts(zip: JSZip): Promise<ChartResult[]>;
export declare function parseAllDiagrams(zip: JSZip): Promise<DiagramResult[]>;
