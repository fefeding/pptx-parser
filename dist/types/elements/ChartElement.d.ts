import { BaseElement } from './BaseElement';
import type { ParsedChartElement, RelsMap } from '../types';
export interface ChartDataPoint {
    value?: number;
    category?: string;
}
export interface ChartSeries {
    name?: string;
    points?: ChartDataPoint[];
    color?: string;
}
export interface ChartData {
    type: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart';
    title?: string;
    xTitle?: string;
    yTitle?: string;
    series?: ChartSeries[];
}
export declare class ChartElement extends BaseElement {
    type: "chart";
    chartType?: 'lineChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'areaChart' | 'scatterChart' | 'unknown';
    chartData?: ChartData;
    relId: string;
    static fromNode(node: Element, relsMap: RelsMap): ChartElement | null;
    private detectChartType;
    toHTML(): string;
    private getChartLabel;
    private renderChartData;
    toParsedElement(): ParsedChartElement;
}
