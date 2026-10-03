/**
 * chart-renderer.js 的类型声明
 * 该模块依赖全局 echarts，按 ECharts 渲染从 PPTX 解析出的图表数据。
 */

/** 解析器返回的图表信息 */
export interface ChartInfo {
    /** 图表容器 ID（对应幻灯片 HTML 中占位 div 的 id） */
    chartId: string;
    /** 图表类型 */
    type: string;
    /** 图表数据 */
    data?: unknown;
    /** 图表样式 */
    style?: Record<string, unknown>;
    [key: string]: unknown;
}

export class ChartRenderer {
    /** chartId -> echarts 实例 */
    chartInstances: Map<string, unknown>;

    /**
     * 渲染所有图表
     * @param charts - 解析器返回的图表数据数组
     * @param container - 作用域容器；传入后只在该容器内查找图表占位元素
     *                    （同一份幻灯片 HTML 被渲染多份时必须传入，否则会互相串台）
     */
    renderCharts(charts: ChartInfo[], container?: HTMLElement | null): void;

    /**
     * 渲染单个图表
     * @param chartInfo - 图表信息
     * @param container - 作用域容器，为 null 时回退到 document
     */
    renderChart(chartInfo: ChartInfo, container?: HTMLElement | null): void;

    /**
     * 在指定容器内按 ID 查找图表占位元素
     */
    findChartElement(chartId: string, container?: HTMLElement | null): HTMLElement | null;
}

/** 默认共享实例 */
export const chartRenderer: ChartRenderer;

export default ChartRenderer;
