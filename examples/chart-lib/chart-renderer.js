/**
 * Chart Renderer Module
 *
 * 这个模块负责在浏览器中使用 NVD3/D3 渲染图表
 * 从 PPTX 解析器中提取的图表数据和样式将被用来创建交互式图表
 *
 * 使用方法:
 * import { ChartRenderer } from './chart-renderer.js';
 * ChartRenderer.renderCharts(charts, container);
 */

/**
 * Chart Renderer 类
 */
export class ChartRenderer {
    constructor() {
        this.chartInstances = new Map(); // 存储图表实例
    }

    /**
     * 渲染所有图表
     * @param {Array} charts - 图表数据数组，从 PPTX 解析器返回
     * @param {HTMLElement} container - 容器元素（可选）
     */
    renderCharts(charts, container = null) {
        if (!charts || charts.length === 0) {
            console.log('No charts to render');
            return;
        }

        console.log(`Rendering ${charts.length} chart(s)...`);

        charts.forEach(chartInfo => {
            try {
                this.renderChart(chartInfo);
            } catch (err) {
                console.error(`Error rendering chart ${chartInfo.chartId}:`, err);
            }
        });
    }

    /**
     * 渲染单个图表
     * @param {Object} chartInfo - 图表信息对象
     * @param {string} chartInfo.chartId - 图表容器 ID
     * @param {string} chartInfo.type - 图表类型
     * @param {Array} chartInfo.data - 图表数据
     * @param {Object} chartInfo.style - 图表样式
     */
    renderChart(chartInfo) {
        const chartElement = d3.select("#" + chartInfo.chartId);
        if (chartElement.empty()) {
            console.warn(`Chart element not found: #${chartInfo.chartId}`);
            return;
        }

        if (!chartInfo.data) {
            console.warn(`Chart data is null or undefined for chart: ${chartInfo.chartId}`);
            return;
        }

        // 准备图表数据
        const preparedData = this.prepareChartData(chartInfo);
        if (!preparedData) {
            return;
        }

        // 创建图表
        const chart = this.createChart(chartInfo);
        if (!chart) {
            return;
        }

        // 应用样式
        this.applyChartStyles(chart, chartInfo, preparedData);

        // 渲染图表
        this.renderChartElement(chart, chartElement, preparedData);

        // 存储图表实例以便后续更新
        this.chartInstances.set(chartInfo.chartId, chart);

        console.log(`Chart ${chartInfo.chartId} rendered successfully`);
    }

    /**
     * 准备图表数据
     * @param {Object} chartInfo - 图表信息
     * @returns {Array} 准备好的数据
     */
    prepareChartData(chartInfo) {
        const chartData = chartInfo.data;
        const isPieChart = chartInfo.type === "pieChart" || chartInfo.type === "pie3DChart";

        if (!Array.isArray(chartData)) {
            console.error(`Chart data is not an array for chart ${chartInfo.chartId}`);
            return null;
        }

        // 饼图需要特殊的数据格式转换
        if (isPieChart) {
            console.log("Converting pie chart data");
            if (chartData.length > 0 && chartData[0]?.values) {
                const series = chartData[0];
                if (Array.isArray(series.values)) {
                    return series.values.map((item, index) => {
                        let label = `Item ${index}`;
                        if (series.xlabels && series.xlabels[index] !== undefined) {
                            label = series.xlabels[index];
                        }
                        let value = 0;
                        if (item && item.y !== undefined) {
                            value = parseFloat(item.y);
                        }
                        return { x: label, y: value };
                    });
                }
            }
            return [];
        } else {
            // 非饼图使用原始数据格式
            return chartData;
        }
    }

    /**
     * 创建图表实例
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} NVD3 图表实例
     */
    createChart(chartInfo) {
        let chart = null;
        const legendPosition = this.getLegendPosition(chartInfo);

        switch (chartInfo.type) {
            case "lineChart":
                chart = nv.models.lineChart().useInteractiveGuideline(true);
                this.configureChartMargins(chart, legendPosition, chartInfo.type);
                break;

            case "barChart":
                chart = nv.models.multiBarChart();
                this.configureChartMargins(chart, legendPosition, chartInfo.type);
                break;

            case "pieChart":
            case "pie3DChart":
                chart = nv.models.pieChart()
                    .x(d => d.x)
                    .y(d => d.y)
                    .showLabels(true);
                break;

            case "areaChart":
                chart = nv.models.stackedAreaChart()
                    .clipEdge(true)
                    .useInteractiveGuideline(true);
                this.configureChartMargins(chart, legendPosition, chartInfo.type);
                break;

            case "scatterChart":
                chart = nv.models.scatterChart()
                    .showDistX(true)
                    .showDistY(true)
                    .color(d3.scale.category10().range());
                chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
                chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
                this.configureChartMargins(chart, legendPosition, chartInfo.type);
                break;

            default:
                console.warn(`Unknown chart type: ${chartInfo.type}`);
                return null;
        }

        return chart;
    }

    /**
     * 配置图表边距和图例位置
     * @param {Object} chart - 图表实例
     * @param {string} legendPosition - 图例位置
     * @param {string} chartType - 图表类型
     */
    configureChartMargins(chart, legendPosition, chartType) {
        // 饼图不支持图例位置配置
        if (chartType === "pieChart" || chartType === "pie3DChart") {
            return;
        }

        const marginConfig = {
            right: { top: 30, right: 100, bottom: 50, left: 60 },
            left: { top: 30, right: 50, bottom: 50, left: 100 },
            top: { top: 50, right: 50, bottom: 50, left: 60 },
            bottom: { top: 30, right: 50, bottom: 80, left: 60 }
        };

        const margins = marginConfig[legendPosition] || marginConfig.right;

        if (legendPosition === 'right') {
            chart.legend.rightAlign(true).align(true).height(30);
        } else if (legendPosition === 'left') {
            chart.legend.rightAlign(false).align(false).height(30);
        } else if (legendPosition === 'top') {
            chart.legend.rightAlign(false).align(false).height(30);
        } else if (legendPosition === 'bottom') {
            chart.legend.rightAlign(false).align(false).height(30);
        }

        chart.margin(margins);
    }

    /**
     * 获取图例位置
     * @param {Object} chartInfo - 图表信息
     * @returns {string} 图例位置
     */
    getLegendPosition(chartInfo) {
        if (!chartInfo.style?.legend?.position) {
            return 'right';
        }

        const positionMap = {
            'b': 'bottom',
            't': 'top',
            'l': 'left',
            'r': 'right',
            'tr': 'right',
            'tl': 'left',
            'br': 'right',
            'bl': 'left'
        };

        return positionMap[chartInfo.style.legend.position] || 'right';
    }

    /**
     * 应用图表样式
     * @param {Object} chart - 图表实例
     * @param {Object} chartInfo - 图表信息
     * @param {Array} chartData - 图表数据
     */
    applyChartStyles(chart, chartInfo, chartData) {
        const chartStyle = chartInfo.style;
        const isPieChart = chartInfo.type === "pieChart" || chartInfo.type === "pie3DChart";

        // 应用系列颜色
        this.applySeriesColors(chart, chartData, chartStyle, isPieChart);

        // 应用图表区域样式
        this.applyChartAreaStyle(chartStyle);

        // 应用轴样式
        this.applyAxisStyles(chartStyle);

        // 应用标题样式
        this.applyTitleStyle(chartStyle);

        // 应用图例样式
        this.applyLegendStyle(chartStyle);
    }

    /**
     * 应用系列颜色
     * @param {Object} chart - 图表实例
     * @param {Array} chartData - 图表数据
     * @param {Object} chartStyle - 图表样式
     * @param {boolean} isPieChart - 是否为饼图
     */
    applySeriesColors(chart, chartData, chartStyle, isPieChart) {
        // 提取系列样式
        const seriesStyles = [];
        if (Array.isArray(chartData) && !isPieChart) {
            chartData.forEach((series, index) => {
                seriesStyles[index] = series.style || {};
            });
        } else if (isPieChart && chartStyle) {
            seriesStyles[0] = chartStyle;
        }

        // 设置颜色
        if (chart.color) {
            chart.color((d, i) => {
                const style = seriesStyles[i] || seriesStyles[0];
                if (style) {
                    if (style.gradientFill?.color?.length > 0) {
                        return "#" + style.gradientFill.color[0];
                    } else if (style.fillColor) {
                        return style.fillColor;
                    }
                }
                return d3.scale.category10().range()[i % 10];
            });
        }
    }

    /**
     * 应用图表区域样式
     * @param {Object} chartStyle - 图表样式
     */
    applyChartAreaStyle(chartStyle) {
        // 样式在渲染时应用到 SVG 元素
        // 这里只是记录样式信息
        if (chartStyle?.chartArea) {
            console.log("Chart area style:", chartStyle.chartArea);
        }
    }

    /**
     * 应用轴样式
     * @param {Object} chartStyle - 图表样式
     */
    applyAxisStyles(chartStyle) {
        if (chartStyle?.categoryAxis) {
            console.log("Category axis style:", chartStyle.categoryAxis);
        }
        if (chartStyle?.valueAxis) {
            console.log("Value axis style:", chartStyle.valueAxis);
        }
    }

    /**
     * 应用标题样式
     * @param {Object} chartStyle - 图表样式
     */
    applyTitleStyle(chartStyle) {
        if (chartStyle?.title) {
            console.log("Title style:", chartStyle.title);
        }
    }

    /**
     * 应用图例样式
     * @param {Object} chartStyle - 图表样式
     */
    applyLegendStyle(chartStyle) {
        if (chartStyle?.legend) {
            console.log("Legend style:", chartStyle.legend);
        }
    }

    /**
     * 渲染图表元素到 DOM
     * @param {Object} chart - 图表实例
     * @param {Object} chartElement - D3 选择器
     * @param {Array} chartData - 图表数据
     */
    renderChartElement(chart, chartElement, chartData) {
        // 移除已存在的 SVG
        const existingSvg = chartElement.select("svg");
        if (!existingSvg.empty()) {
            existingSvg.remove();
        }

        // 创建新的 SVG 元素
        chartElement.append("svg")
            .datum(chartData)
            .transition().duration(500)
            .call(chart);

        // 启用窗口大小调整
        nv.utils.windowResize(chart.update);

        // 应用图表区域的背景和边框
        this.applyChartAreaStylesToSvg(chartElement);
    }

    /**
     * 应用图表区域样式到 SVG 元素
     * @param {Object} chartElement - D3 选择器
     */
    applyChartAreaStylesToSvg(chartElement) {
        const svgElement = chartElement.select("svg");

        // 样式会在图表数据中传递，但需要在渲染后应用
        // 由于样式信息不在这个函数中直接可用，
        // 实际使用时应该通过参数传递或从外部应用
    }

    /**
     * 更新图表数据
     * @param {string} chartId - 图表 ID
     * @param {Array} newData - 新数据
     */
    updateChart(chartId, newData) {
        const chart = this.chartInstances.get(chartId);
        if (!chart) {
            console.warn(`Chart ${chartId} not found`);
            return;
        }

        // 更新数据
        const chartElement = d3.select("#" + chartId);
        chartElement.datum(newData).transition().duration(500).call(chart);
        chart.update();
    }

    /**
     * 销毁图表
     * @param {string} chartId - 图表 ID
     */
    destroyChart(chartId) {
        const chart = this.chartInstances.get(chartId);
        if (chart) {
            // NVD3 图表会自动清理，但我们可以从 Map 中移除
            this.chartInstances.delete(chartId);
            console.log(`Chart ${chartId} destroyed`);
        }
    }

    /**
     * 销毁所有图表
     */
    destroyAllCharts() {
        this.chartInstances.forEach((chart, chartId) => {
            this.destroyChart(chartId);
        });
    }
}

// 导出单例实例，方便使用
export const chartRenderer = new ChartRenderer();

// 也导出类，允许创建多个实例
export default ChartRenderer;
