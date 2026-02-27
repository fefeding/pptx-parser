/**
 * Chart Renderer Module
 *
 * 这个模块负责在浏览器中使用 ECharts 渲染图表
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
            return;
        }

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
        const chartElement = document.getElementById(chartInfo.chartId);
        if (!chartElement) {
            console.warn(`Chart element not found: #${chartInfo.chartId}`);
            return;
        }

        if (!chartInfo.data) {
            console.warn(`Chart data is null or undefined for chart: ${chartInfo.chartId}`);
            return;
        }

        // 准备 ECharts 配置
        const option = this.prepareEChartsOption(chartInfo);
        if (!option) {
            return;
        }

        // 创建 ECharts 实例
        const chart = echarts.init(chartElement);

        // 设置配置
        chart.setOption(option);

        // 启用窗口大小调整
        window.addEventListener('resize', () => {
            chart.resize();
        });

        // 存储图表实例以便后续更新
        this.chartInstances.set(chartInfo.chartId, chart);
    }

    /**
     * 准备 ECharts 配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} ECharts 配置对象
     */
    prepareEChartsOption(chartInfo) {
        const chartData = chartInfo.data;
        const chartType = this.mapChartType(chartInfo.type);

        if (!chartType) {
            console.warn(`Unsupported chart type: ${chartInfo.type}`);
            return null;
        }

        const isPieChart = chartType === 'pie';
        const is3DPie = chartInfo.type === 'pie3DChart';
        const option = {
            tooltip: {
                trigger: isPieChart ? 'item' : 'axis'
            },
            legend: this.getLegendConfig(chartInfo),
            series: this.prepareSeries(chartInfo, chartType)
        };

        // 应用图表背景
        this.applyChartBackground(option, chartInfo.style);

        // 添加标题
        if (chartInfo.title || chartInfo.style?.title) {
            option.title = this.getTitleConfig(chartInfo.title, chartInfo.style?.title);
        }

        // 为非饼图添加坐标轴配置
        if (!isPieChart) {
            const xAxisData = this.getXAxisData(chartInfo);
            option.xAxis = this.getXAxisConfig(chartInfo, xAxisData);
            option.yAxis = this.getYAxisConfig(chartInfo);

            // 应用图表区域布局
            option.grid = this.getGridConfig(chartInfo);
        } else if (isPieChart) {
            // 为饼图添加布局配置
            option.grid = {
                top: (chartInfo.title || chartInfo.style?.title) ? '15%' : '10%',
                bottom: chartInfo.style?.legend?.position === 'bottom' ? '15%' : '5%',
                left: '5%',
                right: '5%',
                containLabel: true
            };
        }

        return option;
    }

    /**
     * 应用图表背景和边框
     * @param {Object} option - ECharts 配置对象
     * @param {Object} style - 样式对象
     */
    applyChartBackground(option, style) {
        if (!style) return;

        // 应用图表区域背景
        if (style.chartArea?.fillColor) {
            option.backgroundColor = style.chartArea.fillColor;
        } else if (style.chartArea?.gradientFill) {
            // 应用渐变背景
            const gradFill = style.chartArea.gradientFill;
            if (gradFill.color && gradFill.color.length >= 2) {
                option.backgroundColor = new echarts.graphic.LinearGradient(
                    0, 0, 1, 1,  // 渐变方向
                    [{
                        offset: 0,
                        color: gradFill.color[0].startsWith('#') 
                            ? gradFill.color[0] 
                            : '#' + gradFill.color[0]
                    }, {
                        offset: 1,
                        color: gradFill.color[1] 
                            ? (gradFill.color[1].startsWith('#') 
                                ? gradFill.color[1] 
                                : '#' + gradFill.color[1])
                            : gradFill.color[0].startsWith('#') 
                                ? gradFill.color[0] 
                                : '#' + gradFill.color[0]
                    }]
                );
            }
        }

        // 应用边框（通过 ECharts 的 graphic 组件）
        if (style.chartArea?.borderColor && style.chartArea?.borderWidth) {
            option.graphic = option.graphic || {
                elements: []
            };
            option.graphic.elements.push({
                type: 'rect',
                shape: {
                    x: 0,
                    y: 0,
                    width: '100%',
                    height: '100%'
                },
                style: {
                    fill: 'transparent',
                    stroke: style.chartArea.borderColor,
                    lineWidth: style.chartArea.borderWidth
                },
                z: -1  // 放在最底层
            });
        }
    }

    /**
     * 获取标题配置
     * @param {string} titleText - 标题文本
     * @param {Object} titleStyle - 标题样式
     * @returns {Object} 标题配置
     */
    getTitleConfig(titleText, titleStyle) {
        const config = {
            text: titleText || '',
            left: 'center',
            top: 0,
            textStyle: {}
        };

        if (titleStyle) {
            if (titleStyle.color) {
                config.textStyle.color = titleStyle.color;
            }

            if (titleStyle.fontSize) {
                config.textStyle.fontSize = titleStyle.fontSize;
            }

            if (titleStyle.fontWeight) {
                config.textStyle.fontWeight = titleStyle.fontWeight;
            }
        }

        return config;
    }

    /**
     * 获取网格配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} 网格配置
     */
    getGridConfig(chartInfo) {
        const chartArea = chartInfo.style?.chartArea;
        const hasTitle = chartInfo.title || chartInfo.style?.title;

        // 默认边距 - 减少左右边距让图表铺满
        let top = '15%';
        let bottom = '10%';
        let left = '3%';
        let right = '3%';

        // 根据图例位置调整
        const legendPosition = this.getLegendPosition(chartInfo);
        switch (legendPosition) {
            case 'top':
                top = hasTitle ? '20%' : '15%';
                break;
            case 'bottom':
                bottom = '15%';
                break;
            case 'left':
                left = '12%';
                break;
            case 'right':
                right = '12%';
                break;
        }

        return {
            left: left,
            right: right,
            top: top,
            bottom: bottom,
            containLabel: true
        };
    }

    /**
     * 映射 PPTX 图表类型到 ECharts 类型
     * @param {string} pptxType - PPTX 图表类型
     * @returns {string} ECharts 图表类型
     */
    mapChartType(pptxType) {
        const typeMap = {
            'lineChart': 'line',
            'barChart': 'bar',
            'pieChart': 'pie',
            'pie3DChart': 'pie',
            'areaChart': 'line',
            'scatterChart': 'scatter'
        };
        return typeMap[pptxType] || null;
    }

    /**
     * 准备系列数据
     * @param {Object} chartInfo - 图表信息
     * @param {string} echartsType - ECharts 图表类型
     * @returns {Array} 系列配置数组
     */
    prepareSeries(chartInfo, echartsType) {
        const chartData = chartInfo.data;
        const isPieChart = echartsType === 'pie';
        const isAreaChart = chartInfo.type === 'areaChart';

        if (!Array.isArray(chartData)) {
            console.error('Chart data is not an array');
            return [];
        }

        // 饼图需要特殊的数据格式转换
        if (isPieChart) {
            return this.preparePieSeries(chartData, chartInfo);
        }

        // 散点图需要特殊的数据格式
        if (echartsType === 'scatter') {
            return this.prepareScatterSeries(chartData, chartInfo);
        }

        // 折线图、柱状图、面积图
        return chartData.map((series, index) => {
            // ECharts 对于有 xlabels 的图表，data 需要是数值数组
            // 对于散点图，data 需要是 [x, y] 数组
            let seriesData;
            if (echartsType === 'scatter') {
                // 散点图保持原格式 [x, y]
                seriesData = series.values || [];
            } else if (series.xlabels && series.values) {
                // 有 xlabels 的折线图/柱状图，提取 y 值
                seriesData = series.values.map(v => v ? v.y : 0);
            } else {
                // 没有 xlabels，使用原始数据
                seriesData = series.values || [];
            }

            const seriesConfig = {
                name: series.key || `Series ${index}`,
                type: echartsType,
                data: seriesData,
                smooth: isAreaChart,
                areaStyle: isAreaChart ? {} : undefined
            };

            // 应用系列颜色
            if (series.style) {
                if (series.style.fillColor) {
                    seriesConfig.itemStyle = {
                        color: series.style.fillColor
                    };
                } else if (series.style.gradientFill?.color?.length > 0) {
                    seriesConfig.itemStyle = {
                        color: '#' + series.style.gradientFill.color[0]
                    };
                }
            }

            return seriesConfig;
        });
    }

    /**
     * 准备饼图系列
     * @param {Array} chartData - 图表数据
     * @param {Object} chartInfo - 图表信息
     * @returns {Array} 饼图系列配置
     */
    preparePieSeries(chartData, chartInfo) {
        if (chartData.length === 0) {
            return [];
        }

        const series = chartData[0];
        const chartStyle = chartInfo.style || {};
        const is3D = chartInfo.type === 'pie3DChart';
        const data = [];

        if (Array.isArray(series.values)) {
            series.values.forEach((item, index) => {
                let label = `Item ${index}`;
                if (series.xlabels && series.xlabels[index] !== undefined) {
                    label = series.xlabels[index];
                }
                let value = 0;
                if (item && item.y !== undefined) {
                    value = parseFloat(item.y);
                }

                const dataItem = {
                    name: label,
                    value: value
                };

                // 应用数据点样式（包括爆炸效果和渐变填充）
                if (chartStyle.dataPointStyles && chartStyle.dataPointStyles.length > 0) {
                    const dpStyle = chartStyle.dataPointStyles[0][index];
                    if (dpStyle) {
                        // 爆炸效果
                        if (dpStyle.explosion !== undefined && dpStyle.explosion > 0) {
                            dataItem.selected = true;
                            dataItem.selectedOffset = dpStyle.explosion;
                        }
                        
                        // 渐变填充
                        if (dpStyle.gradientFill && dpStyle.gradientFill.color) {
                            dataItem.itemStyle = {
                                color: '#' + dpStyle.gradientFill.color[0]
                            };
                        }
                    }
                }

                data.push(dataItem);
            });
        }

        // 标准 2D 饼图
        const pieConfig = {
            name: series.key || 'Series 1',
            type: 'pie',
            radius: '50%',
            data: data,
            label: {
                show: true,
                formatter: '{b}: {d}%'
            },
            emphasis: {
                itemStyle: {
                    shadowBlur: 10,
                    shadowOffsetX: 0,
                    shadowColor: 'rgba(0, 0, 0, 0.5)'
                }
            }
        };

        // 3D 效果（通过阴影和多层饼图模拟）
        if (is3D) {
            const depthPercent = chartStyle.view3D?.depthPercent || 100;
            const depth = depthPercent / 100;
            
            // 为 3D 饼图存储数据项名称，用于图例显示
            const dataNames = data.map(item => item.name);
            
            // 创建多层饼图模拟 3D 效果
            const series3D = [];
            
            // 第一层：顶部亮色层（这个系列的名称用于图例）
            series3D.push({
                name: series.key || 'Series 1',
                type: 'pie',
                radius: ['30%', '55%'],
                data: data.map(item => ({
                    ...item,
                    itemStyle: {
                        color: item.itemStyle?.color || undefined,
                        borderWidth: 0,
                        opacity: 1
                    }
                })),
                label: {
                    show: true,
                    formatter: '{b}: {d}%'
                },
                z: 2,
                itemStyle: {
                    shadowBlur: 10,
                    shadowColor: 'rgba(0, 0, 0, 0.3)',
                    shadowOffsetY: -2
                }
            });

            // 第二层：中间渐变层
            series3D.push({
                name: series.key || 'Series 1',
                type: 'pie',
                radius: ['30%', '55%'],
                data: data.map(item => ({
                    name: item.name,
                    value: item.value,
                    itemStyle: {
                        color: this.darkenColor(item.itemStyle?.color || '#5470c6', 0.1),
                        borderWidth: 0,
                        opacity: 0.7
                    }
                })),
                label: { show: false },
                z: 1,
                center: ['50%', '50%'],
                legendHoverLink: false,
                emphasis: { disabled: true },
                tooltip: { show: false }
            });

            // 第三层：底部阴影层
            series3D.push({
                name: series.key || 'Series 1',
                type: 'pie',
                radius: ['30%', '55%'],
                data: data.map(item => ({
                    name: item.name,
                    value: item.value,
                    itemStyle: {
                        color: this.darkenColor(item.itemStyle?.color || '#5470c6', 0.3),
                        borderWidth: 0,
                        opacity: 0.4
                    }
                })),
                label: { show: false },
                z: 0,
                center: ['50%', '50%'],
                legendHoverLink: false,
                emphasis: { disabled: true },
                tooltip: { show: false }
            });

            // 如果有全局颜色设置，应用到所有层
            if (chartStyle.fillColor) {
                series3D.forEach(s => {
                    s.data.forEach(item => {
                        if (!item.itemStyle) item.itemStyle = {};
                        item.itemStyle.color = chartStyle.fillColor;
                    });
                });
            }

            // 将数据名称存储在图表样式中，供图例配置使用
            chartStyle._pieDataNames = dataNames;

            return series3D;
        }

        // 如果有全局颜色设置，应用到饼图
        if (chartStyle.fillColor) {
            pieConfig.itemStyle = pieConfig.itemStyle || {};
            pieConfig.itemStyle.color = chartStyle.fillColor;
        }

        // varyColors 控制：如果为 false，使用统一颜色
        if (chartStyle.varyColors === false && !pieConfig.itemStyle?.color) {
            pieConfig.itemStyle = pieConfig.itemStyle || {};
            pieConfig.itemStyle.color = '#5470c6';
        }

        return [pieConfig];
    }

    /**
     * 将颜色变暗，用于模拟 3D 阴影
     * @param {string} color - 颜色值
     * @param {number} amount - 变暗程度 (0-1)
     * @returns {string} 变暗后的颜色
     */
    darkenColor(color, amount) {
        // 如果没有颜色或不是 hex 格式，返回默认值
        if (!color || !color.startsWith('#')) return '#5470c6';
        
        let hex = color.replace('#', '');
        if (hex.length === 3) {
            hex = hex.split('').map(c => c + c).join('');
        }
        
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        
        const newR = Math.max(0, Math.floor(r * (1 - amount)));
        const newG = Math.max(0, Math.floor(g * (1 - amount)));
        const newB = Math.max(0, Math.floor(b * (1 - amount)));
        
        return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
    }

    /**
     * 准备散点图系列
     * @param {Array} chartData - 图表数据
     * @param {Object} chartInfo - 图表信息
     * @returns {Array} 散点图系列配置
     */
    prepareScatterSeries(chartData, chartInfo) {
        return chartData.map((series, index) => {
            const seriesConfig = {
                name: series.key || `Series ${index}`,
                type: 'scatter',
                data: series.values || []
            };

            // 应用系列颜色
            if (series.style) {
                if (series.style.fillColor) {
                    seriesConfig.itemStyle = {
                        color: series.style.fillColor
                    };
                } else if (series.style.gradientFill?.color?.length > 0) {
                    seriesConfig.itemStyle = {
                        color: '#' + series.style.gradientFill.color[0]
                    };
                }
            }

            return seriesConfig;
        });
    }

    /**
     * 获取图例配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} 图例配置
     */
    getLegendConfig(chartInfo) {
        const legendPosition = this.getLegendPosition(chartInfo);
        const legendStyle = chartInfo.style?.legend || {};
        const chartData = chartInfo.data;
        const isPieChart = chartInfo.type === 'pieChart' || chartInfo.type === 'pie3DChart';
        const is3DPie = chartInfo.type === 'pie3DChart';

        // 提取图例数据
        const legendData = [];
        if (isPieChart && Array.isArray(chartData) && chartData.length > 0) {
            // 饼图：使用数据项名称（xlabels）作为图例
            const series = chartData[0];

            // 对于 3D 饼图，不显式设置 legend.data，让 ECharts 自动从数据中提取
            // 因为 3D 饼图有多层系列，显式设置会导致警告
            if (!is3DPie) {
                // 对于 2D 饼图，如果已经存储了数据名称，使用它
                if (chartInfo.style?._pieDataNames) {
                    legendData.push(...chartInfo.style._pieDataNames);
                } else if (series.xlabels && Array.isArray(series.xlabels)) {
                    legendData.push(...series.xlabels);
                } else if (series.values && Array.isArray(series.values)) {
                    // 如果没有 xlabels，使用值的索引
                    series.values.forEach((v, i) => {
                        legendData.push(`Item ${i + 1}`);
                    });
                }
            }
        } else if (Array.isArray(chartData)) {
            // 其他图表：使用系列名称
            chartData.forEach((series, index) => {
                legendData.push(series.key || `Series ${index + 1}`);
            });
        }

        const legendConfig = {};
        // 只有在有数据的情况下才设置 legend.data
        if (legendData.length > 0) {
            legendConfig.data = legendData;
        }

        // 设置图例位置
        switch (legendPosition) {
            case 'top':
                legendConfig.top = 0;
                legendConfig.left = 'center';
                legendConfig.orient = 'horizontal';
                break;
            case 'bottom':
                legendConfig.bottom = 0;
                legendConfig.left = 'center';
                legendConfig.orient = 'horizontal';
                break;
            case 'left':
                legendConfig.left = 0;
                legendConfig.top = 'middle';
                legendConfig.orient = 'vertical';
                break;
            case 'right':
            default:
                legendConfig.right = 0;
                legendConfig.top = 'middle';
                legendConfig.orient = 'vertical';
                break;
        }

        // 应用图例样式
        if (legendStyle.color) {
            legendConfig.textStyle = {
                color: legendStyle.color
            };
        }

        if (legendStyle.fontSize) {
            legendConfig.textStyle = legendConfig.textStyle || {};
            legendConfig.textStyle.fontSize = legendStyle.fontSize;
        }

        return legendConfig;
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
     * 获取 X 轴数据
     * @param {Object} chartInfo - 图表信息
     * @returns {Array} X 轴数据
     */
    getXAxisData(chartInfo) {
        const chartData = chartInfo.data;

        if (!Array.isArray(chartData) || chartData.length === 0) {
            return [];
        }

        // 使用第一个系列的 xlabels 作为 X 轴数据
        const firstSeries = chartData[0];
        if (firstSeries.xlabels && Array.isArray(firstSeries.xlabels)) {
            return firstSeries.xlabels;
        }

        // 如果没有 xlabels，使用值的索引
        if (firstSeries.values && Array.isArray(firstSeries.values)) {
            return firstSeries.values.map((v, i) => i.toString());
        }

        return [];
    }

    /**
     * 获取 X 轴配置
     * @param {Object} chartInfo - 图表信息
     * @param {Array} xAxisData - X 轴数据
     * @returns {Object} X 轴配置
     */
    getXAxisConfig(chartInfo, xAxisData) {
        const xAxisStyle = chartInfo.style?.categoryAxis || {};

        const config = {
            type: 'category',
            data: xAxisData,
            axisLine: {
                show: true,
                lineStyle: {}
            },
            axisLabel: {},
            axisTick: {
                show: true
            }
        };

        // 轴标签颜色和字体大小
        if (xAxisStyle.color) {
            config.axisLabel.color = xAxisStyle.color;
        }
        if (xAxisStyle.fontSize) {
            config.axisLabel.fontSize = xAxisStyle.fontSize;
        }

        // 轴线颜色
        if (xAxisStyle.lineColor) {
            config.axisLine.lineStyle.color = xAxisStyle.lineColor;
        }

        // 网格线配置（如果有的话）
        if (xAxisStyle.gridlineColor) {
            config.splitLine = {
                show: true,
                lineStyle: {
                    color: xAxisStyle.gridlineColor
                }
            };
        } else {
            config.splitLine = {
                show: false
            };
        }

        return config;
    }

    /**
     * 获取 Y 轴配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} Y 轴配置
     */
    getYAxisConfig(chartInfo) {
        const yAxisStyle = chartInfo.style?.valueAxis || {};

        const config = {
            type: 'value',
            axisLine: {
                show: true,
                lineStyle: {}
            },
            axisLabel: {},
            axisTick: {
                show: true
            },
            splitLine: {
                show: true,
                lineStyle: {}
            }
        };

        // 轴标签颜色和字体大小
        if (yAxisStyle.color) {
            config.axisLabel.color = yAxisStyle.color;
        }
        if (yAxisStyle.fontSize) {
            config.axisLabel.fontSize = yAxisStyle.fontSize;
        }

        // 轴线颜色
        if (yAxisStyle.lineColor) {
            config.axisLine.lineStyle.color = yAxisStyle.lineColor;
        }

        // 网格线颜色
        if (yAxisStyle.gridlineColor) {
            config.splitLine.lineStyle.color = yAxisStyle.gridlineColor;
        }
        if (yAxisStyle.gridlineWidth) {
            config.splitLine.lineStyle.width = yAxisStyle.gridlineWidth;
        }

        return config;
    }

    /**
     * 更新图表数据
     * @param {string} chartId - 图表 ID
     * @param {Object} newChartInfo - 新的图表信息
     */
    updateChart(chartId, newChartInfo) {
        const chart = this.chartInstances.get(chartId);
        if (!chart) {
            console.warn(`Chart ${chartId} not found`);
            return;
        }

        const option = this.prepareEChartsOption(newChartInfo);
        if (option) {
            chart.setOption(option);
        }
    }

    /**
     * 销毁图表
     * @param {string} chartId - 图表 ID
     */
    destroyChart(chartId) {
        const chart = this.chartInstances.get(chartId);
        if (chart) {
            chart.dispose();
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
