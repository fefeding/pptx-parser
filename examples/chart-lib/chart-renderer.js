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
                // 把容器作用域传递下去，否则同一份幻灯片 HTML 被渲染多份（缩略图/画布/演示）时会互相串台
                this.renderChart(chartInfo, container);
            } catch (err) {
                console.error(`Error rendering chart ${chartInfo.chartId}:`, err);
            }
        });
    }

    /**
     * 在指定容器内按 ID 查找图表容器
     * @param {string} chartId - 图表容器 ID
     * @param {HTMLElement|null} container - 作用域容器，为 null 时回退到 document
     * @returns {HTMLElement|null}
     */
    findChartElement(chartId, container = null) {
        if (!container) {
            return document.getElementById(chartId);
        }
        if (container.id === chartId) {
            return container;
        }
        return container.querySelector(`[id="${chartId}"]`);
    }

    /**
     * 渲染单个图表
     * @param {Object} chartInfo - 图表信息对象
     * @param {string} chartInfo.chartId - 图表容器 ID
     * @param {string} chartInfo.type - 图表类型
     * @param {Array} chartInfo.data - 图表数据
     * @param {Object} chartInfo.style - 图表样式
     */
    renderChart(chartInfo, container = null) {
        const chartElement = this.findChartElement(chartInfo.chartId, container);
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

        // 容器可能尚未完成布局（clientWidth/clientHeight 为 0），
        // 此时 echarts.init 会创建 0x0 画布并告警。
        // 延迟到容器有可测量尺寸时再初始化；若仍无尺寸则兜底使用默认尺寸，保证图表可见。
        const doInit = () => {
            const w = chartElement.clientWidth;
            const h = chartElement.clientHeight;
            const hasSize = w > 0 && h > 0;
            const chart = hasSize
                ? echarts.init(chartElement)
                : echarts.init(chartElement, null, {
                    width: Math.max(w, 600),
                    height: Math.max(h, 400)
                });

            // 设置配置
            chart.setOption(option);

            // 启用窗口大小调整
            window.addEventListener('resize', () => {
                chart.resize();
            });

            // 容器后续获得真实尺寸时重新适配
            if (!hasSize) {
                requestAnimationFrame(() => {
                    if (chartElement.clientWidth > 0 && chartElement.clientHeight > 0) {
                        chart.resize();
                    }
                });
            }

            // 存储图表实例以便后续更新
            this.chartInstances.set(chartInfo.chartId, chart);
        };

        if (chartElement.clientWidth > 0 && chartElement.clientHeight > 0) {
            doInit();
        } else {
            // 等待下一帧布局完成后再初始化
            requestAnimationFrame(doInit);
        }
    }

    /**
     * 准备 ECharts 配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} ECharts 配置对象
     */
    prepareEChartsOption(chartInfo) {
        const chartData = chartInfo.data;
        const pptxType = chartInfo.type;

        // 真 3D：echarts-gl 可用时优先用 3D 坐标系渲染，否则降级为同族 2D
        if (this.isGL3DAvailable()) {
            const glType = this.map3DChartType(pptxType, chartData);
            if (glType) {
                const opt3d = this.prepare3DOption(chartInfo, glType);
                if (opt3d) return opt3d;
            }
        }

        const chartType = this.mapChartType(pptxType);

        if (!chartType) {
            console.warn(`Unsupported chart type: ${pptxType}`);
            return null;
        }

        // 雷达图：需要独立的 radar + series 结构，没有笛卡尔坐标轴
        if (chartType === 'radar') {
            const radar = this.prepareRadarSeries(chartData, chartInfo);
            const option = {
                tooltip: { trigger: 'item' },
                legend: this.getLegendConfig(chartInfo),
                radar: radar.radar,
                series: radar.series
            };
            this.applyChartBackground(option, chartInfo.style);
            if (chartInfo.title || chartInfo.style?.title) {
                option.title = this.getTitleConfig(chartInfo.title, chartInfo.style?.title);
            }
            return option;
        }

        const isPieChart = chartType === 'pie';
        const is3DPie = pptxType === 'pie3DChart';
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

            // 散点/气泡使用数值 X 轴；股票图无类别时仍用类目轴（以索引为 X），
            // 避免 [open,close,low,high] 在数值轴上被误读为 [x,open,close,low]。
            if (chartType === 'scatter') {
                option.xAxis.type = 'value';
                option.xAxis.data = undefined;
                option.xAxis.boundaryGap = undefined;
            }

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
     * 是否具备真 3D 渲染能力（echarts-gl 已加载并完成 3D 坐标系/系列注册）
     * @returns {boolean}
     */
    isGL3DAvailable() {
        try {
            if (typeof window === 'undefined' || !window.echarts) return false;
            // echarts-gl 的 UMD 构建挂载 window['echarts-gl']，加载时即完成注册
            return !!(window['echarts-gl'] || window.echartsGL);
        } catch (e) {
            return false;
        }
    }

    /**
     * PPTX 3D 图表类型 → echarts-gl 系列类型
     * @param {string} pptxType - PPTX 图表类型
     * @param {Array} chartData - 图表数据（surface 需据规模判断能否构成网格）
     * @returns {string|null} echarts-gl 系列类型；非 3D 类型返回 null
     */
    map3DChartType(pptxType, chartData) {
        if (pptxType === 'bar3DChart') return 'bar3D';
        if (pptxType === 'line3DChart') return 'line3D';
        // echarts-gl 无 area3D，按 3D 折线呈现
        if (pptxType === 'area3DChart') return 'line3D';
        if (pptxType === 'surfaceChart' || pptxType === 'surface3DChart') {
            // surface 需要 M×N 规则点阵才能三角化成网格；单系列退化为 3D 折线
            const isGrid = Array.isArray(chartData) && chartData.length >= 2
                && chartData.every(s => Array.isArray(s.values) && s.values.length >= 2);
            return isGrid ? 'surface' : 'line3D';
        }
        return null;
    }

    /**
     * 从首个系列的 xlabels 取 3D 类别轴（按 idx 升序）
     * @param {Array} chartData - 图表数据
     * @returns {Array<string>} 类别名列表
     */
    get3DCategories(chartData) {
        const first = chartData.find(s => s && s.xlabels);
        if (!first) return [];
        return Object.keys(first.xlabels)
            .sort((a, b) => Number(a) - Number(b))
            .map(k => first.xlabels[k]);
    }

    /**
     * 准备 echarts-gl 真 3D 配置
     * @param {Object} chartInfo - 图表信息
     * @param {string} glType - echarts-gl 系列类型
     * @returns {Object|null} 配置对象；数据不足以构成 3D 时返回 null（调用方降级为 2D）
     */
    prepare3DOption(chartInfo, glType) {
        const chartData = chartInfo.data;
        if (!Array.isArray(chartData) || chartData.length === 0) return null;

        const categories = this.get3DCategories(chartData);
        if (!categories.length) return null;

        const style = chartInfo.style || {};
        const view3D = style.view3D || {};
        const seriesNames = chartData.map((s, i) => s.key || `Series ${i + 1}`);

        // 系列取色顺序与 2D 路径保持一致
        const colorOf = (series) => {
            const st = series && series.style;
            if (!st) return undefined;
            const c = st.fillColor
                || (st.gradientFill && st.gradientFill.color && st.gradientFill.color[0]);
            return c ? (String(c).startsWith('#') ? c : '#' + c) : undefined;
        };

        // 数据点：[x=类别下标, y=系列下标, z=数值]
        const pointAt = (ci, si) => {
            const vals = (chartData[si] && chartData[si].values) || [];
            const v = vals[ci];
            const z = v && v.y !== undefined ? parseFloat(v.y) : 0;
            return isNaN(z) ? 0 : z;
        };

        let series;
        let yAxis3D;
        if (glType === 'line3D') {
            // 每个 PPTX 系列一条 3D 折线，沿 y 轴错开，可用图例区分
            yAxis3D = { type: 'value', name: '' };
            series = chartData.map((s, si) => {
                const cfg = {
                    name: seriesNames[si],
                    type: 'line3D',
                    lineStyle: { width: 3 },
                    data: Array.from({ length: (s.values || []).length },
                        (_, ci) => [ci, si, pointAt(ci, si)])
                };
                const c = colorOf(s);
                if (c) cfg.lineStyle.color = c;
                return cfg;
            });
        } else if (glType === 'surface') {
            yAxis3D = { type: 'category', data: seriesNames, name: '' };
            const data = [];
            chartData.forEach((s, si) => {
                (s.values || []).forEach((_, ci) => data.push([ci, si, pointAt(ci, si)]));
            });
            series = [{
                name: seriesNames[0],
                type: 'surface',
                wireframe: { show: style.wireframe === true },
                shading: 'lambert',
                data
            }];
        } else {
            // bar3D：单一系列承载全部柱体，按所属系列着色
            yAxis3D = { type: 'category', data: seriesNames, name: '' };
            const data = [];
            chartData.forEach((s, si) => {
                (s.values || []).forEach((_, ci) => {
                    const item = { value: [ci, si, pointAt(ci, si)] };
                    const c = colorOf(s);
                    if (c) item.itemStyle = { color: c };
                    data.push(item);
                });
            });
            series = [{
                name: seriesNames[0],
                type: 'bar3D',
                shading: 'lambert',
                bevelSize: 0.3,
                bevelSmoothness: 2,
                data
            }];
        }

        // OOXML view3D：rotX(0-90) 俯仰 / rotY(0-360) 旋转 / depthPercent 厚度 / rAngAx 直角轴
        const rotX = view3D.rotX !== undefined ? Number(view3D.rotX) : 30;
        const rotY = view3D.rotY !== undefined ? Number(view3D.rotY) : 20;
        const depthPercent = view3D.depthPercent !== undefined ? Number(view3D.depthPercent) : 100;

        const option = {
            tooltip: {},
            xAxis3D: { type: 'category', data: categories, name: '' },
            yAxis3D,
            zAxis3D: { type: 'value', name: '' },
            grid3D: {
                boxWidth: 120,
                boxHeight: 70,
                boxDepth: Math.max(20, Math.min(200, 80 * (depthPercent / 100))),
                viewControl: {
                    alpha: Math.max(0, Math.min(90, rotX)),
                    beta: ((rotY % 360) + 360) % 360,
                    // rAngAx=1 为直角轴（正交投影），否则透视投影
                    projection: view3D.rAngAx === true ? 'orthographic' : 'perspective',
                    autoRotate: false
                },
                light: {
                    main: { intensity: 1.2, shadow: true },
                    ambient: { intensity: 0.3 }
                }
            },
            series
            // 说明：3D 系列由单一 series 承载多系列数据，图例无法逐项开关，故不输出 legend
        };

        this.applyChartBackground(option, style);

        if (chartInfo.title || style.title) {
            option.title = this.getTitleConfig(chartInfo.title, style.title);
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
            'bar3DChart': 'bar',
            'pieChart': 'pie',
            'pie3DChart': 'pie',
            'doughnutChart': 'pie',
            'areaChart': 'line',
            'scatterChart': 'scatter',
            'bubbleChart': 'scatter',
            'radarChart': 'radar',
            'stockChart': 'candlestick',
            'surfaceChart': 'line',
            // 3D 变体在 ECharts 无原生对应，降级为同族 2D 类型渲染
            'line3DChart': 'line',
            'area3DChart': 'line',
            'surface3DChart': 'line',
            // 子母饼图按饼图渲染
            'ofPieChart': 'pie'
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
        const isAreaChart = chartInfo.type === 'areaChart' || chartInfo.type === 'area3DChart';
        const chartStyle = chartInfo.style || {};
        // 堆叠 / 百分比堆叠：同一 stack 名使系列堆叠
        const stackName = (chartStyle.grouping === 'stacked' || chartStyle.grouping === 'percentStacked')
            ? 'total' : undefined;

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

        // 股票图（蜡烛图）
        if (echartsType === 'candlestick') {
            return this.prepareStockSeries(chartData, chartInfo);
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
                name: series.key || `Series ${index + 1}`,
                type: echartsType,
                data: seriesData,
                smooth: isAreaChart || chartStyle.smooth === true,
                areaStyle: isAreaChart ? {} : undefined
            };

            // 堆叠（stacked / percentStacked）
            if (stackName) {
                seriesConfig.stack = stackName;
            }

            // 数据标记：显式关闭时隐藏符号
            if (chartStyle.marker === false) {
                seriesConfig.showSymbol = false;
                seriesConfig.symbol = 'none';
            }

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

            // 数据标签：仅当 info.dataLabels 为真时显示
            if (chartInfo.dataLabels) {
                seriesConfig.label = {
                    show: true,
                    position: 'top',
                    formatter: (p) => {
                        const v = Array.isArray(p.value) ? p.value[1] : (p.value && typeof p.value === 'object' ? p.value.y : p.value);
                        return v == null ? '' : String(v);
                    }
                };
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
        const isDoughnut = chartInfo.type === 'doughnutChart';
        const data = [];

        if (Array.isArray(series.values)) {
            series.values.forEach((item, index) => {
                let label = `Item ${index + 1}`;
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

        // 甜甜圈内径取 c:holeSize（百分比，默认 50）；外径固定 68%
        let holePercent = 50;
        if (chartStyle.holeSize !== undefined) {
            holePercent = Number(chartStyle.holeSize);
        }
        if (isNaN(holePercent)) {
            holePercent = 50;
        }
        holePercent = Math.max(1, Math.min(90, holePercent));

        // 标准 2D 饼图 / 甜甜圈图
        const pieConfig = {
            name: series.key || 'Series 1',
            type: 'pie',
            radius: isDoughnut ? [`${holePercent * 0.76}%`, '68%'] : '50%',
            data: data,
            label: {
                show: !!chartInfo.dataLabels,
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
                    show: !!chartInfo.dataLabels,
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
     *
     * chartData 可能形态：
     *  A) [[xVals], [yVals]] —— parser 对散点(xVal/yVal)的输出
     *  B) [{ key, values: [{x, y}], style }] —— 与其它图表一致的统一形态
     * 统一转换为 ECharts 散点所需的 [x, y] 点数组。
     */
    prepareScatterSeries(chartData, chartInfo) {
        // 形态 A：[[xVals], [yVals]]
        const isXYArrays = Array.isArray(chartData) && chartData.length >= 1 &&
            chartData.every(a => Array.isArray(a) && a.every(v => typeof v === 'number'));
        if (isXYArrays) {
            const xs = chartData[0] || [];
            const ys = chartData[1] || [];
            const points = xs.map((xv, i) => [xv, ys[i]]);
            const config = {
                name: 'Series 1',
                type: 'scatter',
                data: points
            };
            return [config];
        }

        // 形态 B：多系列统一形态（含气泡图的 size）
        return chartData.map((series, index) => {
            let data = [];
            const vals = series.values;
            if (Array.isArray(vals)) {
                const hasSize = vals.some(v => v && typeof v === 'object' && typeof v.size === 'number' && isFinite(v.size));
                data = vals.map(v => {
                    if (v && typeof v === 'object' && 'x' in v && 'y' in v) {
                        if (hasSize) {
                            return {
                                value: [Number(v.x), Number(v.y)],
                                symbolSize: Math.max(6, Number(v.size) * 4)
                            };
                        }
                        return [Number(v.x), Number(v.y)];
                    }
                    return Array.isArray(v) ? v : v;
                });
            }

            const seriesConfig = {
                name: series.key || `Series ${index + 1}`,
                type: 'scatter',
                data: data
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
     * 准备雷达图系列
     * @param {Array} chartData - 图表数据（与柱状/折线一致的统一形态）
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} { radar, series }
     */
    prepareRadarSeries(chartData, chartInfo) {
        // indicator 为空数组会让 echarts 在初始化雷达坐标系时抛错，返回最小可用结构
        const emptyRadar = () => ({ radar: { indicator: [{ name: '', max: 1 }] }, series: [] });

        if (!Array.isArray(chartData) || chartData.length === 0) {
            return emptyRadar();
        }
        const first = chartData[0];
        const cats = first.xlabels || {};

        let maxVal = 0;
        chartData.forEach(s => (s.values || []).forEach(v => {
            const y = v && v.y != null ? Number(v.y) : 0;
            if (y > maxVal) maxVal = y;
        }));
        if (maxVal <= 0) maxVal = 1;

        // indicator 必须与数据点数一一对应且非空。
        // OOXML 雷达图常省略 c:cat（即无类别名），此时按数据点下标生成占位名。
        const pointCount = Math.max(
            0,
            ...chartData.map(s => (s.values || []).length),
            Object.keys(cats).length
        );
        if (pointCount === 0) {
            return emptyRadar();
        }
        const names = Array.from({ length: pointCount }, (_, i) => {
            const label = cats[i] !== undefined ? cats[i] : cats[String(i)];
            return label !== undefined && label !== '' ? String(label) : `指标 ${i + 1}`;
        });

        const indicator = names.map(name => ({ name, max: maxVal * 1.1 }));
        const series = chartData.map((s, i) => ({
            name: s.key || `Series ${i + 1}`,
            type: 'radar',
            data: [{
                name: s.key || `Series ${i + 1}`,
                value: (s.values || []).map(v => v && v.y != null ? Number(v.y) : 0)
            }]
        }));

        return { radar: { indicator }, series };
    }

    /**
     * 准备股票图（蜡烛图）系列
     * @param {Array} chartData - 图表数据（values 为 [open,close,low,high] 数组）
     * @param {Object} chartInfo - 图表信息
     * @returns {Array} 系列配置数组
     */
    prepareStockSeries(chartData, chartInfo) {
        if (!Array.isArray(chartData)) {
            console.error('Chart data is not an array');
            return [];
        }

        return chartData.map((series, index) => ({
            name: series.key || `Series ${index + 1}`,
            type: 'candlestick',
            data: Array.isArray(series.values) ? series.values : []
        }));
    }

    /**
     * 获取图例配置
     * @param {Object} chartInfo - 图表信息
     * @returns {Object} 图例配置
     */
    getLegendConfig(chartInfo) {
        const legendPosition = this.getLegendPosition(chartInfo);
        const legendStyle = chartInfo.style?.legend || {};

        const legendConfig = {};

        // 不手动设置 legend.data：交由 ECharts 依据 series 名称 / 饼图数据名
        // 自动生成图例，避免「Legend data should be same with series name」不匹配告警。

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
