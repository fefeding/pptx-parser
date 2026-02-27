/**
 * Chart processing module
 * Handles chart generation and data processing
 */

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';

/**
 * Generate chart HTML and data
 * @param {Object} node - Chart node
 * @param {Object} warpObj - Warp object containing context
 * @returns {Promise<string>} Chart HTML
 */
async function genChart(node, warpObj) {
    const order = node["attrs"]["order"];
    const xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
    const result = "<div id='chart" + warpObj.chartId.value + "' class='block content' style='" +
        PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) + PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
        ` z-index: ${order};'></div>`;

    const rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    const refName = warpObj["slideResObj"][rid]["target"];
    const content = await PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
    const chartSpace = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace"]);
    const chart = PPTXXmlUtils.getTextByPathList(chartSpace, ["c:chart"]);
    const plotArea = PPTXXmlUtils.getTextByPathList(chart, ["c:plotArea"]);

    // 提取3D视图属性
    const view3D = PPTXXmlUtils.getTextByPathList(chart, ["c:view3D"]);
    const view3DProps = {};
    if (view3D) {
        if (view3D["attrs"]?.rotX !== undefined) view3DProps.rotX = parseFloat(view3D["attrs"].rotX);
        if (view3D["attrs"]?.rotY !== undefined) view3DProps.rotY = parseFloat(view3D["attrs"].rotY);
        if (view3D["attrs"]?.depthPercent !== undefined) view3DProps.depthPercent = parseFloat(view3D["attrs"].depthPercent);
        if (view3D["attrs"]?.rAngAx !== undefined) view3DProps.rAngAx = view3D["attrs"].rAngAx === "1";
    }

    // 提取图表类型特定属性
    const chartType = Object.keys(plotArea).find(key => key.startsWith('c:') && key.endsWith('Chart'));
    const varyColors = chartType ? PPTXXmlUtils.getTextByPathList(plotArea[chartType], ["c:varyColors", "attrs", "val"]) : undefined;

    // 提取系列数据点的样式（dPt）和爆炸效果（explosion）
    let dataPointStyles = [];
    if (chartType && plotArea[chartType]["c:ser"]) {
        const serArray = Array.isArray(plotArea[chartType]["c:ser"]) 
            ? plotArea[chartType]["c:ser"] 
            : [plotArea[chartType]["c:ser"]];
        
        serArray.forEach(ser => {
            const dPtArray = ser["c:dPt"];
            if (dPtArray) {
                const dpStyles = {};
                const dpList = Array.isArray(dPtArray) ? dPtArray : [dPtArray];
                dpList.forEach(dp => {
                    const idx = dp["c:idx"]?.["attrs"]?.val;
                    const explosion = dp["c:explosion"]?.["attrs"]?.val;
                    const spPr = dp["c:spPr"];
                    
                    if (idx !== undefined) {
                        const dpStyle = {};
                        if (explosion !== undefined) {
                            dpStyle.explosion = parseFloat(explosion);
                        }
                        if (spPr) {
                            const gradFill = spPr["a:gradFill"];
                            if (gradFill) {
                                dpStyle.gradientFill = PPTXStyleUtils.getGradientFill(gradFill, warpObj);
                            }
                        }
                        dpStyles[idx] = dpStyle;
                    }
                });
                dataPointStyles.push(dpStyles);
            }
        });
    }

    // 提取图表标题
    const chartTitleObj = PPTXStyleUtils.extractChartTitleStyle(chart, warpObj);
    const chartTitle = chartTitleObj.text;

    // 提取图表样式信息
    const chartStyle = {
        chartArea: PPTXStyleUtils.extractChartAreaStyle(chartSpace, warpObj),
        legend: PPTXStyleUtils.extractChartLegendStyle(chart, warpObj),
        categoryAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:catAx", warpObj),
        valueAxis: PPTXStyleUtils.extractChartAxisStyle(plotArea, "c:valAx", warpObj),
        view3D: view3DProps,
        varyColors: varyColors === "1",
        dataPointStyles: dataPointStyles,
        title: chartTitleObj.style
    };

    let chartData = null;
    for (const key in plotArea) {
        switch (key) {
            case "c:lineChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "lineChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:barChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "barChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pieChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "pieChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:pie3DChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "pie3DChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:areaChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "areaChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:scatterChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "scatterChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:catAx":
                break;
            case "c:valAx":
                break;
            default:
        }
    }

    return result;
}

/**
 * Process message queue for charts
 * @param {Array} queue - Message queue
 * @param {Object} result - Result object to store chart data
 */
function processMsgQueue(queue, result) {
    for (const msg of queue) {
        if (msg.type === "chart" || msg.type === "createChart") {
            const chartObj = msg.data;
            result.charts.push({
                chartId: chartObj.chartId,
                type: chartObj.chartType,
                data: chartObj.chartData,
                style: chartObj.style,
                title: chartObj.title
            });
        }
    }
}

/**
 * Process single chart message
 * @param {Object} data - Chart data
 * @param {Object} callbacks - Callback functions
 */
function processSingleMsg(data, callbacks) {
    const { chartId, chartType, chartData } = data;
    let chartDataArray = [];
    let chart = null;

    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
        console.warn(`Invalid chart data for chart ID: ${chartId}`);
        return;
    }

    switch (chartType) {
        case "lineChart":
            chartDataArray = chartData;
            chart = nv.models.lineChart().useInteractiveGuideline(true);
            if (chartData[0]?.xlabels) {
                chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
            }
            break;

        case "barChart":
            chartDataArray = chartData;
            chart = nv.models.multiBarChart();
            if (chartData[0]?.xlabels) {
                chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
            }
            break;

        case "pieChart":
        case "pie3DChart":
            chartDataArray = chartData[0]?.values || [];
            chart = nv.models.pieChart();
            break;

        case "areaChart":
            chartDataArray = chartData;
            chart = nv.models.stackedAreaChart()
                .clipEdge(true)
                .useInteractiveGuideline(true);
            if (chartData[0]?.xlabels) {
                chart.xAxis.tickFormat(d => chartData[0].xlabels[d] || d);
            }
            break;

        case "scatterChart":
            for (let i = 0; i < chartData.length; i++) {
                const arr = [];
                if (Array.isArray(chartData[i])) {
                    for (let j = 0; j < chartData[i].length; j++) {
                        arr.push({ x: j, y: chartData[i][j] });
                    }
                }
                chartDataArray.push({ key: `data${i + 1}`, values: arr });
            }
            chart = nv.models.scatterChart()
                .showDistX(true)
                .showDistY(true)
                .color(d3.scale.category10().range());
            chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
            chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
            break;

        default:
            console.warn(`Unknown chart type: ${chartType}`);
    }

    if (chart !== null && callbacks.onChartReady) {
        callbacks.onChartReady({
            chartId,
            chart,
            data: chartDataArray
        });
    }
}

export {
    genChart,
    processMsgQueue,
    processSingleMsg
};
