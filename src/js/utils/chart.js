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
    const plotArea = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

    let chartData = null;
    for (const key in plotArea) {
        switch (key) {
            case "c:lineChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": "chart" + warpObj.chartId.value++,
                        "chartType": "lineChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
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
                data: chartObj.chartData
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
