/**
 * Chart processing module
 * Handles chart generation and data processing
 */

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * Generate chart HTML and data
 * @param {Object} node - Chart node
 * @param {Object} warpObj - Warp object containing context
 * @param {Object} parentNode - Parent node (for group elements coordinate calculation)
 * @returns {Promise<string>} Chart HTML
 */
async function genChart(node, warpObj, parentNode) {
    const order = node["attrs"]["order"];
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);

    // 处理组合缩放 - 当chart在group-abs类型组合中时需要应用缩放
    let workingXfrmNode = xfrmNode;
    if (warpObj.currentGroupScale && xfrmNode) {
        const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

        // 创建缩放后的xfrmNode
        workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));

        // 缩放尺寸
        if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
            const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
            const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
            workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
            workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
        }

        // 调整位置(相对于childX/childY)
        if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
            const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
            const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

            // 计算相对于childOff的偏移
            const relativeX = originalOffX - (childX / SLIDE_FACTOR);
            const relativeY = originalOffY - (childY / SLIDE_FACTOR);

            // 应用缩放
            workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR + relativeX * scaleX);
            workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR + relativeY * scaleY);
        }
    }

    // 提取位置和尺寸信息
    let offX = 0, offY = 0, extCx = 0, extCy = 0;
    if (workingXfrmNode !== undefined) {
        if (workingXfrmNode['a:off'] && workingXfrmNode['a:off'].attrs) {
            offX = workingXfrmNode['a:off'].attrs.x || 0;
            offY = workingXfrmNode['a:off'].attrs.y || 0;
        }
        if (workingXfrmNode['a:ext'] && workingXfrmNode['a:ext'].attrs) {
            extCx = workingXfrmNode['a:ext'].attrs.cx || 0;
            extCy = workingXfrmNode['a:ext'].attrs.cy || 0;
        }
    }

    // 生成 data- 属性
    const dataAttrs = ` data-node-type="chart" data-off-x="${offX}" data-off-y="${offY}" data-ext-cx="${extCx}" data-ext-cy="${extCy}"`;

    const result = "<div id='chart" + warpObj.chartId.value + "' class='block content' style='" +
        PPTXXmlUtils.getPosition(workingXfrmNode, parentNode || node, undefined, undefined) + PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) +
        ` z-index: ${order};'${dataAttrs}></div>`;

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
