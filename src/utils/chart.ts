/**
 * Chart processing module
 * Handles chart generation and data processing
 */

import { PPTXXmlUtils } from './xml';
import { PPTXStyleUtils } from './style';
import { SLIDE_FACTOR } from '../core/constants';
import type { XmlNode, WarpObject } from '../core/types';

// nvd3 / d3 由浏览器通过 <script> 标签注入，运行时作为全局变量存在
declare const nv: any;
declare const d3: any;

/**
 * Generate chart HTML and data
 * @param {Object} node - Chart node
 * @param {Object} warpObj - Warp object containing context
 * @param {Object} parentNode - Parent node (for group elements coordinate calculation)
 * @returns {Promise<string>} Chart HTML
 */
async function genChart(node: XmlNode | undefined, warpObj: WarpObject, parentNode: XmlNode | undefined) {
    const order = node!["attrs"]!["order"];
    // graphicFrame 的变换元素标准是 <a:xfrm>（ECMA-376），
    // 但部分生成工具（旧版/某些导出器）会写成 <p:xfrm>，需兼容两种写法
    let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["a:xfrm"]) ||
        PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);

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

    const result = `<div id='chart${warpObj.chartId.value}' class='block content' style='${PPTXXmlUtils.getPosition(workingXfrmNode, parentNode || node, undefined, undefined)}${PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined)}` +
        ` z-index: ${order};'${dataAttrs}></div>`;

    const rid = node!["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    const refName = warpObj["slideResObj"][rid]["target"];
    const content = await PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
    // Guard: chart XML file may be missing or unreadable
    if (!content) {
        return result;
    }
    const chartSpace = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace"]);
    if (!chartSpace) {
        return result;
    }
    const chart = PPTXXmlUtils.getTextByPathList(chartSpace, ["c:chart"]);
    const plotArea = PPTXXmlUtils.getTextByPathList(chart, ["c:plotArea"]);
    collectChartEntries(chartSpace, chart, plotArea, warpObj);

    return result;
}

/**
 * 将一个 c:chartSpace 的图表数据提取为 createChart 消息并压入消息队列。
 *
 * 抽离为独立函数，供 HTML 生成链路（genChart）与纯 JSON 解析链路共用，
 * 使 pptxToJson 无需经过 HTML 转换也能拿到图表数据。
 */
function collectChartEntries(chartSpace: any, chart: any, plotArea: any, warpObj: WarpObject) {
    if (!plotArea) return;

    // 提取3D视图属性
    const view3D = PPTXXmlUtils.getTextByPathList(chart, ["c:view3D"]);
    const view3DProps: Record<string, unknown> = {};
    if (view3D) {
        if (view3D["attrs"]?.rotX !== undefined) view3DProps.rotX = parseFloat(view3D["attrs"].rotX);
        if (view3D["attrs"]?.rotY !== undefined) view3DProps.rotY = parseFloat(view3D["attrs"].rotY);
        if (view3D["attrs"]?.depthPercent !== undefined) view3DProps.depthPercent = parseFloat(view3D["attrs"].depthPercent);
        if (view3D["attrs"]?.rAngAx !== undefined) view3DProps.rAngAx = view3D["attrs"].rAngAx === "1";
    }

    // 提取图表类型特定属性
    const chartType = Object.keys(plotArea).find(key => key.startsWith('c:') && key.endsWith('Chart'));
    // 组合图/多 plot 场景下同类型节点可能是数组，统一取首个
    const chartTypeNode = chartType
        ? (Array.isArray(plotArea[chartType]) ? plotArea[chartType][0] : plotArea[chartType])
        : undefined;
    const varyColors = chartTypeNode ? PPTXXmlUtils.getTextByPathList(chartTypeNode, ["c:varyColors", "attrs", "val"]) : undefined;
    // 分组（clustered / stacked / percentStacked / standard）与甜甜圈内径
    const grouping = chartTypeNode ? PPTXXmlUtils.getTextByPathList(chartTypeNode, ["c:grouping", "attrs", "val"]) : undefined;
    const holeSize = chartTypeNode ? PPTXXmlUtils.getTextByPathList(chartTypeNode, ["c:holeSize", "attrs", "val"]) : undefined;
    // 平滑线与数据标记取自首个系列（同类型节点可能有多条 ser）
    const firstSerNode = chartTypeNode
        ? (Array.isArray(chartTypeNode["c:ser"]) ? chartTypeNode["c:ser"][0] : chartTypeNode["c:ser"])
        : undefined;
    const smoothVal = firstSerNode ? PPTXXmlUtils.getTextByPathList(firstSerNode, ["c:smooth", "attrs", "val"]) : undefined;
    const markerSymbol = firstSerNode
        ? PPTXXmlUtils.getTextByPathList(firstSerNode, ["c:marker", "c:symbol", "attrs", "val"])
        : undefined;

    // 提取系列数据点的样式（dPt）和爆炸效果（explosion）
    let dataPointStyles: Array<Record<string, unknown>> = [];
    if (chartType && plotArea[chartType]["c:ser"]) {
        const serArray = Array.isArray(plotArea[chartType]["c:ser"]) 
            ? plotArea[chartType]["c:ser"] 
            : [plotArea[chartType]["c:ser"]];
        
        serArray.forEach(ser => {
            const dPtArray = ser["c:dPt"];
            if (dPtArray) {
                const dpStyles: Record<string, unknown> = {};
                const dpList = Array.isArray(dPtArray) ? dPtArray : [dPtArray];
                dpList.forEach(dp => {
                    const idx = dp["c:idx"]?.["attrs"]?.val;
                    const explosion = dp["c:explosion"]?.["attrs"]?.val;
                    const spPr = dp["c:spPr"];
                    
                    if (idx !== undefined) {
                        const dpStyle: Record<string, unknown> = {};
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
        grouping: grouping,
        holeSize: holeSize !== undefined && holeSize !== "" ? Number(holeSize) : undefined,
        smooth: smoothVal !== undefined ? smoothVal !== "0" : undefined,
        marker: markerSymbol !== undefined ? markerSymbol !== "none" : undefined,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
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
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "scatterChart",
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:bar3DChart":
            case "c:doughnutChart":
            case "c:radarChart":
            case "c:surfaceChart":
            case "c:surface3DChart":
            case "c:line3DChart":
            case "c:area3DChart":
            case "c:ofPieChart": {
                const t = key.replace(/^c:/, '');
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": t,
                        "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"], warpObj),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            }
            case "c:bubbleChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "bubbleChart",
                        "chartData": extractBubbleData(plotArea[key]["c:ser"]),
                        "style": chartStyle,
                        "title": chartTitle
                    }
                };
                warpObj.msgQueue.push(chartData);
                break;
            case "c:stockChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartId": `chart${warpObj.chartId.value++}`,
                        "chartType": "stockChart",
                        "chartData": extractStockData(plotArea[key]["c:ser"]),
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
}

/**
 * 独立扫描一页幻灯片内的图表部件，产出 createChart 消息（不依赖 HTML 生成流程）。
 *
 * 用于 pptxToJson：该链路不做 HTML 转换，若不在此处补扫描，
 * 结果中的 charts 将恒为空数组。
 */
async function extractChartsFromSlide(slideData: any, zip: any): Promise<void> {
    if (!slideData || !zip) return;
    const spTree = slideData.slideContent
        && slideData.slideContent["p:sld"]
        && slideData.slideContent["p:sld"]["p:cSld"]
        && slideData.slideContent["p:sld"]["p:cSld"]["p:spTree"];
    if (!spTree) return;

    const frames: any[] = [];
    collectGraphicFrames(spTree, frames);
    if (!frames.length) return;

    const warpObj = Object.assign({}, slideData, { zip }) as WarpObject;
    const resObj = slideData.slideResObj || {};

    for (const frame of frames) {
        const chartRef = frame && frame["a:graphic"]
            && frame["a:graphic"]["a:graphicData"]
            && frame["a:graphic"]["a:graphicData"]["c:chart"];
        const rid = chartRef && chartRef["attrs"] && chartRef["attrs"]["r:id"];
        if (!rid) continue;

        const target = resObj[rid] && resObj[rid].target;
        if (!target) continue;

        const content = await PPTXXmlUtils.readXmlFile(zip, target);
        const chartSpace = content && PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace"]);
        if (!chartSpace) continue;

        const chart = PPTXXmlUtils.getTextByPathList(chartSpace, ["c:chart"]);
        const plotArea = chart && PPTXXmlUtils.getTextByPathList(chart, ["c:plotArea"]);
        collectChartEntries(chartSpace, chart, plotArea, warpObj);
    }
}

/** 递归收集 spTree 下所有 p:graphicFrame 节点（含组合内的） */
function collectGraphicFrames(node: any, out: any[]): void {
    if (!node || typeof node !== "object") return;
    for (const key of Object.keys(node)) {
        if (key === "attrs" || key === "innerText") continue;
        const child = node[key];
        if (!child || typeof child !== "object") continue;
        const items = Array.isArray(child) ? child : [child];
        for (const item of items) {
            if (key === "p:graphicFrame") {
                out.push(item);
                continue;
            }
            collectGraphicFrames(item, out);
        }
    }
}

/** 从 numRef 缓存提取数值数组（按 idx 排序） */
function numCacheValues(node: any): number[] {
    if (!node || !node["c:numRef"] || !node["c:numRef"]["c:numCache"]) return [];
    const pts = node["c:numRef"]["c:numCache"]["c:pt"];
    if (!pts) return [];
    const arr = Array.isArray(pts) ? pts : [pts];
    return arr
        .map((p: any) => ({ idx: parseInt(p["attrs"]?.["idx"] ?? "0", 10), v: parseFloat(p["c:v"]) }))
        .sort((a: any, b: any) => a.idx - b.idx)
        .map((o: any) => o.v);
}

/** 从系列节点的 c:tx/c:strRef 提取系列名 */
function chartSeriesName(ser: any): string | undefined {
    const pt = PPTXXmlUtils.getTextByPathList(ser, ["c:tx", "c:strRef", "c:strCache", "c:pt"]);
    if (!pt) return undefined;
    return Array.isArray(pt) ? pt[0]["c:v"] : pt["c:v"];
}

/** 气泡图：xVal / yVal / bubbleSize → { key, values:[{x,y,size}] } */
function extractBubbleData(serNode: any): Array<Record<string, unknown>> {
    if (!serNode) return [];
    const sers = Array.isArray(serNode) ? serNode : [serNode];
    return sers.map((ser: any, i: number) => {
        const xs = numCacheValues(ser["c:xVal"]);
        const ys = numCacheValues(ser["c:yVal"]);
        const sizes = numCacheValues(ser["c:bubbleSize"]);
        const n = Math.max(xs.length, ys.length, sizes.length);
        const values = [];
        for (let k = 0; k < n; k++) {
            values.push({ x: xs[k], y: ys[k], size: sizes[k] });
        }
        return { key: chartSeriesName(ser) || `Series ${i + 1}`, values, xlabels: {}, style: {} };
    });
}

/** 股票图：open/high/low/close → { key, values:[[open,close,low,high]...], xlabels }（ECharts 蜡烛图格式） */
function extractStockData(serNode: any): Array<Record<string, unknown>> {
    if (!serNode) return [];
    const sers = Array.isArray(serNode) ? serNode : [serNode];
    return sers.map((ser: any, i: number) => {
        const open = numCacheValues(ser["c:openVal"]);
        const high = numCacheValues(ser["c:highVal"]);
        const low = numCacheValues(ser["c:lowVal"]);
        const close = numCacheValues(ser["c:closeVal"]);
        const n = Math.max(open.length, high.length, low.length, close.length);
        const values: number[][] = [];
        const xlabels: string[] = [];
        for (let k = 0; k < n; k++) {
            values.push([open[k], close[k], low[k], high[k]]);
            xlabels.push(String(k + 1));
        }
        return { key: chartSeriesName(ser) || `Series ${i + 1}`, values, xlabels, style: {} };
    });
}

/**
 * Process message queue for charts
 * @param {Array} queue - Message queue
 * @param {Object} result - Result object to store chart data
 */
function processMsgQueue(queue: Array<{ type?: string; data?: Record<string, unknown> }>, result: { charts: Array<Record<string, unknown>> }) {
    for (const msg of queue) {
        if (msg.type === "chart" || msg.type === "createChart") {
            const chartObj = msg.data;
            if (chartObj) {
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
}

/**
 * Process single chart message
 * @param {Object} data - Chart data
 * @param {Object} callbacks - Callback functions
 */
function processSingleMsg(data: any, callbacks: any) {
    const { chartId, chartType, chartData } = data;
    let chartDataArray = [];
    let chart: any = null;

    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
        return;
    }

    switch (chartType) {
        case "lineChart":
            chartDataArray = chartData;
            
            chart = nv.models.lineChart().useInteractiveGuideline(true);
            if (chartData[0]?.xlabels) {
                chart.xAxis.tickFormat((d: number) => chartData[0].xlabels[d] || d);
            }
            break;

        case "barChart":
            chartDataArray = chartData;
            
            chart = nv.models.multiBarChart();
            if (chartData[0]?.xlabels) {
                chart.xAxis.tickFormat((d: number) => chartData[0].xlabels[d] || d);
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
                chart.xAxis.tickFormat((d: number) => chartData[0].xlabels[d] || d);
            }
            break;

        case "scatterChart":
            for (const i of chartData.keys()){
                const arr = [];
                if (Array.isArray(chartData[i])) {
                    for (const j of chartData[i].keys()){
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
    processSingleMsg,
    extractChartsFromSlide
};
