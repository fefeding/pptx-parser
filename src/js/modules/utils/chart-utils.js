/**
 * 图表工具函数模块
 * 提供图表数据提取和处理功能
 * 依赖 NVD3.js 和 D3.js 库
 */


var PPTXChartUtils = (function() {
let isDone = false;

/**
 * extractChartData - 提取图表数据
 * @param {Object} serNode - 序列节点
 * @returns {Array} 数据矩阵
 */
    function extractChartData(serNode) {
    const dataMat = [];

    if (serNode === undefined) {
        return dataMat;
    }

    if (serNode["c:xVal"] !== undefined) {
        const dataRow = [];
        eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
        const dataRow2 = [];
        eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
            dataRow2.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow2);
    } else {
        eachElement(serNode, function (innerNode, index) {
            const dataRow = [];
            const colName = getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

            // Category (string or number)
            const rowNames = {};
            if (getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function (innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            } else if (getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            }

            // Value
            if (getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                    return "";
                });
            }

            dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
            return "";
        });
    }

    return dataMat;
}

/**
 * processMsgQueue - 处理消息队列（用于图表渲染）
 * @param {Array} queue - 消息队列
 */
    function processMsgQueue(queue) {
    for (let i = 0; i < queue.length; i++) {
        processSingleMsg(queue[i].data);
    }
}

/**
 * processSingleMsg - 处理单个消息（用于渲染单个图表）
 * @param {Object} d - 数据对象
 */
    function processSingleMsg(d) {
    const chartID = d.chartID;
    const chartType = d.chartType;
    const chartData = d.chartData;

    let data = [];
    let chart = null;

    switch (chartType) {
        case "lineChart":
            data = chartData;
            chart = nv.models.lineChart()
                .useInteractiveGuideline(true);
            chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
            break;
        case "barChart":
            data = chartData;
            chart = nv.models.multiBarChart();
            chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
            break;
        case "pieChart":
        case "pie3DChart":
            if (chartData.length > 0) {
                data = chartData[0].values;
            }
            chart = nv.models.pieChart();
            break;
        case "areaChart":
            data = chartData;
            chart = nv.models.stackedAreaChart()
                .clipEdge(true)
                .useInteractiveGuideline(true);
            chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
            break;
        case "scatterChart":
            for (let i = 0; i < chartData.length; i++) {
                const arr = [];
                for (let j = 0; j < chartData[i].length; j++) {
                    arr.push({ x: j, y: chartData[i][j] });
                }
                data.push({ key: 'data' + (i + 1), values: arr });
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

    if (chart !== null) {
        d3.select("#" + chartID)
            .append("svg")
            .datum(data)
            .transition().duration(500)
            .call(chart);

        nv.utils.windowResize(chart.update);
        isDone = true;
    }
}

/**
 * getIsDone - 获取图表渲染完成状态
 * @returns {boolean} 是否完成
 */
    function getIsDone() {
    return isDone;
}

/**
 * setIsDone - 设置图表渲染完成状态
 * @param {boolean} done - 完成状态
 */
    function setIsDone(done) {
    isDone = done;
}


    return {
        extractChartData: extractChartData,
        processMsgQueue: processMsgQueue,
        processSingleMsg: processSingleMsg,
        getIsDone: getIsDone,
        setIsDone: setIsDone
    };
})();