/**
 * HTML 生成模块
 * 提供全局 CSS 生成、表格生成、图表生成等功能
 */
// 全局变量引用
import { PPTXUtils } from './core/utils';
import { PPTXColorUtils } from './core/color';
import { PPTXParser } from './parser';
import { PPTXTableUtils } from './shape/table';

let settings: any = {}; // 配置将由 configure 方法设置
// 图表 ID 计数器
let chartID = 0;
/**
 * 生成全局 CSS
 * @returns {string} 生成的 CSS 文本
 */
function genGlobalCSS() {
    let cssText = "";
    // 从 PPTXParser 获取 styleTable
    const styleTable = PPTXParser.styleTable || {};
    const slideWidth = PPTXParser.slideWidth || 960;
    for (const key in styleTable) {
        const tagname = "";
        //ADD suffix
        cssText += `${tagname} .${styleTable[key]["name"]}${(styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : ""}{${styleTable[key]["text"]}}\n`;
    }
    cssText += " .slide{margin-bottom: 5px;}\n";
    if (settings.slideMode && settings.slideType == "divs2slidesjs") {
        cssText += `#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: ${slideWidth}px;}\n`;
    }
    return cssText;
}
/**
 * 获取填充颜色
 */
function getSolidFill(fillNode, clrMap, phClr, warpObj) {
    return PPTXColorUtils.getSolidFill(fillNode, clrMap, phClr, warpObj);
}
/**
 * 获取形状填充
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {string} 填充样式字符串
 */
function getShapeFill(node, warpObj) {
    if (!node)
        return "";
    const fillType = PPTXColorUtils.getFillType(node);
    let fillColor;
    if (fillType == "NO_FILL") {
        return "";
    }
    else if (fillType == "SOLID_FILL") {
        const shpFill = node["a:solidFill"];
        fillColor = PPTXColorUtils.getSolidFill(shpFill, undefined, undefined, warpObj);
    }
    if (fillColor) {
        return `background-color: #${fillColor};`;
    }
    return "";
}
/**
 * 生成表格 HTML
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {string} 表格 HTML 字符串
 */
function genTable(node, warpObj) {
    const order = node["attrs"]["order"];
    const tableNode = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    const slideFactor = PPTXParser.slideFactor || (96 / 914400);
    const styleTable = PPTXParser.styleTable || {};
    const tableStyles = PPTXParser.tableStyles || {};
    if (!tableNode) {
        const result = "";
        return result;
        return `<div class='block table' style='z-index: ${order};'>表格</div>`;
    }
    const getTblPr = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
    const getColsGrid = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
    let tblDir = "";
    if (getTblPr !== undefined) {
        const isRTL = getTblPr["attrs"]["rtl"];
        tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
    }
    let firstRowAttr = getTblPr["attrs"]["firstRow"];
    const firstColAttr = getTblPr["attrs"]["firstCol"];
    const lastRowAttr = getTblPr["attrs"]["lastRow"];
    const lastColAttr = getTblPr["attrs"]["lastCol"];
    const bandRowAttr = getTblPr["attrs"]["bandRow"];
    const bandColAttr = getTblPr["attrs"]["bandCol"];
    const tblStylAttrObj = {
        isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
        isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
        isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
        isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
        isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
        isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
    };
    let thisTblStyle;
    const tbleStyleId = getTblPr["a:tableStyleId"];
    if (tbleStyleId !== undefined && tableStyles) {
        const tbleStylList = PPTXUtils.getTextByPathList(tableStyles, ["a:tblStyleLst", "a:tblStyle"]);
        if (tbleStylList !== undefined) {
            if (Array.isArray(tbleStylList)) {
                for (let k = 0; k < tbleStylList.length; k++) {
                    if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                        thisTblStyle = tbleStylList[k];
                    }
                }
            }
            else {
                if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                    thisTblStyle = tbleStylList;
                }
            }
        }
    }
    if (thisTblStyle !== undefined) {
        thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
        warpObj["thisTblStyle"] = thisTblStyle;
    }
    const tblStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
    const tblBorderStyl = PPTXUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
    let tbl_borders = "";
    if (tblBorderStyl !== undefined) {
        // @ts-ignore - getTableBorders is from PPTXTableUtils
        tbl_borders = PPTXTableUtils.getTableBorders(tblBorderStyl, warpObj);
    }
    let tbl_bgcolor = "";
    let tbl_bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
    if (tbl_bgFillschemeClr !== undefined) {
        tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgFillschemeClr === undefined) {
        tbl_bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
        tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgcolor !== "") {
        tbl_bgcolor = `background-color: #${tbl_bgcolor};`;
    }
    let tableHtml = `<table ${tblDir} style='border-collapse: collapse;${PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined)}${PPTXUtils.getSize(xfrmNode, undefined, undefined)} z-index: ${order};${tbl_borders};${tbl_bgcolor}'>`;
    let trNodes = tableNode["a:tr"];
    if (!trNodes) {
        tableHtml += "</table>";
        return tableHtml;
    }
    if (!Array.isArray(trNodes)) {
        trNodes = [trNodes];
    }
    const rowSpanAry = [];
    let totalColSpan = 0;
    for (let i = 0; i < trNodes.length; i++) {
        const rowHeightParam = trNodes[i]["attrs"]["h"];
        let rowHeight = 0;
        let rowsStyl = "";
        if (rowHeightParam !== undefined) {
            rowHeight = parseInt(rowHeightParam) * slideFactor;
            rowsStyl += `height:${rowHeight}px;`;
        }
        let fillColor = "";
        let row_borders = "";
        let fontClrPr = "";
        let fontWeight = "";
        if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
            const bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            const rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontColor !== undefined) {
                    fontClrPr = local_fontColor;
                }
                const local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight != "") {
                    fontWeight = local_fontWeight;
                }
            }
        }
        if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {
            let bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            const borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
            if (borderStyl !== undefined) {
                // @ts-ignore - getTableBorders is from PPTXTableUtils
                const local_row_borders = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                if (local_row_borders != "") {
                    row_borders = local_row_borders;
                }
            }
            let rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                const local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    fontClrPr = local_fontClrPr;
                }
                const local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }
        }
        rowsStyl += ((row_borders !== undefined) ? row_borders : "");
        rowsStyl += ((fontClrPr !== undefined) ? ` color: #${fontClrPr};` : "");
        rowsStyl += ((fontWeight != "") ? ` font-weight:${fontWeight};` : "");
        if (fillColor !== undefined && fillColor != "") {
            rowsStyl += `background-color: #${fillColor};`;
        }
        tableHtml += `<tr style='${rowsStyl}'>`;
        const tcNodes = trNodes[i]["a:tc"];
        if (tcNodes !== undefined) {
            if (Array.isArray(tcNodes)) {
                let j = 0;
                if (rowSpanAry.length == 0) {
                    rowSpanAry.length = tcNodes.length;
                    rowSpanAry.fill(0);
                }
                totalColSpan = 0;
                while (j < tcNodes.length) {
                    if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                        let a_sorce = "";
                        if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                            a_sorce = "a:firstCol";
                        }
                        else if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                            a_sorce = "a:lastCol";
                        }
                        const cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj, styleTable);
                        const text = cellParmAry[0];
                        const colStyl = cellParmAry[1];
                        const cssName = cellParmAry[2];
                        const rowSpan = cellParmAry[3];
                        const colSpan = cellParmAry[4];
                        if (rowSpan !== undefined) {
                            tableHtml += `<td class='${cssName}' rowspan ='${parseInt(rowSpan)}' style='${colStyl}'>${text}</td>`;
                            rowSpanAry[j] = parseInt(rowSpan) - 1;
                        }
                        else if (colSpan !== undefined) {
                            tableHtml += `<td class='${cssName}' colspan = '${parseInt(colSpan)}' style='${colStyl}'>${text}</td>`;
                            totalColSpan = parseInt(colSpan) - 1;
                        }
                        else {
                            tableHtml += `<td class='${cssName}' style = '${colStyl}'>${text}</td>`;
                        }
                    }
                    else {
                        if (rowSpanAry[j] != 0) {
                            rowSpanAry[j] -= 1;
                        }
                        if (totalColSpan != 0) {
                            totalColSpan--;
                        }
                    }
                    j++;
                }
            }
            else {
                let a_sorce = "";
                if (tblStylAttrObj["isFrstColAttr"] == 1) {
                    a_sorce = "a:firstCol";
                }
                let cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj, styleTable);
                let text = cellParmAry[0];
                let colStyl = cellParmAry[1];
                let cssName = cellParmAry[2];
                let rowSpan = cellParmAry[3];
                if (rowSpan !== undefined) {
                    tableHtml += `<td class='${cssName}' rowspan='${parseInt(rowSpan)}' style = '${colStyl}'>${text}</td>`;
                }
                else {
                    tableHtml += `<td class='${cssName}' style='${colStyl}'>${text}</td>`;
                }
            }
        }
        tableHtml += "</tr>";
    }
    tableHtml += "</table>";
    return tableHtml;
}
/**
 * 生成图表 HTML
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {string} 图表 HTML 字符串
 */
async function genChart(node, warpObj) {
    const order = node["attrs"]["order"];
    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    const readXmlFile = PPTXParser ? PPTXParser.readXmlFile : (() => null);
    const result = `<div id='chart${chartID}' class='block content' style='${PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined)}${PPTXUtils.getSize(xfrmNode, undefined, undefined)} z-index: ${order};'></div>`;
    const rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    const refName = warpObj["slideResObj"][rid]["target"];
    // 读取图表文件
    const content = await readXmlFile(warpObj["zip"], refName);
    if (!content) {
        chartID++;
        return result;
    }
    const plotArea = PPTXUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);
    if (!plotArea) {
        chartID++;
        return result;
    }
    // 收集所有有效的图表数据
    const chartDatas = [];
    // 处理不同类型的图表
    const chartTypes = [
        { key: "c:lineChart", type: "lineChart" },
        { key: "c:barChart", type: "barChart" },
        { key: "c:pieChart", type: "pieChart" },
        { key: "c:pie3DChart", type: "pie3DChart" },
        { key: "c:areaChart", type: "areaChart" },
        { key: "c:scatterChart", type: "scatterChart" }
    ];
    for (let i = 0; i < chartTypes.length; i++) {
        const chartType = chartTypes[i];
        const seriesNode = plotArea[chartType.key];
        if (seriesNode) {
            // 确保 seriesNode 是数组
            let seriesArray = seriesNode;
            if (!Array.isArray(seriesArray)) {
                seriesArray = [seriesArray];
            }
            // 过滤掉空的系列
            const validSeries = seriesArray.filter((series) => series && series["c:ser"]);
            if (validSeries.length > 0) {
                const chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": `chart${chartID}`,
                        "chartType": chartType.type,
                        "chartData": extractChartData(validSeries[0]["c:ser"]),
                        "hasMultipleSeries": validSeries.length > 1
                    }
                };
                // 如果有多个系列，只使用第一个系列
                if (validSeries.length > 1) {
                    // silently ignore additional series
                }
                chartDatas.push(chartData);
            }
        }
    }
    // 如果没有找到任何图表数据，尝试更宽松的搜索
    if (chartDatas.length === 0) {
        // fallback extraction
        // 查找任何包含 c:ser 的节点
        for (const key in plotArea) {
            if (key.indexOf('Chart') > -1 && plotArea[key]["c:ser"]) {
                const fallbackData = {
                    "type": "createChart",
                    "data": {
                        "chartID": `chart${chartID}`,
                        "chartType": "lineChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                chartDatas.push(fallbackData);
                break;
            }
        }
    }
    // Store all chart data for later processing
    if (chartDatas.length > 0) {
        if (!PPTXHtml.MsgQueue) {
            PPTXHtml.MsgQueue = [];
        }
        // 将所有图表数据添加到队列
        for (let j = 0; j < chartDatas.length; j++) {
            PPTXHtml.MsgQueue.push(chartDatas[j]);
        }
    }
    chartID++;
    return result;
}
/**
 * 生成图表数据 - 增强版本，更好的错误处理和数据验证
 * @param {Object} serNode - 系列节点
 * @returns {Array} 提取的图表数据
 */
function extractChartData(serNode) {
    const dataMat = [];
    // 输入验证
    if (serNode === undefined || serNode === null) {
        return dataMat;
    }
    // 确保 serNode 是数组
    let seriesArray = serNode;
    if (!Array.isArray(seriesArray)) {
        seriesArray = [seriesArray];
    }
    if (seriesArray.length === 0) {
        return dataMat;
    }
    // 辅助函数：安全获取路径值
    const safeGetPath = (obj, path, defaultValue) => {
        try {
            let result = obj;
            for (let i = 0; i < path.length; i++) {
                if (result === undefined || result === null)
                    return defaultValue;
                result = result[path[i]];
            }
            return result !== undefined ? result : defaultValue;
        }
        catch (e) {
            return defaultValue;
        }
    };
    // 辅助函数：安全遍历节点数组
    const safeEachElement = (nodes, processFunc) => {
        if (!nodes || !Array.isArray(nodes)) {
            if (nodes) {
                return processFunc(nodes, 0);
            }
            return '';
        }
        let result = '';
        for (let i = 0; i < nodes.length; i++) {
            if (nodes[i]) {
                result += processFunc(nodes[i], i);
            }
        }
        return result;
    };
    // 处理第一种图表格式：有 c:xVal 和 c:yVal 的简单格式
    const xValNode = safeGetPath(seriesArray[0], ["c:xVal"], null);
    const yValNode = safeGetPath(seriesArray[0], ["c:yVal"], null);
    if (xValNode && yValNode) {
        try {
            // 处理 X 值
            const xCache = safeGetPath(xValNode, ["c:numRef", "c:numCache", "c:pt"], null);
            if (xCache) {
                const xDataRow = [];
                safeEachElement(xCache, (pointNode) => {
                    const value = safeGetPath(pointNode, ["c:v"], null);
                    if (value !== null) {
                        const numValue = parseFloat(value);
                        if (!isNaN(numValue)) {
                            xDataRow.push(numValue);
                        }
                    }
                });
                if (xDataRow.length > 0) {
                    dataMat.push(xDataRow);
                }
            }
            // 处理 Y 值
            const yCache = safeGetPath(yValNode, ["c:numRef", "c:numCache", "c:pt"], null);
            if (yCache) {
                const yDataRow = [];
                safeEachElement(yCache, (pointNode) => {
                    const value = safeGetPath(pointNode, ["c:v"], null);
                    if (value !== null) {
                        const numValue = parseFloat(value);
                        if (!isNaN(numValue)) {
                            yDataRow.push(numValue);
                        }
                    }
                });
                if (yDataRow.length > 0) {
                    dataMat.push(yDataRow);
                }
            }
            // 如果成功提取到数据，返回
            if (dataMat.length >= 2) {
                return dataMat;
            }
        }
        catch (e) {
            // extraction failed
        }
    }
    // 处理第二种图表格式：复杂的多系列格式
    try {
        safeEachElement(seriesArray, (seriesItem) => {
            if (!seriesItem)
                return '';
            const dataRow = [];
            let colName = safeGetPath(seriesItem, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"], null);
            if (colName === null) {
                // 如果没有名称，使用索引
                const seriesIndex = seriesArray.indexOf(seriesItem);
                colName = `Series ${seriesIndex + 1}`;
            }
            // 提取类别标签
            const rowNames = {};
            const catStrRef = safeGetPath(seriesItem, ["c:cat", "c:strRef", "c:strCache", "c:pt"], null);
            const catNumRef = safeGetPath(seriesItem, ["c:cat", "c:numRef", "c:numCache", "c:pt"], null);
            const catPoints = catStrRef || catNumRef;
            if (catPoints) {
                safeEachElement(catPoints, (pointNode) => {
                    const idx = safeGetPath(pointNode, ["attrs", "idx"], null);
                    const val = safeGetPath(pointNode, ["c:v"], null);
                    if (idx !== null && val !== null) {
                        rowNames[idx] = val;
                    }
                });
            }
            // 提取值数据
            const valNode = safeGetPath(seriesItem, ["c:val", "c:numRef", "c:numCache", "c:pt"], null);
            if (valNode) {
                safeEachElement(valNode, (pointNode) => {
                    const idx = safeGetPath(pointNode, ["attrs", "idx"], null);
                    const val = safeGetPath(pointNode, ["c:v"], null);
                    if (idx !== null && val !== null) {
                        const numValue = parseFloat(val);
                        if (!isNaN(numValue)) {
                            dataRow.push({ x: parseInt(idx), y: numValue });
                        }
                    }
                });
            }
            // 只有当有实际数据时才添加到结果中
            if (dataRow.length > 0) {
                dataMat.push({
                    key: colName,
                    values: dataRow,
                    xlabels: rowNames
                });
            }
            return '';
        });
        if (dataMat.length > 0) {
            return dataMat;
        }
    }
    catch (e) {
        // extraction failed
    }
    return [];
}
/**
 * Convert plain numeric lists to proper HTML numbered lists
 * @param {NodeList|Array} elem - 要处理的元素列表
 */
function setNumericBullets(elem) {
    if (PPTXUtils && PPTXUtils.getNumTypeNum) {
        const prgrphs_arry = elem;
        for (let i = 0; i < prgrphs_arry.length; i++) {
            const element = prgrphs_arry[i];
            const buSpan = element.querySelectorAll('.numeric-bullet-style');
            if (buSpan.length > 0) {
                let prevBultTyp = "";
                let prevBultLvl = "";
                let buletIndex = 0;
                const tmpArry = [];
                let tmpArryIndx = 0;
                const buletTypSrry = [];
                for (let j = 0; j < buSpan.length; j++) {
                    const span = buSpan[j];
                    const bult_typ = span.dataset.bulltname;
                    const bult_lvl = span.dataset.bulltlvl;
                    if (buletIndex == 0) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        tmpArry[tmpArryIndx] = buletIndex;
                        buletTypSrry[tmpArryIndx] = bult_typ;
                        buletIndex++;
                    }
                    else {
                        if (bult_typ == prevBultTyp && bult_lvl == prevBultLvl) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            buletIndex++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                        }
                        else if (bult_typ != prevBultTyp && bult_lvl == prevBultLvl) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex = 1;
                        }
                        else if (bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx++;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex = 1;
                        }
                        else if (bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArryIndx--;
                            buletIndex = tmpArry[tmpArryIndx] + 1;
                        }
                    }
                    const numIdx = PPTXUtils.getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                    span.innerHTML = numIdx;
                }
            }
        }
    }
    else {
        // Fallback to simple list conversion if PPTXUtils is not available
        let elements = elem;
        if (!elements.length && elements.nodeType === 1) {
            elements = [elements];
        }
        for (let i = 0; i < elements.length; i++) {
            let element = elements[i];
            const lis = element.querySelectorAll('li');
            for (let j = 0; j < lis.length; j++) {
                const li = lis[j];
                const html = li.innerHTML;
                // If it starts with a number and a dot, treat as numbered list item
                if (/^\d+\.\s/.test(html)) {
                    // Ensure parent is ol if not already
                    const parent = li.parentNode;
                    if (parent && parent.tagName !== 'OL') {
                        if (parent.tagName === 'UL') {
                            const ol = document.createElement('ol');
                            while (parent.firstChild) {
                                ol.appendChild(parent.firstChild);
                            }
                            parent.parentNode.replaceChild(ol, parent);
                        }
                    }
                }
            }
        }
    }
}
/**
 * Process message queue and update UI accordingly
 * @param {Array} msgQueue - 消息队列
 */
function processMsgQueue(msgQueue) {
    if (!msgQueue || msgQueue.length === 0)
        return;
    // Process each message
    for (let i = 0; i < msgQueue.length; i++) {
        const msg = msgQueue[i];
        if (msg && msg.type === "createChart" && msg.data) {
            processSingleMsg(msg.data);
        }
    }
    // Clear after processing
    msgQueue.length = 0;
}
/**
 * 处理单个消息
 * @param {Object} d - 消息数据
 * @returns {boolean} 是否处理成功
 */
function processSingleMsg(d: any): boolean {
    // 检查外部图表库是否可用
    // @ts-ignore - External libraries nv and d3
    if (typeof (globalThis as any).nv === 'undefined' || typeof (globalThis as any).d3 === 'undefined') {
        return false;
    }
    // @ts-ignore
    let chartID = d.chartID;
    let chartType = d.chartType;
    let chartData = d.chartData;
    let data: any[] = [];
    let chart: any = null;
    let isDone = false;
    switch (chartType) {
        case "lineChart":
            data = chartData;
            // @ts-ignore
            chart = (globalThis as any).nv.models.lineChart()
                .useInteractiveGuideline(true);
            chart.xAxis.tickFormat((val) => chartData[0].xlabels[val] || val);
            break;
        case "barChart":
            data = chartData;
            // @ts-ignore
            chart = (globalThis as any).nv.models.multiBarChart();
            chart.xAxis.tickFormat((val) => chartData[0].xlabels[val] || val);
            break;
        case "pieChart":
        case "pie3DChart":
            if (chartData.length > 0) {
                data = chartData[0].values;
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.pieChart();
            break;
        case "areaChart":
            data = chartData;
            // @ts-ignore
            chart = (globalThis as any).nv.models.stackedAreaChart()
                .clipEdge(true)
                .useInteractiveGuideline(true);
            chart.xAxis.tickFormat((val) => chartData[0].xlabels[val] || val);
            break;
        case "scatterChart":
            for (let i = 0; i < chartData.length; i++) {
                const arr = [];
                for (let j = 0; j < chartData[i].length; j++) {
                    arr.push({ x: j, y: chartData[i][j] });
                }
                data.push({ key: `data${i + 1}`, values: arr });
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.scatterChart()
                .showDistX(true)
                .showDistY(true)
                .color((globalThis as any).d3.scale.category10().range());
            chart.xAxis.axisLabel('X').tickFormat((globalThis as any).d3.format('.02f'));
            chart.yAxis.axisLabel('Y').tickFormat((globalThis as any).d3.format('.02f'));
            break;
        default:
    }
    if (chart !== null) {
        const chartElement = document.getElementById(chartID);
        if (chartElement) {
            // @ts-ignore
            (globalThis as any).d3.select(`#${chartID}`)
                .append("svg")
                .datum(data)
                .transition().duration(500)
                .call(chart);
            // @ts-ignore
            (globalThis as any).nv.utils.windowResize(chart.update);
            isDone = true;
        }
        else {
            // chart element not found
        }
    }
    return isDone;
}
// 获取背景
// getBackground 和 getSlideBackgroundFill 已移至 PPTXBackgroundUtils 模块
// 更新加载进度条 - 使用回调而不是直接操作 DOM
let progressCallback = null;
/**
 * 设置进度回调函数
 * @param {Function} callback - 回调函数
 */
function setProgressCallback(callback) {
    progressCallback = callback;
}
/**
 * 更新进度条
 * @param {number} percent - 进度百分比
 */
function updateProgressBar(percent) {
    if (progressCallback) {
        progressCallback(percent);
    }
}
// 公开 API
/**
 * PPTXHtml 模块
 * @namespace PPTXHtml
 */
const PPTXHtml = {
    genGlobalCSS,
    genTable,
    genChart,
    setNumericBullets,
    processMsgQueue,
    processSingleMsg,
    extractChartData,
    setProgressCallback,
    MsgQueue: null // 图表数据队列，将在 processMsgQueue 中初始化
};
export { PPTXHtml };
//# sourceMappingURL=html.js.map