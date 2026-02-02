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
    
    const chartSpace = PPTXUtils.getTextByPathList(content, ["c:chartSpace"]);
    const chart = PPTXUtils.getTextByPathList(content, ["c:chartSpace", "c:chart"]);
    const plotArea = PPTXUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);
    
    // 检查是否有外部颜色和样式引用
    let colorStyleData = null;
    let chartStyleData = null;
    
    // 检查是否有外部数据引用（可能包含颜色和样式）
    const externalDataRef = PPTXUtils.getTextByPathList(content, ["c:chartSpace", "c:externalData"]);
    if (externalDataRef && externalDataRef["attrs"] && externalDataRef["attrs"]["r:id"]) {
        const extRid = externalDataRef["attrs"]["r:id"];
        const extRefName = warpObj["slideResObj"][extRid]?.["target"];
        if (extRefName) {
            try {
                const extContent = await readXmlFile(warpObj["zip"], extRefName);
                if (extContent) {
                    // 尝试解析颜色样式文件
                    if (extRefName.includes('colors')) {
                        colorStyleData = extContent;
                    } else if (extRefName.includes('style')) {
                        chartStyleData = extContent;
                    }
                }
            } catch (e) {
                console.warn('无法读取外部数据文件:', extRefName, e);
            }
        }
    }
    
    // 尝试直接查找可能的colors和style文件
    // 在PPTX文件中，这些通常在charts文件夹中
    try {
        // 检查是否有颜色样式文件的引用
        const relsPath = refName.replace('charts/', 'charts/_rels/').replace('.xml', '.xml.rels');
        const relsContent = await readXmlFile(warpObj["zip"], relsPath);
        
        if (relsContent && relsContent["Relationships"]) {
            const relationships = Array.isArray(relsContent["Relationships"]["Relationship"]) 
                ? relsContent["Relationships"]["Relationship"] 
                : [relsContent["Relationships"]["Relationship"]];
            
            for (const rel of relationships) {
                if (rel && rel["attrs"]) {
                    const relType = rel["attrs"]["Type"];
                    const relTarget = rel["attrs"]["Target"];
                    const relId = rel["attrs"]["Id"];
                    
                    // 检查是否是颜色或样式引用
                    if (relTarget && (relTarget.includes('colors') || relTarget.includes('style'))) {
                        const fullTargetPath = `charts/${relTarget}`;
                        try {
                            const styleContent = await readXmlFile(warpObj["zip"], fullTargetPath);
                            if (styleContent) {
                                if (relTarget.includes('colors')) {
                                    colorStyleData = styleContent;
                                } else if (relTarget.includes('style')) {
                                    chartStyleData = styleContent;
                                }
                            }
                        } catch (e) {
                            console.warn('无法读取样式文件:', fullTargetPath, e);
                        }
                    }
                }
            }
        }
    } catch (e) {
        console.warn('无法读取关系文件:', e);
    }
    
    if (!plotArea) {
        chartID++;
        return result;
    }
    
    // 提取图表标题
    const titleNode = PPTXUtils.getTextByPathList(chart, ["c:title"]);
    let chartTitle = null;
    if (titleNode) {
        // 检查是否存在c:tx节点
        let txNode = titleNode["c:tx"];
        
        // 如果直接在title下没有找到c:tx，尝试其他可能的位置
        if (!txNode) {
            // 遍历titleNode的所有键来找c:tx
            for (const key in titleNode) {
                if (key === 'c:tx' && typeof titleNode[key] === 'object') {
                    txNode = titleNode[key];
                    break;
                }
                // 或者检查嵌套结构
                if (typeof titleNode[key] === 'object' && titleNode[key]['c:tx']) {
                    txNode = titleNode[key]['c:tx'];
                    break;
                }
            }
        }
        
        // 如果还是找不到c:tx，可能在某些PPTX文件中标题结构不同
        if (!txNode) {
            // 检查是否在其他结构中，如直接在title节点下有文本
            // 从你提供的结构来看，文本可能在 a:p -> a:r -> a:t 结构中
            const findAllTextNodes = (node) => {
                let texts = [];
                if (typeof node === 'object' && node !== null) {
                    for (const key in node) {
                        if (key === 'a:t' && typeof node[key] === 'string') {
                            texts.push(node[key]);
                        } else if (typeof node[key] === 'object') {
                            texts = texts.concat(findAllTextNodes(node[key]));
                        }
                    }
                }
                return texts;
            };
            
            const allTexts = findAllTextNodes(titleNode);
            if (allTexts.length > 0) {
                chartTitle = allTexts.join('');
            }
        } else {
            // c:tx 存在，按原来逻辑处理
            // 尝试多种可能的标题路径
            let titleText = null;
            // 由于 a:r 是数组，需要遍历处理
            const richNode = PPTXUtils.getTextByPathList(txNode, ["c:rich"]);
            if (richNode && richNode["a:p"]) {
                const paragraphs = Array.isArray(richNode["a:p"]) ? richNode["a:p"] : [richNode["a:p"]];
                const texts = [];
                for (const p of paragraphs) {
                    if (p && p["a:r"]) {
                        const textRuns = Array.isArray(p["a:r"]) ? p["a:r"] : [p["a:r"]];
                        for (const textRun of textRuns) {
                            if (textRun && textRun["a:t"]) {
                                texts.push(textRun["a:t"]);
                            }
                        }
                    }
                }
                if (texts.length > 0) {
                    titleText = texts.join('');
                }
            }
            
            // 如果上面的路径没有找到标题，尝试另一种格式
            if (!titleText) {
                const pNodes = PPTXUtils.getTextByPathList(txNode, ["c:rich", "a:p"]);
                if (pNodes) {
                    const pArray = Array.isArray(pNodes) ? pNodes : [pNodes];
                    const texts = [];
                    for (const p of pArray) {
                        if (p && p["a:r"]) {
                            const textRuns = Array.isArray(p["a:r"]) ? p["a:r"] : [p["a:r"]];
                            for (const textRun of textRuns) {
                                if (textRun && textRun["a:t"]) {
                                    texts.push(textRun["a:t"]);
                                }
                            }
                        }
                    }
                    if (texts.length > 0) {
                        titleText = texts.join('');
                    }
                }
            }
            
            // 如果还是没有找到，尝试strRef格式
            if (!titleText) {
                const strRefNode = PPTXUtils.getTextByPathList(txNode, ["c:strRef", "c:strCache", "c:pt"]);
                if (strRefNode) {
                    const ptNodes = Array.isArray(strRefNode) ? strRefNode : [strRefNode];
                    titleText = ptNodes.map(pt => pt["c:v"] || "").join(" ");
                }
            }
            
            // 如果还是没有找到，尝试提取段落中的文本运行
            if (!titleText) {
                const richNode = PPTXUtils.getTextByPathList(txNode, ["c:rich"]);
                if (richNode && richNode["a:p"]) {
                    const paragraphs = Array.isArray(richNode["a:p"]) ? richNode["a:p"] : [richNode["a:p"]];
                    const texts = [];
                    for (const p of paragraphs) {
                        if (p && p["a:r"]) {
                            const textRuns = Array.isArray(p["a:r"]) ? p["a:r"] : [p["a:r"]];
                            for (const textRun of textRuns) {
                                if (textRun && textRun["a:t"]) {
                                    texts.push(textRun["a:t"]);
                                }
                            }
                        }
                    }
                    if (texts.length > 0) {
                        titleText = texts.join('');
                    }
                }
            }
            
            // 再尝试另一种结构：可能直接在 a:p 下有 a:r 和 a:t
            if (!titleText) {
                const pNodes = PPTXUtils.getTextByPathList(txNode, ["a:p"]);
                if (pNodes) {
                    const texts = [];
                    const pArray = Array.isArray(pNodes) ? pNodes : [pNodes];
                    for (const pNode of pArray) {
                        if (pNode && pNode["a:r"]) {
                            const rNodes = Array.isArray(pNode["a:r"]) ? pNode["a:r"] : [pNode["a:r"]];
                            for (const rNode of rNodes) {
                                if (rNode && rNode["a:t"]) {
                                    texts.push(rNode["a:t"]);
                                }
                            }
                        }
                    }
                    if (texts.length > 0) {
                        titleText = texts.join('');
                    }
                }
            }
            
            // 最后的兜底方案：深度遍历节点寻找所有a:t标签
            if (!titleText) {
                const findAllTextNodes = (node) => {
                    let texts = [];
                    if (typeof node === 'object' && node !== null) {
                        for (const key in node) {
                            if (key === 'a:t' && typeof node[key] === 'string') {
                                texts.push(node[key]);
                            } else if (typeof node[key] === 'object') {
                                texts = texts.concat(findAllTextNodes(node[key]));
                            }
                        }
                    }
                    return texts;
                };
                
                const allTexts = findAllTextNodes(txNode);
                if (allTexts.length > 0) {
                    titleText = allTexts.join('');
                }
            }
            
            if (titleText) {
                chartTitle = Array.isArray(titleText) ? titleText.join(' ') : titleText;
            }
        }
    }
    
    // 提取Y轴（值轴）标题
    const valAxNode = PPTXUtils.getTextByPathList(plotArea, ["c:valAx"]);
    let yAxisTitle = null;
    if (valAxNode && valAxNode["c:title"]) {
        const yAxisTitleNode = valAxNode["c:title"];
        const yAxisTxNode = yAxisTitleNode["c:tx"];
        
        if (yAxisTxNode) {
            // 从Y轴标题中提取文本，处理多个 a:r 节点的情况
            const yAxisRichNode = PPTXUtils.getTextByPathList(yAxisTxNode, ["c:rich"]);
            if (yAxisRichNode && yAxisRichNode["a:p"]) {
                const yAxisParagraphs = Array.isArray(yAxisRichNode["a:p"]) ? yAxisRichNode["a:p"] : [yAxisRichNode["a:p"]];
                const yAxisTexts = [];
                for (const p of yAxisParagraphs) {
                    if (p && p["a:r"]) {
                        const yAxisTextRuns = Array.isArray(p["a:r"]) ? p["a:r"] : [p["a:r"]];
                        for (const textRun of yAxisTextRuns) {
                            if (textRun && textRun["a:t"]) {
                                yAxisTexts.push(textRun["a:t"]);
                            }
                        }
                    }
                }
                if (yAxisTexts.length > 0) {
                    yAxisTitle = yAxisTexts.join('');
                }
            }
        }
    }
    
    // 提取X轴（分类轴）标题
    const catAxNode = PPTXUtils.getTextByPathList(plotArea, ["c:catAx"]);
    let xAxisTitle = null;
    if (catAxNode && catAxNode["c:title"]) {
        const xAxisTitleNode = catAxNode["c:title"];
        const xAxisTxNode = xAxisTitleNode["c:tx"];
        
        if (xAxisTxNode) {
            // 从X轴标题中提取文本，处理多个 a:r 节点的情况
            const xAxisRichNode = PPTXUtils.getTextByPathList(xAxisTxNode, ["c:rich"]);
            if (xAxisRichNode && xAxisRichNode["a:p"]) {
                const xAxisParagraphs = Array.isArray(xAxisRichNode["a:p"]) ? xAxisRichNode["a:p"] : [xAxisRichNode["a:p"]];
                const xAxisTexts = [];
                for (const p of xAxisParagraphs) {
                    if (p && p["a:r"]) {
                        const xAxisTextRuns = Array.isArray(p["a:r"]) ? p["a:r"] : [p["a:r"]];
                        for (const textRun of xAxisTextRuns) {
                            if (textRun && textRun["a:t"]) {
                                xAxisTexts.push(textRun["a:t"]);
                            }
                        }
                    }
                }
                if (xAxisTexts.length > 0) {
                    xAxisTitle = xAxisTexts.join('');
                }
            }
        }
    }
    
    // 提取图例位置
    let legendPos = "r"; // 默认位置
    const legendNode = PPTXUtils.getTextByPathList(chart, ["c:legend"]);
    if (legendNode && legendNode["c:legendPos"]) {
        legendPos = PPTXUtils.getTextByPathList(legendNode["c:legendPos"], ["attrs", "val"]); 
    }
    
    // 提取图表颜色方案
    const colorSchemes = [];
    
    // 首先尝试从颜色样式文件（colors1.xml）中提取颜色
    if (colorStyleData) {
        // 解析颜色样式文件，提取颜色列表
        const colorStyleRoot = colorStyleData["cs:colorStyle"] || colorStyleData["a:clrScheme"];
        if (colorStyleRoot) {
            // 处理Microsoft Office颜色样式文件格式
            const colorKeys = [
                "a:schemeClr", "a:sysClr", "a:srgbClr"
            ];
            
            for (const colorKey of colorKeys) {
                if (colorStyleRoot[colorKey]) {
                    const colorElements = Array.isArray(colorStyleRoot[colorKey]) 
                        ? colorStyleRoot[colorKey] 
                        : [colorStyleRoot[colorKey]];
                    
                    for (const colorEl of colorElements) {
                        if (colorEl && colorEl["attrs"]) {
                            let colorValue;
                            
                            if (colorEl["attrs"]["val"]) {
                                // 对于 schemeClr，使用主题颜色映射
                                if (colorKey === "a:schemeClr") {
                                    const themeColor = PPTXColorUtils.getSchemeColorFromTheme(`a:${colorEl["attrs"]["val"]}`, undefined, undefined, warpObj);
                                    if (themeColor) {
                                        colorSchemes.push(themeColor);
                                    }
                                } else if (colorKey === "a:srgbClr") {
                                    // 对于 srgbClr，直接使用 val 属性作为颜色值
                                    colorValue = colorEl["attrs"]["val"];
                                    if (colorValue) {
                                        colorSchemes.push(colorValue);
                                    }
                                } else if (colorKey === "a:sysClr") {
                                    // 对于 sysClr，使用 lastClr 属性
                                    colorValue = colorEl["attrs"]["lastClr"] || colorEl["attrs"]["val"];
                                    if (colorValue) {
                                        colorSchemes.push(colorValue);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            // 处理颜色变体
            if (colorStyleRoot["cs:variation"]) {
                const variations = Array.isArray(colorStyleRoot["cs:variation"]) 
                    ? colorStyleRoot["cs:variation"] 
                    : [colorStyleRoot["cs:variation"]];
                
                for (const variation of variations) {
                    if (variation && variation["a:lumMod"]) {
                        const lumMod = variation["a:lumMod"];
                        if (lumMod && lumMod["attrs"] && lumMod["attrs"]["val"]) {
                            // 这里可以根据亮度调整值来调整颜色，暂时跳过
                        }
                    }
                }
            }
        }
    }
    
    // 如果从样式文件中没有获得足够的颜色，或者没有样式文件，则从图表数据中提取
    if (colorSchemes.length === 0) {
        const plotAreaCharts = Object.keys(plotArea).filter(key => key.includes('Chart'));
        for (const chartKey of plotAreaCharts) {
            const chartData = plotArea[chartKey];
            if (chartData && chartData['c:ser']) {
                const seriesArray = Array.isArray(chartData['c:ser']) ? chartData['c:ser'] : [chartData['c:ser']];
                for (const series of seriesArray) {
                    if (series && series['c:spPr']) {
                        // 提取系列颜色
                        const solidFill = series['c:spPr']['a:solidFill'];
                        if (solidFill) {
                            const color = PPTXColorUtils.getSolidFill(solidFill, undefined, undefined, warpObj);
                            if (color) {
                                colorSchemes.push(color);
                            }
                        }
                        // 检查是否有其他填充类型
                        else {
                            const fillTypes = ['a:gradFill', 'a:pattFill', 'a:blipFill'];
                            for (const fillType of fillTypes) {
                                if (series['c:spPr'][fillType]) {
                                    const color = PPTXColorUtils.getSolidFill(series['c:spPr'][fillType], undefined, undefined, warpObj);
                                    if (color) {
                                        colorSchemes.push(color);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    // 同时检查系列本身是否有颜色定义
                    if (series && series['c:spPr'] && series['c:spPr']['a:ln']) {
                        // 检查线条颜色
                        const lineSolidFill = series['c:spPr']['a:ln']['a:solidFill'];
                        if (lineSolidFill) {
                            const color = PPTXColorUtils.getSolidFill(lineSolidFill, undefined, undefined, warpObj);
                            if (color) {
                                colorSchemes.push(color);
                            }
                        }
                    }
                }
            }
        }
    }
    
    // 提取图表样式信息（从style1.xml）
    let chartStyleInfo: any = {};
    if (chartStyleData && chartStyleData["cs:chartStyle"]) {
        // 提取轴标题样式
        const axisTitleStyle = chartStyleData["cs:chartStyle"]["cs:axisTitle"];
        if (axisTitleStyle) {
            chartStyleInfo.axisTitleStyle = axisTitleStyle;
        }
        
        // 提取分类轴样式
        const catAxisStyle = chartStyleData["cs:chartStyle"]["cs:categoryAxis"];
        if (catAxisStyle) {
            chartStyleInfo.catAxisStyle = catAxisStyle;
        }
        
        // 提取值轴样式
        const valAxisStyle = chartStyleData["cs:chartStyle"]["cs:valueAxis"];
        if (valAxisStyle) {
            chartStyleInfo.valAxisStyle = valAxisStyle;
        }
        
        // 提取图例样式
        const legendStyle = chartStyleData["cs:chartStyle"]["cs:legend"];
        if (legendStyle) {
            chartStyleInfo.legendStyle = legendStyle;
        }
        
        // 提取数据标签样式
        const dataLabelStyle = chartStyleData["cs:chartStyle"]["cs:dataLabel"];
        if (dataLabelStyle) {
            chartStyleInfo.dataLabelStyle = dataLabelStyle;
        }
        
        // 提取网格线样式
        const gridlineMajorStyle = chartStyleData["cs:chartStyle"]["cs:gridlineMajor"];
        if (gridlineMajorStyle) {
            chartStyleInfo.gridlineMajorStyle = gridlineMajorStyle;
        }
        
        const gridlineMinorStyle = chartStyleData["cs:chartStyle"]["cs:gridlineMinor"];
        if (gridlineMinorStyle) {
            chartStyleInfo.gridlineMinorStyle = gridlineMinorStyle;
        }
        
        // 提取图表标题样式
        const titleStyle = chartStyleData["cs:chartStyle"]["cs:title"];
        if (titleStyle) {
            chartStyleInfo.titleStyle = titleStyle;
        }
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
                const extractedSeriesData = [];
                // 提取所有系列的数据，而不仅仅是第一个
                for (let k = 0; k < validSeries.length; k++) {
                    const series = validSeries[k];
                    const seriesData = extractChartData(series["c:ser"]);
                    // 添加系列名称
                    const seriesName = extractSeriesName(series);
                    extractedSeriesData.push({
                        data: seriesData,
                        name: seriesName
                    });
                }
                
                const chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": `chart${chartID}`,
                        "chartType": chartType.type,
                        "chartData": extractedSeriesData,
                        "hasMultipleSeries": validSeries.length > 1,
                        "title": chartTitle,
                        "yAxisTitle": yAxisTitle,
                        "xAxisTitle": xAxisTitle,
                        "legendPos": legendPos,
                        "colorSchemes": colorSchemes,
                        "chartStyleInfo": chartStyleInfo
                    }
                };
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
                const extractedSeriesData = [];
                const seriesArray = Array.isArray(plotArea[key]["c:ser"]) ? plotArea[key]["c:ser"] : [plotArea[key]["c:ser"]];
                
                for (let k = 0; k < seriesArray.length; k++) {
                    const series = seriesArray[k];
                    const seriesData = extractChartData(series);
                    const seriesName = extractSeriesName({"c:ser": series});
                    extractedSeriesData.push({
                        data: seriesData,
                        name: seriesName
                    });
                }
                
                const fallbackData = {
                    "type": "createChart",
                    "data": {
                        "chartID": `chart${chartID}`,
                        "chartType": "lineChart",
                        "chartData": extractedSeriesData,
                        "title": chartTitle,
                        "yAxisTitle": yAxisTitle,
                        "xAxisTitle": xAxisTitle,
                        "legendPos": legendPos,
                        "chartStyleInfo": chartStyleInfo
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
// 提取系列名称的辅助函数
function extractSeriesName(seriesNode) {
    if (!seriesNode || !seriesNode["c:ser"]) return `Series ${Math.floor(Math.random() * 1000)}`;
    
    const series = seriesNode["c:ser"];
    const txNode = PPTXUtils.getTextByPathList(series, ["c:tx", "c:strRef", "c:strCache", "c:pt"]);
    
    if (txNode) {
        if (Array.isArray(txNode)) {
            return txNode.map(item => {
                if (item && item["c:v"]) {
                    return item["c:v"];
                } else {
                    // 尝试从attrs中获取val
                    return PPTXUtils.getTextByPathList(item, ["attrs", "val"]) || '';
                }
            }).join(' ');
        } else {
            if (txNode["c:v"]) {
                return txNode["c:v"];
            } else {
                // 尝试从attrs中获取val
                return PPTXUtils.getTextByPathList(txNode, ["attrs", "val"]) || `Series ${Math.floor(Math.random() * 1000)}`;
            }
        }
    }
    
    return `Series ${Math.floor(Math.random() * 1000)}`;
}

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
                    // 尝试获取值，优先从c:v获取，如果没有则从attrs.val获取
                    let value = safeGetPath(pointNode, ["c:v"], null);
                    if (value === null) {
                        value = safeGetPath(pointNode, ["attrs", "val"], null);
                    }
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
                    // 尝试获取值，优先从c:v获取，如果没有则从attrs.val获取
                    let value = safeGetPath(pointNode, ["c:v"], null);
                    if (value === null) {
                        value = safeGetPath(pointNode, ["attrs", "val"], null);
                    }
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
    // 处理第三种图表格式：从XML中直接提取数据点
    try {
        // 提取所有系列数据
        for (let i = 0; i < seriesArray.length; i++) {
            const seriesItem = seriesArray[i];
            if (!seriesItem) continue;
            
            // 提取类别标签（X轴）
            const catNode = safeGetPath(seriesItem, ["c:cat", "c:strRef", "c:strCache"], null) || 
                           safeGetPath(seriesItem, ["c:cat", "c:numRef", "c:numCache"], null);
            
            // 提取值数据（Y轴）
            const valNode = safeGetPath(seriesItem, ["c:val", "c:numRef", "c:numCache"], null);
            
            const categories = [];
            const values = [];
            
            // 提取类别名称
            if (catNode && catNode["c:pt"]) {
                const catPoints = Array.isArray(catNode["c:pt"]) ? catNode["c:pt"] : [catNode["c:pt"]];
                for (let j = 0; j < catPoints.length; j++) {
                    const pt = catPoints[j];
                    if (pt) {
                        // 检查是否有c:v值
                        let value = pt["c:v"];
                        // 如果没有直接的c:v，尝试从attrs获取
                        if (!value) {
                            value = PPTXUtils.getTextByPathList(pt, ["attrs", "val"]);
                        }
                        if (value !== undefined && value !== null) {
                            categories.push(value);
                        }
                    }
                }
            }
            
            // 提取数值
            if (valNode && valNode["c:pt"]) {
                const valPoints = Array.isArray(valNode["c:pt"]) ? valNode["c:pt"] : [valNode["c:pt"]];
                for (let j = 0; j < valPoints.length; j++) {
                    const pt = valPoints[j];
                    if (pt) {
                        // 检查是否有c:v值
                        let value = pt["c:v"];
                        // 如果没有直接的c:v，尝试从attrs获取
                        if (!value) {
                            value = PPTXUtils.getTextByPathList(pt, ["attrs", "val"]);
                        }
                        if (value !== undefined && value !== null) {
                            const numValue = parseFloat(value);
                            if (!isNaN(numValue)) {
                                values.push(numValue);
                            }
                        }
                    }
                }
            }
            
            // 创建数据序列 - 现在将类别标签和值配对为对象格式
            if (categories.length > 0 && values.length > 0) {
                const seriesNameNode = safeGetPath(seriesItem, ["c:tx", "c:strRef", "c:strCache", "c:pt"], null);
                let seriesName;
                if (seriesNameNode) {
                    if (Array.isArray(seriesNameNode)) {
                        seriesName = seriesNameNode.map(item => {
                            if (item && item["c:v"]) {
                                return item["c:v"];
                            } else {
                                // 尝试从attrs中获取val
                                return PPTXUtils.getTextByPathList(item, ["attrs", "val"]) || `Series ${i + 1}-part`;
                            }
                        }).join(' ');
                    } else {
                        if (seriesNameNode["c:v"]) {
                            seriesName = seriesNameNode["c:v"];
                        } else {
                            // 尝试从attrs中获取val
                            seriesName = PPTXUtils.getTextByPathList(seriesNameNode, ["attrs", "val"]) || `Series ${i + 1}`;
                        }
                    }
                } else {
                    seriesName = `Series ${i + 1}`;
                }
                
                // 将类别和值配对成 {x: category, y: value} 格式
                const pairedValues = [];
                for (let k = 0; k < Math.min(categories.length, values.length); k++) {
                    pairedValues.push({
                        x: categories[k],
                        y: values[k]
                    });
                }
                
                dataMat.push({
                    key: seriesName,
                    values: pairedValues,
                    labels: categories
                });
            } else if (values.length > 0) {
                // 如果只有值没有类别标签，使用默认标签
                const seriesNameNode = safeGetPath(seriesItem, ["c:tx", "c:strRef", "c:strCache", "c:pt"], null);
                let seriesName;
                if (seriesNameNode) {
                    if (Array.isArray(seriesNameNode)) {
                        seriesName = seriesNameNode.map(item => {
                            if (item && item["c:v"]) {
                                return item["c:v"];
                            } else {
                                // 尝试从attrs中获取val
                                return PPTXUtils.getTextByPathList(item, ["attrs", "val"]) || `Series ${i + 1}-part`;
                            }
                        }).join(' ');
                    } else {
                        if (seriesNameNode["c:v"]) {
                            seriesName = seriesNameNode["c:v"];
                        } else {
                            // 尝试从attrs中获取val
                            seriesName = PPTXUtils.getTextByPathList(seriesNameNode, ["attrs", "val"]) || `Series ${i + 1}`;
                        }
                    }
                } else {
                    seriesName = `Series ${i + 1}`;
                }
                
                // 使用默认类别名称
                const pairedValues = [];
                for (let k = 0; k < values.length; k++) {
                    pairedValues.push({
                        x: `Item ${k + 1}`,
                        y: values[k]
                    });
                }
                
                dataMat.push({
                    key: seriesName,
                    values: pairedValues,
                    labels: []
                });
            }
        }
        
        if (dataMat.length > 0) {
            return dataMat;
        }
    }
    catch (e) {
        // extraction failed
    }
    return dataMat;
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
function processSingleMsg(d) {
    // 检查外部图表库是否可用
    // @ts-ignore - External libraries nv and d3
    if (typeof (globalThis as any).nv === 'undefined' || typeof (globalThis as any).d3 === 'undefined') {
        console.warn('External chart libraries nv and d3 are not available.');
        return false;
    }
    // @ts-ignore
    let chartID = d.chartID;
    let chartType = d.chartType;
    let chartData = d.chartData;
    let title = d.title;
    let yAxisTitle = d.yAxisTitle;
    let xAxisTitle = d.xAxisTitle;
    let legendPos = d.legendPos;
    let colorSchemes = d.colorSchemes || [];
    let chartStyleInfo = d.chartStyleInfo || {};
    let data = [];
    let chart = null;
    let isDone = false;
    
    // 处理chartData格式，确保是正确的格式
    let processedChartData = [];
    
    if (Array.isArray(chartData) && chartData.length > 0) {
        if (typeof chartData[0].data !== 'undefined') {
            // 如果是新的数据格式（包含名称和数据）
            for (let i = 0; i < chartData.length; i++) {
                const series = chartData[i];
                if (Array.isArray(series.data) && series.data.length > 0 && typeof series.data[0].values !== 'undefined') {
                    // 如果是复杂数据格式
                    for (let j = 0; j < series.data.length; j++) {
                        const subSeries = series.data[j];
                        processedChartData.push({
                            key: subSeries.key || series.name || `Series ${i + 1}-${j + 1}`,
                            values: subSeries.values || subSeries
                        });
                    }
                } else if (series.data && series.data.labels && series.data.values) {
                    // 如果是标签和值的格式
                    const seriesData = [];
                    for (let j = 0; j < Math.min(series.data.labels.length, series.data.values.length); j++) {
                        // nvd3期望值是对象，有x和y属性
                        seriesData.push({
                            x: series.data.labels[j],
                            y: series.data.values[j]
                        });
                    }
                    processedChartData.push({
                        key: series.name || `Series ${i + 1}`,
                        values: seriesData
                    });
                } else {
                    // 简单数据格式 - 需要转换为nvd3期望的格式
                    const seriesData = [];
                    const seriesValues = Array.isArray(series.data) ? series.data : [series.data];
                    
                    // 尝试从series中获取标签，否则使用默认标签
                    const labels = series.data && series.data.labels ? series.data.labels : [];
                    
                    for (let j = 0; j < seriesValues.length; j++) {
                        seriesData.push({
                            x: labels[j] || `Item ${j + 1}`,
                            y: seriesValues[j]
                        });
                    }
                    
                    processedChartData.push({
                        key: series.name || `Series ${i + 1}`,
                        values: seriesData
                    });
                }
            }
        } else {
            // 如果是旧的数据格式 - 需要转换为nvd3期望的格式
            for (let i = 0; i < chartData.length; i++) {
                const item = chartData[i];
                if (item && typeof item === 'object' && item.values) {
                    // 已经是正确格式
                    processedChartData.push(item);
                } else {
                    // 转换为正确格式
                    if (Array.isArray(item)) {
                        const seriesData = [];
                        for (let j = 0; j < item.length; j++) {
                            seriesData.push({
                                x: `Item ${j + 1}`,
                                y: item[j]
                            });
                        }
                        processedChartData.push({
                            key: `Series ${i + 1}`,
                            values: seriesData
                        });
                    } else {
                        // 单个值处理
                        processedChartData.push({
                            key: `Series ${i + 1}`,
                            values: [{x: 'Item 1', y: item}]
                        });
                    }
                }
            }
        }
    } else {
        // 如果是简单的数组格式
        if (Array.isArray(chartData)) {
            for (let i = 0; i < chartData.length; i++) {
                const item = chartData[i];
                if (item && typeof item === 'object' && item.values) {
                    processedChartData.push(item);
                } else {
                    if (Array.isArray(item)) {
                        const seriesData = [];
                        for (let j = 0; j < item.length; j++) {
                            seriesData.push({
                                x: `Item ${j + 1}`,
                                y: item[j]
                            });
                        }
                        processedChartData.push({
                            key: `Series ${i + 1}`,
                            values: seriesData
                        });
                    } else {
                        processedChartData.push({
                            key: `Series ${i + 1}`,
                            values: [{x: 'Item 1', y: item}]
                        });
                    }
                }
            }
        }
    }
    
    switch (chartType) {
        case "lineChart":
            // 确保线图数据格式正确
            data = [];
            if (Array.isArray(processedChartData)) {
                for (let i = 0; i < processedChartData.length; i++) {
                    const series = processedChartData[i];
                    if (series && series.values && Array.isArray(series.values)) {
                        // 确保每个值都是对象格式 {x: ..., y: ...}
                        const formattedValues = series.values.map((val, idx) => {
                            if (typeof val === 'object' && val.x !== undefined && val.y !== undefined) {
                                return val;
                            } else if (typeof val === 'number') {
                                // 如果直接是数字，则转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val};
                            } else {
                                // 其他情况也转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val.y !== undefined ? val.y : val};
                            }
                        });
                        
                        data.push({
                            key: series.key || `Series ${i + 1}`,
                            values: formattedValues
                        });
                    }
                }
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.lineChart()
                .useInteractiveGuideline(true)
                .margin({top: 20})
                .showControls(false)
                .showXAxis(true)
                .showYAxis(true);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有从样式文件中获取到颜色，则使用默认的nvd3颜色方案
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 应用样式信息，如果有的话
            if (chartStyleInfo && chartStyleInfo.lineChartStyle) {
                // 应用线条图表特定的样式
            }
            
            // 设置坐标轴标签
            try {
                // 使用从XML解析出的轴标题，如果不存在则使用默认值
                chart.xAxis.axisLabel(xAxisTitle || 'X轴');
                chart.yAxis.axisLabel(yAxisTitle || 'Y轴');
                
                // 应用轴标题样式
                if (chartStyleInfo.axisTitleStyle) {
                    // 可以根据样式信息进一步定制轴标题
                }
                
                // 应用分类轴样式
                if (chartStyleInfo.catAxisStyle) {
                    // 可以根据样式信息定制分类轴
                }
                
                // 应用值轴样式
                if (chartStyleInfo.valAxisStyle) {
                    // 可以根据样式信息定制值轴
                }
            } catch(e) {
                console.warn('无法设置坐标轴标签:', e);
            }
            break;
        case "barChart":
            // 确保柱状图数据格式正确
            data = [];
            if (Array.isArray(processedChartData)) {
                for (let i = 0; i < processedChartData.length; i++) {
                    const series = processedChartData[i];
                    if (series && series.values && Array.isArray(series.values)) {
                        // 确保每个值都是对象格式 {x: ..., y: ...}
                        const formattedValues = series.values.map((val, idx) => {
                            if (typeof val === 'object' && val.x !== undefined && val.y !== undefined) {
                                return val;
                            } else if (typeof val === 'number') {
                                // 如果直接是数字，则转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val};
                            } else {
                                // 其他情况也转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val.y !== undefined ? val.y : val};
                            }
                        });
                        
                        data.push({
                            key: series.key || `Series ${i + 1}`,
                            values: formattedValues
                        });
                    }
                }
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.multiBarChart()
                .reduceXTicks(false)
                .rotateLabels(-45)
                .margin({top: 20})
                .showControls(false)
                .showXAxis(true)
                .showYAxis(true);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有从样式文件中获取到颜色，则使用默认的nvd3颜色方案
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 应用样式信息，如果有的话
            if (chartStyleInfo && chartStyleInfo.barChartStyle) {
                // 应用柱状图表特定的样式
            }
            
            // 设置坐标轴标签
            try {
                // 使用从XML解析出的轴标题，如果不存在则使用默认值
                chart.xAxis.axisLabel(xAxisTitle || 'X轴');
                chart.yAxis.axisLabel(yAxisTitle || 'Y轴');
                
                // 应用轴标题样式
                if (chartStyleInfo.axisTitleStyle) {
                    // 可以根据样式信息进一步定制轴标题
                }
                
                // 应用分类轴样式
                if (chartStyleInfo.catAxisStyle) {
                    // 可以根据样式信息定制分类轴
                }
                
                // 应用值轴样式
                if (chartStyleInfo.valAxisStyle) {
                    // 可以根据样式信息定制值轴
                }
            } catch(e) {
                console.warn('无法设置坐标轴标签:', e);
            }
            break;
        case "pieChart":
        case "pie3DChart":
            if (processedChartData.length > 0) {
                // 对于饼图，需要特别处理数据格式
                if (processedChartData[0].values) {
                    // 如果是带有标签和值的对象格式
                    data = processedChartData[0].values.map((item, index) => {
                        let label = item.x || processedChartData[0].labels?.[index] || `Item ${index + 1}`;
                        let value = typeof item.y !== 'undefined' ? item.y : item;
                        return { label, value };
                    });
                } else {
                    // 简单数值数组格式
                    data = [];
                    for (let i = 0; i < processedChartData.length; i++) {
                        const series = processedChartData[i];
                        if (series.values) {
                            for (let j = 0; j < series.values.length; j++) {
                                const label = series.labels && series.labels[j] ? series.labels[j] : `Item ${j + 1}`;
                                const value = series.values[j];
                                data.push({ label, value });
                            }
                        } else {
                            data.push({ label: `Item ${i + 1}`, value: Array.isArray(series) ? series[0] : series });
                        }
                    }
                }
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.pieChart()
                .x(function(d) { return d.label })
                .y(function(d) { return d.value })
                .showLabels(true)
                .labelThreshold(.05)
                .labelType('key')
                .margin({top: 20})
                .showControls(false);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有从样式文件中获取到颜色，则使用默认的nvd3颜色方案
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 饼图通常不设置坐标轴，但可以设置标签格式
            try {
                chart.tooltip.contentGenerator(function(obj) {
                    return '<h3>' + obj.data.label + '</h3>' +
                           '<p>' + obj.data.value + '</p>';
                });
            } catch(e) {
                console.warn('无法设置饼图工具提示:', e);
            }
            
            // 饼图不需要坐标轴标签，跳过这部分
            break;
        case "areaChart":
            // 确保区域图数据格式正确
            data = [];
            if (Array.isArray(processedChartData)) {
                for (let i = 0; i < processedChartData.length; i++) {
                    const series = processedChartData[i];
                    if (series && series.values && Array.isArray(series.values)) {
                        // 确保每个值都是对象格式 {x: ..., y: ...}
                        const formattedValues = series.values.map((val, idx) => {
                            if (typeof val === 'object' && val.x !== undefined && val.y !== undefined) {
                                return val;
                            } else if (typeof val === 'number') {
                                // 如果直接是数字，则转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val};
                            } else {
                                // 其他情况也转换为对象格式
                                return {x: `Item ${idx + 1}`, y: val.y !== undefined ? val.y : val};
                            }
                        });
                        
                        data.push({
                            key: series.key || `Series ${i + 1}`,
                            values: formattedValues
                        });
                    }
                }
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.stackedAreaChart()
                .clipEdge(true)
                .useInteractiveGuideline(true)
                .margin({top: 20})
                .showControls(false)
                .showXAxis(true)
                .showYAxis(true);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有从样式文件中获取到颜色，则使用默认的nvd3颜色方案
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 应用样式信息，如果有的话
            if (chartStyleInfo && chartStyleInfo.areaChartStyle) {
                // 应用区域图表特定的样式
            }
            
            // 设置坐标轴标签
            try {
                // 使用从XML解析出的轴标题，如果不存在则使用默认值
                chart.xAxis.axisLabel(xAxisTitle || 'X轴');
                chart.yAxis.axisLabel(yAxisTitle || 'Y轴');
                
                // 应用轴标题样式
                if (chartStyleInfo.axisTitleStyle) {
                    // 可以根据样式信息进一步定制轴标题
                }
                
                // 应用分类轴样式
                if (chartStyleInfo.catAxisStyle) {
                    // 可以根据样式信息定制分类轴
                }
                
                // 应用值轴样式
                if (chartStyleInfo.valAxisStyle) {
                    // 可以根据样式信息定制值轴
                }
            } catch(e) {
                console.warn('无法设置坐标轴标签:', e);
            }
            break;
        case "scatterChart":
            // 确保散点图数据格式正确
            data = [];
            if (Array.isArray(processedChartData)) {
                for (let i = 0; i < processedChartData.length; i++) {
                    const series = processedChartData[i];
                    if (series && series.values && Array.isArray(series.values)) {
                        // 确保每个值都是对象格式 {x: ..., y: ...}
                        const formattedValues = series.values.map((val, idx) => {
                            if (typeof val === 'object' && val.x !== undefined && val.y !== undefined) {
                                return val;
                            } else if (typeof val === 'number') {
                                // 如果直接是数字，则转换为对象格式
                                return {x: idx, y: val};
                            } else {
                                // 其他情况也转换为对象格式
                                return {x: val.x !== undefined ? val.x : idx, y: val.y !== undefined ? val.y : val};
                            }
                        });
                        
                        data.push({
                            key: series.key || `Series ${i + 1}`,
                            values: formattedValues
                        });
                    }
                }
            }
            // @ts-ignore
            chart = (globalThis as any).nv.models.scatterChart()
                .showDistX(true)
                .showDistY(true)
                .margin({top: 20})
                .showControls(false)
                .showXAxis(true)
                .showYAxis(true);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有自定义颜色方案，使用默认的d3颜色比例尺
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 应用样式信息，如果有的话
            if (chartStyleInfo && chartStyleInfo.scatterChartStyle) {
                // 应用散点图表特定的样式
            }
            
            // 设置坐标轴标签
            try {
                // 使用从XML解析出的轴标题，如果不存在则使用默认值
                chart.xAxis.axisLabel(xAxisTitle || 'X轴');
                chart.yAxis.axisLabel(yAxisTitle || 'Y轴');
                
                // 应用轴标题样式
                if (chartStyleInfo.axisTitleStyle) {
                    // 可以根据样式信息进一步定制轴标题
                }
                
                // 应用分类轴样式
                if (chartStyleInfo.catAxisStyle) {
                    // 可以根据样式信息定制分类轴
                }
                
                // 应用值轴样式
                if (chartStyleInfo.valAxisStyle) {
                    // 可以根据样式信息定制值轴
                }
            } catch(e) {
                console.warn('无法设置坐标轴标签:', e);
            }
            break;
        default:
            // 如果未识别的图表类型，默认为多柱状图
            data = processedChartData;
            // @ts-ignore
            chart = (globalThis as any).nv.models.multiBarChart()
                .margin({top: 20})
                .showControls(false)
                .showXAxis(true)
                .showYAxis(true);
            
            // 设置颜色方案，如果有的话
            if (colorSchemes && colorSchemes.length > 0) {
                chart.color(colorSchemes);
            } else {
                // 如果没有从样式文件中获取到颜色，则使用默认的nvd3颜色方案
                chart.color((globalThis as any).d3.scale.category10().range());
            }
            
            // 应用样式信息，如果有的话
            if (chartStyleInfo && chartStyleInfo.defaultChartStyle) {
                // 应用默认图表特定的样式
            }
            
            // 设置坐标轴标签
            try {
                // 使用从XML解析出的轴标题，如果不存在则使用默认值
                chart.xAxis.axisLabel(xAxisTitle || 'X轴');
                chart.yAxis.axisLabel(yAxisTitle || 'Y轴');
                
                // 应用轴标题样式
                if (chartStyleInfo.axisTitleStyle) {
                    // 可以根据样式信息进一步定制轴标题
                }
                
                // 应用分类轴样式
                if (chartStyleInfo.catAxisStyle) {
                    // 可以根据样式信息定制分类轴
                }
                
                // 应用值轴样式
                if (chartStyleInfo.valAxisStyle) {
                    // 可以根据样式信息定制值轴
                }
            } catch(e) {
                console.warn('无法设置坐标轴标签:', e);
            }
            break;
    }
    
    if (chart !== null) {
        const chartElement = document.getElementById(chartID);
        if (chartElement) {
            // 清空图表容器
            chartElement.innerHTML = '';
            
            // 如果有标题，创建一个包含标题的包装元素
            if (title) {
                // 添加标题元素
                const titleDiv = document.createElement('div');
                titleDiv.className = 'chart-title';
                titleDiv.textContent = title;
                titleDiv.style.textAlign = 'center';
                titleDiv.style.fontWeight = 'bold';
                titleDiv.style.marginBottom = '10px';
                titleDiv.style.fontSize = '16px';
                
                // 将标题添加到图表容器中
                chartElement.appendChild(titleDiv);
            }
            
            // 创建SVG元素用于图表
            // @ts-ignore
            (globalThis as any).d3.select(chartElement)
                .append("svg")
                .attr('width', '100%')
                .attr('height', '400px')
                .datum(data)
                .transition().duration(500)
                .call(chart);
            // @ts-ignore
            (globalThis as any).nv.utils.windowResize(chart.update);
            
            // 如果有图例位置信息，调整图例
            if (legendPos) {
                try {
                    // 尝试设置图例位置
                    const svgElement = chartElement.querySelector('svg');
                    if (svgElement) {
                        // 根据图例位置调整布局
                        // 注意：nvd3中图例位置通常是通过CSS或特定的图例组件来控制的
                    }
                } catch (e) {
                    console.warn('无法设置图例位置:', e);
                }
            }
            
            isDone = true;
        }
        else {
            console.warn(`Chart element with id ${chartID} not found.`);
        }
    } else {
        console.warn(`Chart type ${chartType} not supported or data is invalid.`);
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