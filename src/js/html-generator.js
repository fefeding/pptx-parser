import { PPTXUtils } from './utils/utils.js';
import { PPTXColorUtils } from './core/color-utils.js';
import { PPTXTableUtils } from './table/table-utils.js';

    // 全局变量引用
    var settings = window.settings; // 将在 pptxjs.js 中设置
    var PPTXParser = window.PPTXParser; // 从 PPTXParser 获取变量

    // 图表 ID 计数器
    var chartID = 0;

    // Helper function: getTextByPathList
    var getTextByPathList = PPTXUtils.getTextByPathList;

    // Helper functions for position and size - use from PPTXUtils
    var getPosition = window.PPTXUtils ? PPTXUtils.getPosition : function() { return ""; };
    var getSize = window.PPTXUtils ? PPTXUtils.getSize : function() { return ""; };

    // 从 PPTXParser 获取全局变量
    var slideFactor = window.PPTXParser ? window.PPTXParser.slideFactor || (96 / 914400) : (96 / 914400);
    var styleTable = PPTXParser.styleTable || {};
    var tableStyles = PPTXParser.tableStyles || {};
    var defaultTextStyle = PPTXParser.defaultTextStyle || null;

    // 生成全局 CSS
    function genGlobalCSS() {
        var cssText = "";
        // 从 PPTXParser 获取 styleTable
        var styleTable = PPTXParser.styleTable || {};
        var slideWidth = PPTXParser.slideWidth || 960;
        for (var key in styleTable) {
            // 跳过无效的样式(没有 name 的情况)
            if (!styleTable[key] || !styleTable[key]["name"]) {
                continue;
            }
            var tagname = "";
            // if (settings.slideMode && settings.slideType == "revealjs") {
            //     tagname = "section";
            // } else {
            //     tagname = "div";
            // }
            //ADD suffix
            cssText += tagname + " ." + styleTable[key]["name"] +
                ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
                "{" + styleTable[key]["text"] + "}\n"; //section > div
        }
        cssText += " .slide{margin-bottom: 5px;}\n";

        if (settings.slideMode && settings.slideType == "divs2slidesjs") {
            //divId
            //console.log("slideWidth: ", slideWidth)
            cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n";
        }
        return cssText;
    }



    // 获取填充颜色
    function getSolidFill(fillNode, clrMap, phClr, warpObj) {
        return PPTXColorUtils.getSolidFill(fillNode, clrMap, phClr, warpObj);
    }

    // 获取形状填充
    function getShapeFill(node, warpObj) {
        if (!node) return "";
        var fillType = PPTXColorUtils.getFillType(node);
        var fillColor;
        
        if (fillType == "NO_FILL") {
            return "";
        } else if (fillType == "SOLID_FILL") {
            var shpFill = node["a:solidFill"];
            fillColor = PPTXColorUtils.getSolidFill(shpFill, undefined, undefined, warpObj);
        }
        
        if (fillColor) {
            return "background-color: #" + fillColor + ";";
        }
        return "";
    }



    // 生成表格 HTML
    function genTable(node, warpObj) {
        var order = getTextByPathList(node, ["attrs", "order"]) || 0;
        var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        
        if (!tableNode) {
            return "<div class='block table' style='z-index: " + order + ";'>表格</div>";
        }

        var getTblPr = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        var getColsGrid = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
        
        var tblDir = "";
        if (getTblPr !== undefined) {
            var isRTL = getTblPr["attrs"]["rtl"];
            tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
        }
        
        var firstRowAttr = getTblPr["attrs"]["firstRow"];
        var firstColAttr = getTblPr["attrs"]["firstCol"];
        var lastRowAttr = getTblPr["attrs"]["lastRow"];
        var lastColAttr = getTblPr["attrs"]["lastCol"];
        var bandRowAttr = getTblPr["attrs"]["bandRow"];
        var bandColAttr = getTblPr["attrs"]["bandCol"];
        
        var tblStylAttrObj = {
            isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
            isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
            isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
            isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
            isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
            isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
        };

        var thisTblStyle;
        var tbleStyleId = getTblPr["a:tableStyleId"];
        if (tbleStyleId !== undefined && tableStyles) {
            var tbleStylList = getTextByPathList(tableStyles, ["a:tblStyleLst", "a:tblStyle"]);
            if (tbleStylList !== undefined) {
                if (tbleStylList.constructor === Array) {
                    for (var k = 0; k < tbleStylList.length; k++) {
                        if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList[k];
                        }
                    }
                } else {
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
        
        var tblStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
        var tblBorderStyl = getTextByPathList(tblStyl, ["a:tcBdr"]);
        var tbl_borders = "";
        if (tblBorderStyl !== undefined) {
            tbl_borders = PPTXUtils.getTableBorders(tblBorderStyl, warpObj);
        }
        var tbl_bgcolor = "";
        var tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
        if (tbl_bgFillschemeClr !== undefined) {
            tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
        }
        if (tbl_bgFillschemeClr === undefined) {
            tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
        }
        if (tbl_bgcolor !== "") {
            tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
        }
        
        var tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
            getPosition(xfrmNode, node, undefined, undefined) +
            getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";" +
            tbl_borders + ";" +
            tbl_bgcolor + "'>";

        var trNodes = tableNode["a:tr"];
        if (!trNodes) {
            tableHtml += "</table>";
            return tableHtml;
        }
        
        if (trNodes.constructor !== Array) {
            trNodes = [trNodes];
        }

        var rowSpanAry = [];
        var totalColSpan = 0;

        for (var i = 0; i < trNodes.length; i++) {
            var rowHeightParam = trNodes[i]["attrs"]["h"];
            var rowHeight = 0;
            var rowsStyl = "";
            if (rowHeightParam !== undefined) {
                rowHeight = parseInt(rowHeightParam) * slideFactor;
                rowsStyl += "height:" + rowHeight + "px;";
            }
            var fillColor = "";
            var row_borders = "";
            var fontClrPr = "";
            var fontWeight = "";

            if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontColor !== undefined) {
                        fontClrPr = local_fontColor;
                    }
                    var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }

            if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {
                var bgFillschemeClr = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                var borderStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    var local_row_borders = PPTXUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                var rowTxtStyl = getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontClrPr = getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    var local_fontWeight = ((getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }

            rowsStyl += ((row_borders !== undefined) ? row_borders : "");
            rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
            rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
            if (fillColor !== undefined && fillColor != "") {
                rowsStyl += "background-color: #" + fillColor + ";";
            }
            tableHtml += "<tr style='" + rowsStyl + "'>";

            var tcNodes = trNodes[i]["a:tc"];
            if (tcNodes !== undefined) {
                if (tcNodes.constructor === Array) {
                    var j = 0;
                    if (rowSpanAry.length == 0) {
                        rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(function() { return 0; });
                    }
                    var totalColSpan = 0;
                    while (j < tcNodes.length) {
                        if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                            var a_sorce = "";
                            
                            if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                a_sorce = "a:firstCol";
                            } else if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                a_sorce = "a:lastCol";
                            }

                            var cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj, styleTable);
                            var text = cellParmAry[0];
                            var colStyl = cellParmAry[1];
                            var cssName = cellParmAry[2];
                            var rowSpan = cellParmAry[3];
                            var colSpan = cellParmAry[4];

                            if (rowSpan !== undefined) {
                                tableHtml += "<td class='" + cssName + "' rowspan ='" + parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                rowSpanAry[j] = parseInt(rowSpan) - 1;
                            } else if (colSpan !== undefined) {
                                tableHtml += "<td class='" + cssName + "' colspan = '" + parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                totalColSpan = parseInt(colSpan) - 1;
                            } else {
                                tableHtml += "<td class='" + cssName + "' style = '" + colStyl + "'>" + text + "</td>";
                            }
                        } else {
                            if (rowSpanAry[j] != 0) {
                                rowSpanAry[j] -= 1;
                            }
                            if (totalColSpan != 0) {
                                totalColSpan--;
                            }
                        }
                        j++;
                    }
                } else {
                    var a_sorce = "";
                    if (tblStylAttrObj["isFrstColAttr"] == 1) {
                        a_sorce = "a:firstCol";
                    }
                    var cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj, styleTable);
                    var text = cellParmAry[0];
                    var colStyl = cellParmAry[1];
                    var cssName = cellParmAry[2];
                    var rowSpan = cellParmAry[3];

                    if (rowSpan !== undefined) {
                        tableHtml += "<td class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                    } else {
                        tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                    }
                }
            }
            tableHtml += "</tr>";
        }

        tableHtml += "</table>";
        return tableHtml;
    }

    // 生成图表 HTML
    function genChart(node, warpObj) {
        var order = getTextByPathList(node, ["attrs", "order"]) || 0;
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);

        var readXmlFile = PPTXParser ? PPTXParser.readXmlFile : function() { return null; };

        var result = "<div id='chart" + chartID + "' class='block content' style='" +
            getPosition(xfrmNode, node, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
        var refName = warpObj["slideResObj"][rid]["target"];
        
        // 读取图表文件
        var content = readXmlFile(warpObj["zip"], refName);
        if (!content) {
            chartID++;
            return result;
        }
        
        var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);
        if (!plotArea) {
            chartID++;
            return result;
        }

        // 收集所有有效的图表数据
        var chartDatas = [];
        
        // 处理不同类型的图表
        var chartTypes = [
            { key: "c:lineChart", type: "lineChart" },
            { key: "c:barChart", type: "barChart" },
            { key: "c:pieChart", type: "pieChart" },
            { key: "c:pie3DChart", type: "pie3DChart" },
            { key: "c:areaChart", type: "areaChart" },
            { key: "c:scatterChart", type: "scatterChart" }
        ];
        
        for (var i = 0; i < chartTypes.length; i++) {
            var chartType = chartTypes[i];
            var seriesNode = plotArea[chartType.key];
            
            if (seriesNode) {
                // 确保 seriesNode 是数组
                var seriesArray = seriesNode;
                if (seriesArray.constructor !== Array) {
                    seriesArray = [seriesArray];
                }
                
                // 过滤掉空的系列
                var validSeries = seriesArray.filter(function(series) {
                    return series && series["c:ser"];
                });
                
                if (validSeries.length > 0) {
                    var chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": chartType.type,
                            "chartData": extractChartData(validSeries[0]["c:ser"]),
                            "hasMultipleSeries": validSeries.length > 1
                        }
                    };
                    
                    // 如果有多个系列，记录警告但只使用第一个系列
                    if (validSeries.length > 1) {
                        console.warn('Chart has multiple series, using only the first one:', chartType.type);
                    }
                    
                    chartDatas.push(chartData);
                }
            }
        }
        
        // 如果没有找到任何图表数据，尝试更宽松的搜索
        if (chartDatas.length === 0) {
            console.warn('No standard chart types found, trying fallback extraction');
            
            // 查找任何包含 c:ser 的节点
            for (var key in plotArea) {
                if (key.indexOf('Chart') > -1 && plotArea[key]["c:ser"]) {
                    var fallbackData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "lineChart", // 默认使用折线图
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
            if (!window.MsgQueue) {
                window.MsgQueue = [];
            }
            
            // 将所有图表数据添加到队列
            for (var j = 0; j < chartDatas.length; j++) {
                window.MsgQueue.push(chartDatas[j]);
            }
        }

        chartID++;
        return result;
    }

    // 生成图表数据 - 增强版本，更好的错误处理和数据验证
    function extractChartData(serNode) {
        var dataMat = new Array();

        // 输入验证
        if (serNode === undefined || serNode === null) {
            console.warn('extractChartData: serNode is undefined or null');
            return dataMat;
        }
        
        // 确保 serNode 是数组
        var seriesArray = serNode;
        if (seriesArray.constructor !== Array) {
            seriesArray = [seriesArray];
        }
        
        if (seriesArray.length === 0) {
            console.warn('extractChartData: serNode array is empty');
            return dataMat;
        }

        // 辅助函数：安全获取路径值
        var safeGetPath = function(obj, path, defaultValue) {
            try {
                var result = obj;
                for (var i = 0; i < path.length; i++) {
                    if (result === undefined || result === null) return defaultValue;
                    result = result[path[i]];
                }
                return result !== undefined ? result : defaultValue;
            } catch (e) {
                return defaultValue;
            }
        };
        
        // 辅助函数：安全遍历节点数组
        var safeEachElement = function(nodes, processFunc) {
            if (!nodes || nodes.constructor !== Array) {
                if (nodes) {
                    return processFunc(nodes, 0);
                }
                return '';
            }
            
            var result = '';
            for (var i = 0; i < nodes.length; i++) {
                if (nodes[i]) {
                    result += processFunc(nodes[i], i);
                }
            }
            return result;
        };

        // 处理第一种图表格式：有 c:xVal 和 c:yVal 的简单格式
        var xValNode = safeGetPath(seriesArray[0], ["c:xVal"], null);
        var yValNode = safeGetPath(seriesArray[0], ["c:yVal"], null);
        
        if (xValNode && yValNode) {
            try {
                // 处理 X 值
                var xCache = safeGetPath(xValNode, ["c:numRef", "c:numCache", "c:pt"], null);
                if (xCache) {
                    var xDataRow = [];
                    safeEachElement(xCache, function(pointNode) {
                        var value = safeGetPath(pointNode, ["c:v"], null);
                        if (value !== null) {
                            var numValue = parseFloat(value);
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
                var yCache = safeGetPath(yValNode, ["c:numRef", "c:numCache", "c:pt"], null);
                if (yCache) {
                    var yDataRow = [];
                    safeEachElement(yCache, function(pointNode) {
                        var value = safeGetPath(pointNode, ["c:v"], null);
                        if (value !== null) {
                            var numValue = parseFloat(value);
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
                    console.log('Successfully extracted simple chart data:', dataMat.length, 'rows');
                    return dataMat;
                }
            } catch (e) {
                console.warn('Error extracting simple chart data:', e);
            }
        }

        // 处理第二种图表格式：复杂的多系列格式
        try {
            safeEachElement(seriesArray, function(seriesItem) {
                if (!seriesItem) return '';
                
                var dataRow = [];
                var colName = safeGetPath(seriesItem, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"], null);
                if (colName === null) {
                    // 如果没有名称，使用索引
                    var seriesIndex = seriesArray.indexOf(seriesItem);
                    colName = 'Series ' + (seriesIndex + 1);
                }

                // 提取类别标签
                var rowNames = {};
                var catStrRef = safeGetPath(seriesItem, ["c:cat", "c:strRef", "c:strCache", "c:pt"], null);
                var catNumRef = safeGetPath(seriesItem, ["c:cat", "c:numRef", "c:numCache", "c:pt"], null);
                
                var catPoints = catStrRef || catNumRef;
                if (catPoints) {
                    safeEachElement(catPoints, function(pointNode) {
                        var idx = safeGetPath(pointNode, ["attrs", "idx"], null);
                        var val = safeGetPath(pointNode, ["c:v"], null);
                        if (idx !== null && val !== null) {
                            rowNames[idx] = val;
                        }
                    });
                }

                // 提取值数据
                var valNode = safeGetPath(seriesItem, ["c:val", "c:numRef", "c:numCache", "c:pt"], null);
                if (valNode) {
                    safeEachElement(valNode, function(pointNode) {
                        var idx = safeGetPath(pointNode, ["attrs", "idx"], null);
                        var val = safeGetPath(pointNode, ["c:v"], null);
                        if (idx !== null && val !== null) {
                            var numValue = parseFloat(val);
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
                console.log('Successfully extracted complex chart data:', dataMat.length, 'series');
                return dataMat;
            }
        } catch (e) {
            console.warn('Error extracting complex chart data:', e);
        }
        
        // 如果所有方法都失败，记录错误并返回空数组
        console.error('Failed to extract chart data from series:', seriesArray);
        return [];
    }

    // Convert plain numeric lists to proper HTML numbered lists
    function setNumericBullets(elem) {
        if (PPTXUtils && PPTXUtils.getNumTypeNum) {
            var prgrphs_arry = elem;
            for (var i = 0; i < prgrphs_arry.length; i++) {
                var element = prgrphs_arry[i];
                var buSpan = element.querySelectorAll('.numeric-bullet-style');
                if (buSpan.length > 0) {
                    //console.log("DIV-"+i+":");
                    var prevBultTyp = "";
                    var prevBultLvl = "";
                    var buletIndex = 0;
                    var tmpArry = new Array();
                    var tmpArryIndx = 0;
                    var buletTypSrry = new Array();
                    for (var j = 0; j < buSpan.length; j++) {
                        var span = buSpan[j];
                        var bult_typ = span.dataset.bulltname;
                        var bult_lvl = span.dataset.bulltlvl;
                        //console.log(j+" - "+bult_typ+" lvl: "+bult_lvl );
                        if (buletIndex == 0) {
                            prevBultTyp = bult_typ;
                            prevBultLvl = bult_lvl;
                            tmpArry[tmpArryIndx] = buletIndex;
                            buletTypSrry[tmpArryIndx] = bult_typ;
                            buletIndex++;
                        } else {
                            if (bult_typ == prevBultTyp && bult_lvl == prevBultLvl) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                buletIndex++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                            } else if (bult_typ != prevBultTyp && bult_lvl == prevBultLvl) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            } else if (bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx++;
                                tmpArry[tmpArryIndx] = buletIndex;
                                buletTypSrry[tmpArryIndx] = bult_typ;
                                buletIndex = 1;
                            } else if (bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)) {
                                prevBultTyp = bult_typ;
                                prevBultLvl = bult_lvl;
                                tmpArryIndx--;
                                buletIndex = tmpArry[tmpArryIndx] + 1;
                            }
                        }
                        //console.log(buletTypSrry[tmpArryIndx]+" - "+buletIndex);
                        var numIdx = PPTXUtils.getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                        span.innerHTML = numIdx;
                    }
                }
            }
        } else {
            // Fallback to simple list conversion if PPTXUtils is not available
            // Convert elem to a NodeList or array if needed
            var elements = elem;
            if (!elements.length && elements.nodeType === 1) {
                elements = [elements];
            }
            for (var i = 0; i < elements.length; i++) {
                var element = elements[i];
                var lis = element.querySelectorAll('li');
                for (var j = 0; j < lis.length; j++) {
                    var li = lis[j];
                    var html = li.innerHTML;
                    // If it starts with a number and a dot, treat as numbered list item
                    if (/^\d+\.\s/.test(html)) {
                        // Ensure parent is ol if not already
                        var parent = li.parentNode;
                        if (parent && parent.tagName !== 'OL') {
                            if (parent.tagName === 'UL') {
                                var ol = document.createElement('ol');
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

    // Process message queue and update UI accordingly
    function processMsgQueue(msgQueue) {
        if (!msgQueue || msgQueue.length === 0) return;

        // Process each message
        for (var i = 0; i < msgQueue.length; i++) {
            var msg = msgQueue[i];
            if (msg && msg.type === "createChart" && msg.data) {
                processSingleMsg(msg.data);
            } else {
                console.log("PPTXjs Message:", msg);
            }
        }
        // Clear after processing
        msgQueue.length = 0;
    }

    // 处理单个消息
    function processSingleMsg(d) {
        var chartID = d.chartID;
        var chartType = d.chartType;
        var chartData = d.chartData;

        var data = [];
        var chart = null;
        var isDone = false;

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
                for (var i = 0; i < chartData.length; i++) {
                    var arr = [];
                    for (var j = 0; j < chartData[i].length; j++) {
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

        return isDone;
    }

    // 获取背景
    // getBackground 和 getSlideBackgroundFill 已移至 PPTXBackgroundUtils 模块

    // 更新加载进度条
    function updateProgressBar(percent) {
        var progressBarElemtnt = document.querySelector(".slides-loading-progress-bar");
        if (progressBarElemtnt) {
            progressBarElemtnt.style.width = percent + "%";
            progressBarElemtnt.innerHTML = "<span style='text-align: center;'>Loading...(" + percent + "%)</span>";
        }
    }

    // 公开 API
    const PPTXHtml = {
        genGlobalCSS: genGlobalCSS,
        genTable: genTable,
        genChart: genChart,
        setNumericBullets: setNumericBullets,
        processMsgQueue: processMsgQueue,
        processSingleMsg: processSingleMsg,
        extractChartData: extractChartData
    };


export { PPTXHtml };

// Also export to global scope for backward compatibility
window.PPTXHtml = PPTXHtml;
