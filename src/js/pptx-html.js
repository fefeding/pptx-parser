/**
 * PPTXHtml - HTML 转换逻辑模块
 * 提取自 pptxjs.js
 */

(function () {
    var $ = window.jQuery;

    // 全局变量引用
    var PPTXUtils = window.PPTXUtils;
    var settings = window.settings; // 将在 pptxjs.js 中设置
    var PPTXParser = window.PPTXParser; // 从 PPTXParser 获取变量

    // 图表 ID 计数器
    var chartID = 0;

    // Helper function: getTextByPathList
    var getTextByPathList = window.PPTXUtils ? window.PPTXUtils.getTextByPathList : function(node, path) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }
        if (node === undefined || node === null) {
            return undefined;
        }
        var l = path.length;
        for (var i = 0; i < l; i++) {
            node = node[path[i]];
            if (node === undefined || node === null) {
                return undefined;
            }
        }
        return node;
    };

    // Helper functions for position and size - use from PPTXUtils
    var getPosition = window.PPTXUtils ? window.PPTXUtils.getPosition : function() { return ""; };
    var getSize = window.PPTXUtils ? window.PPTXUtils.getSize : function() { return ""; };

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
        //cssText += " .slide{margin-bottom: 5px;}\n"; // TODO

        if (settings.slideMode && settings.slideType == "divs2slidesjs") {
            //divId
            //console.log("slideWidth: ", slideWidth)
            cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n"; // TODO
        }
        return cssText;
    }

    // 获取单元格文本（简化版，仅用于表格）
    function getTableCellText(tcNode) {
        if (!tcNode) return "";
        var textBody = tcNode["a:txBody"];
        if (!textBody) return "";
        
        var paragraphs = textBody["a:p"];
        if (!paragraphs) return "";
        
        if (paragraphs.constructor !== Array) {
            paragraphs = [paragraphs];
        }
        
        var cellText = "";
        paragraphs.forEach(function(pNode) {
            var runs = pNode["a:r"];
            if (runs) {
                if (runs.constructor !== Array) {
                    runs = [runs];
                }
                runs.forEach(function(rNode) {
                    var textNode = rNode["a:t"];
                    if (textNode) {
                        var text = textNode["text"] || "";
                        // 处理空白字符
                        text = text.replace(/\s/g, "&nbsp;");
                        cellText += text;
                    }
                });
            }
            // 添加换行
            cellText += "<br/>";
        });
        
        return cellText;
    }

    // 获取填充颜色
    function getSolidFill(fillNode, clrMap, phClr, warpObj) {
        return window.PPTXColorUtils.getSolidFill(fillNode, clrMap, phClr, warpObj);
    }

    // 获取形状填充
    function getShapeFill(node, warpObj) {
        if (!node) return "";
        var fillType = window.PPTXColorUtils.getFillType(node);
        var fillColor;
        
        if (fillType == "NO_FILL") {
            return "";
        } else if (fillType == "SOLID_FILL") {
            var shpFill = node["a:solidFill"];
            fillColor = window.PPTXColorUtils.getSolidFill(shpFill, undefined, undefined, warpObj);
        }
        
        if (fillColor) {
            return "background-color: #" + fillColor + ";";
        }
        return "";
    }

    // 获取单元格参数
    function getTableCellParams(tcNodes, getColsGrid, row_idx, col_idx, thisTblStyle, cellSource, warpObj) {
        var rowSpan = getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
        var colSpan = getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
        var colStyl = "word-wrap: break-word;";
        
        // 计算列宽
        var colSapnInt = parseInt(colSpan);
        var total_col_width = 0;
        if (getColsGrid !== undefined && !isNaN(colSapnInt) && colSapnInt > 1) {
            for (var k = 0; k < colSapnInt; k++) {
                var gridCol = getColsGrid[col_idx + k];
                if (gridCol !== undefined) {
                    var colWidthAttr = getTextByPathList(gridCol, ["attrs", "w"]);
                    if (colWidthAttr !== undefined) {
                        total_col_width += parseInt(colWidthAttr);
                    }
                }
            }
        } else if (getColsGrid !== undefined) {
            var gridCol = (col_idx === undefined) ? getColsGrid : getColsGrid[col_idx];
            if (gridCol !== undefined) {
                total_col_width = getTextByPathList(gridCol, ["attrs", "w"]);
            }
        }
        
        // 获取单元格文本
        var text = getTableCellText(tcNodes);
        
        // 设置列宽
        if (total_col_width != 0) {
            colWidth = parseInt(total_col_width) * slideFactor;
            colStyl += "width:" + colWidth + "px;";
        }
        
        // 单元格边框
        var lin_bottm = getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
        if (lin_bottm === undefined && cellSource !== undefined && thisTblStyle !== undefined) {
            lin_bottm = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            if (lin_bottm === undefined && thisTblStyle !== undefined) {
                lin_bottm = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
            }
        }
        
        var lin_top = getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
        if (lin_top === undefined && cellSource !== undefined && thisTblStyle !== undefined) {
            lin_top = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            if (lin_top === undefined && thisTblStyle !== undefined) {
                lin_top = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
            }
        }
        
        var lin_left = getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
        if (lin_left === undefined && cellSource !== undefined && thisTblStyle !== undefined) {
            lin_left = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            if (lin_left === undefined && thisTblStyle !== undefined) {
                lin_left = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
            }
        }
        
        var lin_right = getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
        if (lin_right === undefined && cellSource !== undefined && thisTblStyle !== undefined) {
            lin_right = getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            if (lin_right === undefined && thisTblStyle !== undefined) {
                lin_right = getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
            }
        }
        
        // 应用边框
        if (lin_bottm !== undefined && lin_bottm != "") {
            var bottom_line_border = PPTXUtils.getBorder(lin_bottm, undefined, false, "", warpObj);
            if (bottom_line_border != "") {
                colStyl += "border-bottom:" + bottom_line_border + ";";
            }
        }
        if (lin_top !== undefined && lin_top != "") {
            var top_line_border = PPTXUtils.getBorder(lin_top, undefined, false, "", warpObj);
            if (top_line_border != "") {
                colStyl += "border-top: " + top_line_border + ";";
            }
        }
        if (lin_left !== undefined && lin_left != "") {
            var left_line_border = PPTXUtils.getBorder(lin_left, undefined, false, "", warpObj);
            if (left_line_border != "") {
                colStyl += "border-left: " + left_line_border + ";";
            }
        }
        if (lin_right !== undefined && lin_right != "") {
            var right_line_border = PPTXUtils.getBorder(lin_right, undefined, false, "", warpObj);
            if (right_line_border != "") {
                colStyl += "border-right:" + right_line_border + ";";
            }
        }
        
        // 单元格填充色
        var celFillColor = "";
        var getCelFill = getTextByPathList(tcNodes, ["a:tcPr"]);
        if (getCelFill !== undefined) {
            var cellObj = { "p:spPr": getCelFill };
            celFillColor = getShapeFill(cellObj, warpObj);
        }
        
        // 单元格填充色（主题）
        if (celFillColor == "" || celFillColor == "background-color: inherit;") {
            if (cellSource !== undefined && thisTblStyle !== undefined) {
                var bgFillschemeClr = getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        celFillColor = " background-color: #" + local_fillColor + ";";
                    }
                }
            }
        }
        
        var cssName = "";
        if (celFillColor !== undefined && celFillColor != "") {
            if (celFillColor in styleTable) {
                cssName = styleTable[celFillColor]["name"];
            } else {
                cssName = "_tbl_cell_css_" + (Object.keys(styleTable).length + 1);
                styleTable[celFillColor] = {
                    "name": cssName,
                    "text": celFillColor
                };
            }
        }
        
        colStyl += celFillColor;
        
        return [text, colStyl, cssName, rowSpan, colSpan];
    }

    // 生成表格 HTML
    function genTable(node, warpObj) {
        var order = node["attrs"]["order"];
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

                            var cellParmAry = getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj);
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
                    var cellParmAry = getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj);
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
        var order = node["attrs"]["order"];
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);

        var readXmlFile = PPTXParser ? PPTXParser.readXmlFile : function() { return null; };

        var result = "<div id='chart" + chartID + "' class='block content' style='" +
            getPosition(xfrmNode, node, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
        var refName = warpObj["slideResObj"][rid]["target"];
        var content = readXmlFile(warpObj["zip"], refName);
        var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

        var chartData = null;
        for (var key in plotArea) {
            switch (key) {
                case "c:lineChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "lineChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:barChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "barChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pieChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "pieChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pie3DChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "pie3DChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:areaChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "areaChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:scatterChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "scatterChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:catAx":
                    break;
                case "c:valAx":
                    break;
                default:
            }
        }

        // Store chart data for later processing
        if (chartData !== null) {
            if (!window.MsgQueue) {
                window.MsgQueue = [];
            }
            window.MsgQueue.push(chartData);
        }

        chartID++;
        return result;
    }

    // 生成图表数据
    function extractChartData(serNode) {
        var dataMat = new Array();

        if (serNode === undefined) {
            return dataMat;
        }

        if (serNode["c:xVal"] !== undefined) {
            var dataRow = new Array();
            var eachElement = function(node, doFunction) {
                if (node === undefined) {
                    return;
                }
                var result = "";
                if (node.constructor === Array) {
                    var l = node.length;
                    for (var i = 0; i < l; i++) {
                        result += doFunction(node[i], i);
                    }
                } else {
                    result += doFunction(node, 0);
                }
                return result;
            };

            eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
            dataRow = new Array();
            eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                dataRow.push(parseFloat(innerNode["c:v"]));
                return "";
            });
            dataMat.push(dataRow);
        } else {
            var eachElement = function(node, doFunction) {
                if (node === undefined) {
                    return;
                }
                var result = "";
                if (node.constructor === Array) {
                    var l = node.length;
                    for (var i = 0; i < l; i++) {
                        result += doFunction(node[i], i);
                    }
                } else {
                    result += doFunction(node, 0);
                }
                return result;
            };

            eachElement(serNode, function (innerNode, index) {
                var dataRow = new Array();
                var colName = getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                // Category (string or number)
                var rowNames = {};
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

    // Convert plain numeric lists to proper HTML numbered lists
    function setNumericBullets(elem) {
        if (PPTXUtils && PPTXUtils.getNumTypeNum) {
            var prgrphs_arry = elem;
            for (var i = 0; i < prgrphs_arry.length; i++) {
                var buSpan = $(prgrphs_arry[i]).find('.numeric-bullet-style');
                if (buSpan.length > 0) {
                    //console.log("DIV-"+i+":");
                    var prevBultTyp = "";
                    var prevBultLvl = "";
                    var buletIndex = 0;
                    var tmpArry = new Array();
                    var tmpArryIndx = 0;
                    var buletTypSrry = new Array();
                    for (var j = 0; j < buSpan.length; j++) {
                        var bult_typ = $(buSpan[j]).data("bulltname");
                        var bult_lvl = $(buSpan[j]).data("bulltlvl");
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
                        $(buSpan[j]).html(numIdx);
                    }
                }
            }
        } else {
            // Fallback to simple list conversion if PPTXUtils is not available
            jqSelector.find('li').each(function () {
                var $li = $(this);
                var html = $li.html();
                // If it starts with a number and a dot, treat as numbered list item
                if (/^\d+\.\s/.test(html)) {
                    // Ensure parent is ol if not already
                    var $parent = $li.parent();
                    if (!$parent.is('ol')) {
                        $parent.each(function () {
                            if (!$(this).is('ol')) {
                                $(this).filter('ul').replaceWith(function () {
                                    return $('<ol></ol>').append($(this).contents());
                                });
                            }
                        });
                    }
                }
            });
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
    function getBackground(warpObj, slideSize, index) {
        var bgResult = "";
        if (warpObj.processFullTheme === true) {
            // 读取 slide 节点中的背景
            var bgNode = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    // 纯色填充
                    var solidFill = getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = PPTXUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgResult = "<div class='slide-background-" + index + "' style='position:absolute;width:" + slideSize.width + "px;height:" + slideSize.height + "px;background-color:" + color + ";'></div>";
                        }
                    }
                    // 图片填充等可在此扩展
                }
            }
        }
        return bgResult;
    }

    // 获取幻灯片背景填充
    function getSlideBackgroundFill(warpObj, index) {
        var bgColor = "";
        if (warpObj.processFullTheme == "colorsAndImageOnly") {
            var bgNode = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    var solidFill = getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = PPTXUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgColor = "background-color:" + color + ";";
                        }
                    }
                }
            }
        }
        return bgColor;
    }

    // 更新加载进度条
    function updateProgressBar(percent) {
        var progressBarElemtnt = $(".slides-loading-progress-bar");
        progressBarElemtnt.width(percent + "%");
        progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
    }

    // 公开 API
    window.PPTXHtml = {
        genGlobalCSS: genGlobalCSS,
        genTable: genTable,
        genChart: genChart,
        setNumericBullets: setNumericBullets,
        processMsgQueue: processMsgQueue,
        processSingleMsg: processSingleMsg,
        getBackground: getBackground,
        getSlideBackgroundFill: getSlideBackgroundFill,
        extractChartData: extractChartData
    };

})();