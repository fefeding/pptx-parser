import { PPTXUtils } from '../utils/utils.js';
import { PPTXShapeFillsUtils } from '../shape/pptx-shape-fills-utils.js';
import { PPTXColorUtils } from '../core/pptx-color-utils.js';
import { PPTXConstants } from '../core/pptx-constants.js';
import { PPTXStyleManager } from '../core/pptx-style-manager.js';
import { PPTXTextElementUtils } from '../text/pptx-text-element-utils.js';

class PPTXTableUtils {
    /**
     * Get table borders style
     * @param {Object} node - The table borders node
     * @param {Object} warpObj - The warp object
     * @returns {String} CSS border style
     */
static getTableBorders(node, warpObj) {
    var borderStyle = "";
    if (node["a:bottom"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:bottom"]["a:ln"]
            }
        }
        var borders = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-bottom");
    }
    if (node["a:top"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:top"]["a:ln"]
            }
        }
        var borders = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-top");
    }
    if (node["a:right"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:right"]["a:ln"]
            }
        }
        var borders = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-right");
    }
    if (node["a:left"] !== undefined) {
        var obj = {
            "p:spPr": {
                "a:ln": node["a:left"]["a:ln"]
            }
        }
        var borders = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
        borderStyle += borders.replace("border", "border-left");
    }

    return borderStyle;
};

    /**
 * Generate internal table HTML
 * @param {Object} node - The table node
 * @param {Object} warpObj - The warp object
 * @param {Object} styleTable - The style table object
 * @returns {String} Table HTML string
 */
static genTableInternal(node, warpObj, styleTable) {
    var order = node["attrs"]["order"];
    var tableNode = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    var xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    
    var getTblPr = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
    var getColsGrid = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
    var tblDir = "";
    if (getTblPr !== undefined) {
        var isRTL = getTblPr["attrs"]["rtl"];
        tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
    }
    
    var firstRowAttr = getTblPr !== undefined ? getTblPr["attrs"]["firstRow"] : undefined;
    var firstColAttr = getTblPr !== undefined ? getTblPr["attrs"]["firstCol"] : undefined;
    var lastRowAttr = getTblPr !== undefined ? getTblPr["attrs"]["lastRow"] : undefined;
    var lastColAttr = getTblPr !== undefined ? getTblPr["attrs"]["lastCol"] : undefined;
    var bandRowAttr = getTblPr !== undefined ? getTblPr["attrs"]["bandRow"] : undefined;
    var bandColAttr = getTblPr !== undefined ? getTblPr["attrs"]["bandCol"] : undefined;
    
    var tblStylAttrObj = {
        isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
        isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
        isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
        isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
        isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
        isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
    };

    var thisTblStyle;
    var tbleStyleId = getTblPr !== undefined ? getTblPr["a:tableStyleId"] : undefined;
    if (tbleStyleId !== undefined) {
        var tbleStylList = PPTXUtils.getTextByPathList(warpObj, ["tableStyles", "a:tblStyleLst", "a:tblStyle"]);
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
        warpObj["thisTbiStyle"] = thisTblStyle;
    }
    
    var tblStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
    var tblBorderStyl = PPTXUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
    var tbl_borders = "";
    if (tblBorderStyl !== undefined) {
        tbl_borders = PPTXTableUtils.getTableBorders(tblBorderStyl, warpObj);
    }
    var tbl_bgcolor = "";
    var tbl_bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
    if (tbl_bgFillschemeClr !== undefined) {
        tbl_bgcolor = PPTXColorUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgFillschemeClr === undefined) {
        tbl_bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
        tbl_bgcolor = PPTXColorUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
    }
    if (tbl_bgcolor !== "") {
        tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
    }
    
    var tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
        PPTXUtils.getPosition(xfrmNode, node, undefined, undefined) +
        PPTXUtils.getSize(xfrmNode, undefined, undefined) +
        " z-index: " + order + ";" +
        tbl_borders + ";" +
        tbl_bgcolor + "'>";

    var trNodes = tableNode["a:tr"];
    if (trNodes.constructor !== Array) {
        trNodes = [trNodes];
    }
    
    var totalrowSpan = 0;
    var rowSpanAry = [];
    for (var i = 0; i < trNodes.length; i++) {
        var rowHeightParam = trNodes[i]["attrs"]["h"];
        var rowHeight = 0;
        var rowsStyl = "";
        if (rowHeightParam !== undefined) {
            rowHeight = parseInt(rowHeightParam) * PPTXConstants.SLIDE_FACTOR;
            rowsStyl += "height:" + rowHeight + "px;";
        }
        var fillColor = "";
        var row_borders = "";
        var fontClrPr = "";
        var fontWeight = "";
        var band_1H_fillColor;
        var band_2H_fillColor;

        if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
            var bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            var rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                var local_fontColor = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontColor !== undefined) {
                    fontClrPr = local_fontColor;
                }

                var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight != "") {
                    fontWeight = local_fontWeight
                }
            }
        }

        if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

            var bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            var borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
            if (borderStyl !== undefined) {
                var local_row_borders = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                if (local_row_borders != "") {
                    row_borders = local_row_borders;
                }
            }
            var rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                var local_fontClrPr = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    fontClrPr = local_fontClrPr;
                }
                var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }

        } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
            fillColor = "";
            row_borders = undefined;
            if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                var bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== "") {
                        fillColor = local_fillColor;
                        band_2H_fillColor = local_fillColor;
                    }
                }

                var borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    var local_row_borders = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                var rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontClrPr = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                }

                var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                if (local_fontWeight !== "") {
                    fontWeight = local_fontWeight;
                }
            }
            if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                var bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                        band_1H_fillColor = local_fillColor;
                    }
                }
                var borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    var local_row_borders = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                var rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    var local_fontClrPr = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }

        }
        if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
            var bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    fillColor = local_fillColor;
                }
            }
            var borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
            if (borderStyl !== undefined) {
                var local_row_borders = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                if (local_row_borders != "") {
                    row_borders = local_row_borders;
                }
            }
            var rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
            if (rowTxtStyl !== undefined) {
                var local_fontClrPr = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    fontClrPr = local_fontClrPr;
                }

                var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
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
                    rowSpanAry = new Array(tcNodes.length).fill(0);
                }
                var totalColSpan = 0;
                while (j < tcNodes.length) {
                    if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                        var a_sorce;
                        if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                            a_sorce = "a:firstCol";
                            if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                PPTXUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                a_sorce = "a:seCell";
                            } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                PPTXUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                a_sorce = "a:neCell";
                            }
                        } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                            !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                            !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                            j != (tcNodes.length - 1)) {

                            if ((j % 2) != 0) {
                                var aBandNode = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                if (aBandNode === undefined) {
                                    aBandNode = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                    if (aBandNode !== undefined) {
                                        a_sorce = "a:band2V";
                                    }
                                } else {
                                    a_sorce = "a:band2V";
                                }
                            }
                        }

                        if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                            a_sorce = "a:lastCol";
                            if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && PPTXUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                a_sorce = "a:swCell";
                            } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && PPTXUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                a_sorce = "a:nwCell";
                            }
                        }

                        var cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj, styleTable)
                        var text = cellParmAry[0];
                        var colStyl = cellParmAry[1];
                        var cssName = cellParmAry[2];
                        var rowSpan = cellParmAry[3];
                        var colSpan = cellParmAry[4];

                        if (rowSpan !== undefined) {
                            totalrowSpan++;
                            rowSpanAry[j] = parseInt(rowSpan) - 1;
                            tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                        } else if (colSpan !== undefined) {
                            tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                            totalColSpan = parseInt(colSpan) - 1;
                        } else {
                            tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
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
                var a_sorce;
                if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                    a_sorce = "a:firstCol";

                } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                    var aBandNode = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                    if (aBandNode === undefined) {
                        aBandNode = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                        if (aBandNode !== undefined) {
                            a_sorce = "a:band2V";
                        }
                    } else {
                        a_sorce = "a:band2V";
                    }
                }

                if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                    a_sorce = "a:lastCol";
                }

                var cellParmAry = PPTXTableUtils.getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj, styleTable)
                var text = cellParmAry[0];
                var colStyl = cellParmAry[1];
                var cssName = cellParmAry[2];
                var rowSpan = cellParmAry[3];

                if (rowSpan !== undefined) {
                    tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                } else {
                    tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                }
            }
        }
        tableHtml += "</tr>";
    }

    return tableHtml;
};

    /**
 * Get table cell parameters
 * @param {Object} tcNodes - The table cell nodes
 * @param {Object} getColsGrid - The columns grid
 * @param {number} row_idx - Row index
 * @param {number} col_idx - Column index
 * @param {Object} thisTblStyle - Table style object
 * @param {string} cellSource - Cell source identifier
 * @param {Object} warpObj - The warp object
 * @param {Object} styleTable - The style table object
 * @returns {Array} [text, colStyl, cssName, rowSpan, colSpan]
 */
static getTableCellParams(tcNodes, getColsGrid, row_idx, col_idx, thisTblStyle, cellSource, warpObj, styleTable) {
    //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
    //text, cell-width, cell-borders, 
    //var text = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
    var rowSpan = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
    var colSpan = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
    var vMerge = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
    var hMerge = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
    var colStyl = "word-wrap: break-word;";
    var colWidth;
    var celFillColor = "";
    var col_borders = "";
    var colFontClrPr = "";
    var colFontWeight = "";
    var lin_bottm = "",
        lin_top = "",
        lin_left = "",
        lin_right = "",
        lin_bottom_left_to_top_right = "",
        lin_top_left_to_bottom_right = "";
    
    var colSapnInt = parseInt(colSpan);
    var total_col_width = 0;
    if (!isNaN(colSapnInt) && colSapnInt > 1){
        for (var k = 0; k < colSapnInt ; k++) {
            total_col_width += parseInt(PPTXUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
        }
    }else{
        total_col_width = PPTXUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
    }
    

    var text = PPTXTextElementUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width, styleTable);//tableStyles

    if (total_col_width != 0 /*&& row_idx == 0*/) {
        colWidth = parseInt(total_col_width) * PPTXConstants.SLIDE_FACTOR;
        colStyl += "width:" + colWidth + "px;";
    }

    //cell bords
    lin_bottm = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
    if (lin_bottm === undefined && cellSource !== undefined) {
        if (cellSource !== undefined)
            lin_bottm = PPTXUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        if (lin_bottm === undefined) {
            lin_bottm = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
        }
    }
    lin_top = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
    if (lin_top === undefined) {
        if (cellSource !== undefined)
            lin_top = PPTXUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        if (lin_top === undefined) {
            lin_top = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
        }
    }
    lin_left = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
    if (lin_left === undefined) {
        if (cellSource !== undefined)
            lin_left = PPTXUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        if (lin_left === undefined) {
            lin_left = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
        }
    }
    lin_right = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
    if (lin_right === undefined) {
        if (cellSource !== undefined)
            lin_right = PPTXUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        if (lin_right === undefined) {
            lin_right = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
        }
    }
    lin_bottom_left_to_top_right = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
    lin_top_left_to_bottom_right = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

    if (lin_bottm !== undefined && lin_bottm != "") {
        var bottom_line_border = PPTXStyleManager.getBorder(lin_bottm, undefined, false, "", warpObj)
        if (bottom_line_border != "") {
            colStyl += "border-bottom:" + bottom_line_border + ";";
        }
    }
    if (lin_top !== undefined && lin_top != "") {
        var top_line_border = PPTXStyleManager.getBorder(lin_top, undefined, false, "", warpObj);
        if (top_line_border != "") {
            colStyl += "border-top: " + top_line_border + ";";
        }
    }
    if (lin_left !== undefined && lin_left != "") {
        var left_line_border = PPTXStyleManager.getBorder(lin_left, undefined, false, "", warpObj)
        if (left_line_border != "") {
            colStyl += "border-left: " + left_line_border + ";";
        }
    }
    if (lin_right !== undefined && lin_right != "") {
        var right_line_border = PPTXStyleManager.getBorder(lin_right, undefined, false, "", warpObj)
        if (right_line_border != "") {
            colStyl += "border-right:" + right_line_border + ";";
        }
    }

    //cell fill color custom
    var getCelFill = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
    if (getCelFill !== undefined && getCelFill != "") {
        var cellObj = {
            "p:spPr": getCelFill
        };
        celFillColor = PPTXShapeFillsUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide")
    }

    //cell fill color theme
    if (celFillColor == "" || celFillColor == "background-color: inherit;") {
        var bgFillschemeClr;
        if (cellSource !== undefined)
            bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
        if (bgFillschemeClr !== undefined) {
            var local_fillColor = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
            if (local_fillColor !== undefined) {
                celFillColor = " background-color: #" + local_fillColor + ";";
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

    //border
    // var borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
    // if (borderStyl !== undefined) {
    //     var local_col_borders = getTableBorders(borderStyl, warpObj);
    //     if (local_col_borders != "") {
    //         col_borders = local_col_borders;
    //     }
    // }
    // if (col_borders != "") {
    //     colStyl += col_borders;
    // }

    //Text style
    var rowTxtStyl;
    if (cellSource !== undefined) {
        rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
    }
    // if (rowTxtStyl === undefined) {
    //     rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
    // }
    if (rowTxtStyl !== undefined) {
        var local_fontClrPr = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
        if (local_fontClrPr !== undefined) {
            colFontClrPr = local_fontClrPr;
        }
        var local_fontWeight = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
        if (local_fontWeight !== "") {
            colFontWeight = local_fontWeight;
        }
    }
    colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
    colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

    return [text, colStyl, cssName, rowSpan, colSpan];
};
}
    // Export to window
// window.PPTXTableUtils = PPTXTableUtils; // Removed for ES modules


export { PPTXTableUtils };