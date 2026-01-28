import { PPTXUtils } from '../core/utils.js';
import { PPTXShapeFillsUtils } from '../shape/fills.js';
import { PPTXColorUtils } from '../core/color.js';
import { PPTXConstants } from '../core/constants.js';
import { PPTXStyleManager } from '../core/style-manager.js';
import { PPTXTextElementUtils } from '../text/element.js';

class PPTXTableUtils {
    /**
     * Get table borders style
     * @param {Object} node - The table borders node
     * @param {Object} warpObj - The warp object
     * @returns {String} CSS border style
     */
    static getTableBorders(node: any, warpObj: any): string {
        let borderStyle: string = "";
        if (node["a:bottom"] !== undefined) {
            const obj = {
                "p:spPr": {
                    "a:ln": node["a:bottom"]["a:ln"]
                }
            };
            const borders: string = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-bottom");
        }
        if (node["a:top"] !== undefined) {
            const obj = {
                "p:spPr": {
                    "a:ln": node["a:top"]["a:ln"]
                }
            };
            const borders: string = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-top");
        }
        if (node["a:right"] !== undefined) {
            const obj = {
                "p:spPr": {
                    "a:ln": node["a:right"]["a:ln"]
                }
            };
            const borders: string = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-right");
        }
        if (node["a:left"] !== undefined) {
            const obj = {
                "p:spPr": {
                    "a:ln": node["a:left"]["a:ln"]
                }
            };
            const borders: string = PPTXStyleManager.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-left");
        }

        return borderStyle;
    }

    /**
     * Generate internal table HTML
     * @param {Object} node - The table node
     * @param {Object} warpObj - The warp object
     * @param {Object} styleTable - The style table object
     * @returns {String} Table HTML string
     */
    static async genTableInternal(node: any, warpObj: any, styleTable: any): Promise<string> {
        const order: any = node["attrs"]["order"];
        const tableNode: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        const xfrmNode: any = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);

        const getTblPr: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        const getColsGrid: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
        let tblDir: string = "";
        if (getTblPr !== undefined) {
            const isRTL: any = getTblPr["attrs"]["rtl"];
            tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
        }

        const firstRowAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["firstRow"] : undefined;
        const firstColAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["firstCol"] : undefined;
        const lastRowAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["lastRow"] : undefined;
        const lastColAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["lastCol"] : undefined;
        const bandRowAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["bandRow"] : undefined;
        const bandColAttr: any = getTblPr !== undefined ? getTblPr["attrs"]["bandCol"] : undefined;

        const tblStylAttrObj: any = {
            isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
            isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
            isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
            isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
            isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
            isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
        };

        let thisTblStyle: any;
        const tbleStyleId: any = getTblPr !== undefined ? getTblPr["a:tableStyleId"] : undefined;
        if (tbleStyleId !== undefined) {
            const tbleStylList: any = PPTXUtils.getTextByPathList(warpObj, ["tableStyles", "a:tblStyleLst", "a:tblStyle"]);
            if (tbleStylList !== undefined) {
                if (tbleStylList.constructor === Array) {
                    for (let k = 0; k < tbleStylList.length; k++) {
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

        const tblStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
        const tblBorderStyl: any = PPTXUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
        let tbl_borders: string = "";
        if (tblBorderStyl !== undefined) {
            tbl_borders = PPTXTableUtils.getTableBorders(tblBorderStyl, warpObj);
        }
        let tbl_bgcolor: string = "";
        let tbl_bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
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

        let tableHtml: string = "<table " + tblDir + " style='border-collapse: collapse;" +
            PPTXUtils.getPosition(xfrmNode, node, undefined, undefined, undefined) +
            PPTXUtils.getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";" +
            tbl_borders + ";" +
            tbl_bgcolor + "'>";

        let trNodes: any = tableNode["a:tr"];
        if (trNodes.constructor !== Array) {
            trNodes = [trNodes];
        }

        let totalrowSpan: number = 0;
        let rowSpanAry: number[] = [];
        for (let i = 0; i < trNodes.length; i++) {
            const rowHeightParam: any = trNodes[i]["attrs"]["h"];
            let rowHeight: number = 0;
            let rowsStyl: string = "";
            if (rowHeightParam !== undefined) {
                rowHeight = parseInt(rowHeightParam) * PPTXConstants.SLIDE_FACTOR;
                rowsStyl += "height:" + rowHeight + "px;";
            }
            let fillColor: string = "";
            let row_borders: string = "";
            let fontClrPr: string = "";
            let fontWeight: string = "";
            let band_1H_fillColor: string;
            let band_2H_fillColor: string;

            if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                const bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                const rowTxtStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    const local_fontColor: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontColor !== undefined) {
                        fontClrPr = local_fontColor;
                    }

                    const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight != "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }

            if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                const bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                const borderStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    const local_row_borders: string = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                const rowTxtStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    const local_fontClrPr: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }
                    const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }

            } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                fillColor = "";
                row_borders = "";
                if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                    const bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== "") {
                            fillColor = local_fillColor;
                            band_2H_fillColor = local_fillColor;
                        }
                    }

                    const borderStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        const local_row_borders: string = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    const rowTxtStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        const local_fontClrPr: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                    }

                    const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
                if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                    const bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                    if (bgFillschemeClr !== undefined) {
                        const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                        if (local_fillColor !== undefined) {
                            fillColor = local_fillColor;
                            band_1H_fillColor = local_fillColor;
                        }
                    }
                    const borderStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                    if (borderStyl !== undefined) {
                        const local_row_borders: string = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                        if (local_row_borders != "") {
                            row_borders = local_row_borders;
                        }
                    }
                    const rowTxtStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                    if (rowTxtStyl !== undefined) {
                        const local_fontClrPr: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                        if (local_fontClrPr !== undefined) {
                            fontClrPr = local_fontClrPr;
                        }
                        const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                        if (local_fontWeight != "") {
                            fontWeight = local_fontWeight;
                        }
                    }
                }

            }
            if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                const bgFillschemeClr: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        fillColor = local_fillColor;
                    }
                }
                const borderStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                if (borderStyl !== undefined) {
                    const local_row_borders: string = PPTXTableUtils.getTableBorders(borderStyl, warpObj);
                    if (local_row_borders != "") {
                        row_borders = local_row_borders;
                    }
                }
                const rowTxtStyl: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                if (rowTxtStyl !== undefined) {
                    const local_fontClrPr: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                    if (local_fontClrPr !== undefined) {
                        fontClrPr = local_fontClrPr;
                    }

                    const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                    if (local_fontWeight !== "") {
                        fontWeight = local_fontWeight;
                    }
                }
            }
            rowsStyl += ((row_borders !== "") ? row_borders : "");
            rowsStyl += ((fontClrPr !== "") ? " color: #" + fontClrPr + ";" : "");
            rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
            if (fillColor !== "" && fillColor !== undefined) {
                rowsStyl += "background-color: #" + fillColor + ";";
            }
            tableHtml += "<tr style='" + rowsStyl + "'>";

            const tcNodes: any = trNodes[i]["a:tc"];
            if (tcNodes !== undefined) {
                if (tcNodes.constructor === Array) {
                    let j: number = 0;
                    if (rowSpanAry.length == 0) {
                        const tempAry: number[] = new Array(tcNodes.length).fill(0);
                        rowSpanAry = tempAry;
                    }
                    let totalColSpan: number = 0;
                    while (j < tcNodes.length) {
                        if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                            let a_sorce: string;
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
                                    let aBandNode: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
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

                            const cellParmAry: any = PPTXTableUtils.getTableCellParams(tcNodes[j], getColsGrid, i, j, thisTblStyle, a_sorce, warpObj, styleTable);
                            const text: string = cellParmAry[0];
                            const colStyl: string = cellParmAry[1];
                            const cssName: string = cellParmAry[2];
                            const rowSpan: any = cellParmAry[3];
                            const colSpan: any = cellParmAry[4];

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
                    let a_sorce: string;
                    if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                        a_sorce = "a:firstCol";

                    } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                        let aBandNode: any = PPTXUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
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

                    const cellParmAry: any = PPTXTableUtils.getTableCellParams(tcNodes, getColsGrid, i, undefined, thisTblStyle, a_sorce, warpObj, styleTable);
                    const text: string = cellParmAry[0];
                    const colStyl: string = cellParmAry[1];
                    const cssName: string = cellParmAry[2];
                    const rowSpan: any = cellParmAry[3];

                    if (rowSpan !== undefined) {
                        tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
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
    static async getTableCellParams(tcNodes: any, getColsGrid: any, row_idx: number, col_idx: number | undefined, thisTblStyle: any, cellSource: string | undefined, warpObj: any, styleTable: any): Promise<any[]> {
        //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
        //text, cell-width, cell-borders,
        //const text = genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
        const rowSpan: any = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
        const colSpan: any = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
        const vMerge: any = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
        const hMerge: any = PPTXUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
        let colStyl: string = "word-wrap: break-word;";
        let colWidth: number;
        let celFillColor: string = "";
        const col_borders: string = "";
        const colFontClrPr: string = "";
        const colFontWeight: string = "";
        let lin_bottm: any,
            lin_top: any,
            lin_left: any,
            lin_right: any,
            lin_bottom_left_to_top_right: any,
            lin_top_left_to_bottom_right: any;

        const colSapnInt: number = parseInt(colSpan);
        let total_col_width: any = 0;
        if (!isNaN(colSapnInt) && colSapnInt > 1) {
            for (let k = 0; k < colSapnInt; k++) {
                total_col_width += parseInt(PPTXUtils.getTextByPathList(getColsGrid[col_idx! + k], ["attrs", "w"]));
            }
        } else {
            total_col_width = PPTXUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
        }


        const text: string = await PPTXTextElementUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width, styleTable);//tableStyles

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
            const bottom_line_border: string = PPTXStyleManager.getBorder(lin_bottm, undefined, false, "", warpObj);
            if (bottom_line_border != "") {
                colStyl += "border-bottom:" + bottom_line_border + ";";
            }
        }
        if (lin_top !== undefined && lin_top != "") {
            const top_line_border: string = PPTXStyleManager.getBorder(lin_top, undefined, false, "", warpObj);
            if (top_line_border != "") {
                colStyl += "border-top: " + top_line_border + ";";
            }
        }
        if (lin_left !== undefined && lin_left != "") {
            const left_line_border: string = PPTXStyleManager.getBorder(lin_left, undefined, false, "", warpObj);
            if (left_line_border != "") {
                colStyl += "border-left: " + left_line_border + ";";
            }
        }
        if (lin_right !== undefined && lin_right != "") {
            const right_line_border: string = PPTXStyleManager.getBorder(lin_right, undefined, false, "", warpObj);
            if (right_line_border != "") {
                colStyl += "border-right:" + right_line_border + ";";
            }
        }

        //cell fill color custom
        const getCelFill: any = PPTXUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
        if (getCelFill !== undefined && getCelFill != "") {
            const cellObj = {
                "p:spPr": getCelFill
            };
            celFillColor = PPTXShapeFillsUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide");
        }

        //cell fill color theme
        if (celFillColor == "" || celFillColor == "background-color: inherit;") {
            let bgFillschemeClr: any;
            if (cellSource !== undefined)
                bgFillschemeClr = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
            if (bgFillschemeClr !== undefined) {
                const local_fillColor: string = PPTXColorUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                if (local_fillColor !== undefined) {
                    celFillColor = " background-color: #" + local_fillColor + ";";
                }
            }
        }
        let cssName: string = "";
        if (celFillColor !== "" && celFillColor !== undefined) {
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
        // const borderStyl = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
        // if (borderStyl !== undefined) {
        //     const local_col_borders = getTableBorders(borderStyl, warpObj);
        //     if (local_col_borders != "") {
        //         col_borders = local_col_borders;
        //     }
        // }
        // if (col_borders != "") {
        //     colStyl += col_borders;
        // }

        //Text style
        let rowTxtStyl: any;
        if (cellSource !== undefined) {
            rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
        }
        // if (rowTxtStyl === undefined) {
        //     rowTxtStyl = PPTXUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
        // }
        if (rowTxtStyl !== undefined) {
            const local_fontClrPr: string = PPTXColorUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
            if (local_fontClrPr !== undefined) {
                // colFontClrPr = local_fontClrPr;
                colStyl += "color: #" + local_fontClrPr + ";";
            }
            const local_fontWeight: string = ((PPTXUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
            if (local_fontWeight !== "") {
                // colFontWeight = local_fontWeight;
                colStyl += " font-weight:" + local_fontWeight + ";";
            }
        }
        // colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
        // colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

        return [text, colStyl, cssName, rowSpan, colSpan];
    }
}

export { PPTXTableUtils };