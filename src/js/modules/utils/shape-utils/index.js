/**
 * 形状工具模块主入口
 * 提供统一的接口供外部使用
 */

var PPTXShapeUtils = (function() {
    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;

    /**
     * genShape - 生成形状的主函数
     * @param {Object} node - 节点对象
     * @param {Object} pNode - 父节点对象
     * @param {Object} slideLayoutSpNode - 幻灯片布局节点
     * @param {Object} slideMasterSpNode - 幻灯片母版节点
     * @param {string} id - ID
     * @param {string} name - 名称
     * @param {number} idx - 索引
     * @param {string} type - 类型
     * @param {number} order - 顺序
     * @param {Object} warpObj - 包装对象
     * @param {boolean} isUserDrawnBg - 是否用户绘制背景
     * @param {string} sType - 形状类型
     * @param {string} source - 来源
     * @param {Object} settings - 设置
     * @returns {string} 生成的HTML字符串
     */
    function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source, settings) {
        var result = "";

        if (node["p:spPr"] === undefined) {
            return result;
        }

        var xfrmNode = node["p:spPr"]["a:xfrm"];
        if (xfrmNode === undefined) {
            return result;
        }

        var shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        var custShapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);

        var xfrmList = ["p:spPr", "a:xfrm"];
        var slideXfrmNode = PPTXXmlUtils.getTextByPathList(node, xfrmList);
        var slideLayoutXfrmNode = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
        var slideMasterXfrmNode = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, xfrmList);

        var isFlipV = false;
        var isFlipH = false;
        var flip = "";
        if (slideXfrmNode !== undefined) {
            if (PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1") {
                isFlipV = true;
            }
            if (PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
                isFlipH = true;
            }
        }
        if (isFlipH && !isFlipV) {
            flip = " scale(-1,1)";
        } else if (!isFlipH && isFlipV) {
            flip = " scale(1,-1)";
        } else if (isFlipH && isFlipV) {
            flip = " scale(-1,-1)";
        }

        var offX = 0, offY = 0, extCx = 0, extCy = 0;
        if (xfrmNode["a:off"] !== undefined && xfrmNode["a:off"]["attrs"] !== undefined) {
            offX = parseInt(xfrmNode["a:off"]["attrs"]["x"]) || 0;
            offY = parseInt(xfrmNode["a:off"]["attrs"]["y"]) || 0;
        }
        if (xfrmNode["a:ext"] !== undefined && xfrmNode["a:ext"]["attrs"] !== undefined) {
            extCx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) || 0;
            extCy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) || 0;
        }

        var x = offX * slideFactor;
        var y = offY * slideFactor;
        var w = extCx * slideFactor;
        var h = extCy * slideFactor;

        var shpId = PPTXXmlUtils.getTextByPathList(node, ["attrs", "order"]);
        var fillColor = PPTXStyleUtils.getShapeFill(node, pNode, true, warpObj, source);
        if (fillColor === undefined) {
            fillColor = "#ffffff";
        }
        var grndFillFlg = false;
        var imgFillFlg = false;
        var clrFillType = PPTXStyleUtils.getFillType(PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
        if (clrFillType == "GROUP_FILL") {
            clrFillType = PPTXStyleUtils.getFillType(PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
        }

        var svgGrdnt = "";
        var svgBgImg = "";
        if (clrFillType == "GRADIENT_FILL") {
            grndFillFlg = true;
            var color_arry = (fillColor !== undefined && fillColor.color !== undefined) ? fillColor.color : [];
            var angl = (fillColor !== undefined && fillColor.rot !== undefined) ? fillColor.rot + 90 : 90;
            svgGrdnt = PPTXStyleUtils.getSvgGradient(w, h, angl, color_arry, shpId);
            fillColor = "url(#linGrd_" + shpId + ")";
        } else if (clrFillType == "PIC_FILL") {
            imgFillFlg = true;
            svgBgImg = PPTXStyleUtils.getSvgImagePattern(node, fillColor, shpId, warpObj);
            fillColor = "url(#imgPtrn_" + shpId + ")";
        } else if (clrFillType == "PATTERN_FILL") {
            fillColor = "none";
        }

        var border = PPTXStyleUtils.getBorder(node, pNode, true, "shape", warpObj);
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }

        var rotate = 0;
        if (slideXfrmNode !== undefined) {
            rotate = PPTXXmlUtils.angleToDegrees(PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));
        }
        var txtRotate;
        var txtXframeNode = PPTXXmlUtils.getTextByPathList(node, ["p:txXfrm"]);
        if (txtXframeNode !== undefined) {
            var txtXframeRot = PPTXXmlUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
            if (txtXframeRot !== undefined) {
                txtRotate = PPTXXmlUtils.angleToDegrees(txtXframeRot) + 90;
            }
        } else {
            txtRotate = rotate;
        }

        var outerShdwNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
        var oShadowSvgUrlStr = "";
        if (outerShdwNode !== undefined) {
            var chdwClrNode = PPTXStyleUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
            var outerShdwAttrs = outerShdwNode["attrs"];
            var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
            var dist = parseInt(outerShdwAttrs["dist"]) * slideFactor;
            var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor) : "";
            var vx = dist * Math.sin(dir * Math.PI / 180);
            var hx = dist * Math.cos(dir * Math.PI / 180);
            var oShadowId = "outerhadow_" + shpId;
            oShadowSvgUrlStr = "filter='url(#" + oShadowId + ")'";
        }

        var headEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
        var tailEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);

        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1_val = 0;
        var sAdj2_val = 0;
        var sAdj3_val = 0;
        var sAdj4_val = 0;
        var sAdj5_val = 0;
        var sAdj6_val = 0;
        var sAdj7_val = 0;
        var sAdj8_val = 0;

        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    var sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj1 !== undefined) {
                        sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj2") {
                    var sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj2 !== undefined) {
                        sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj3") {
                    var sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj3 !== undefined) {
                        sAdj3_val = parseInt(sAdj3.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj4") {
                    var sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj4 !== undefined) {
                        sAdj4_val = parseInt(sAdj4.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj5") {
                    var sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj5 !== undefined) {
                        sAdj5_val = parseInt(sAdj5.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj6") {
                    var sAdj6 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj6 !== undefined) {
                        sAdj6_val = parseInt(sAdj6.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj7") {
                    var sAdj7 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj7 !== undefined) {
                        sAdj7_val = parseInt(sAdj7.substr(4)) / 50000;
                    }
                } else if (sAdj_name == "adj8") {
                    var sAdj8 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    if (sAdj8 !== undefined) {
                        sAdj8_val = parseInt(sAdj8.substr(4)) / 50000;
                    }
                }
            }
        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
            var sAdj = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            if (sAdj !== undefined) {
                sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
            }
        }

        result += "<div class='drawing' style='position:absolute; left:" + x + "px; top:" + y + "px; width:" + w + "px; height:" + h + "px; z-index:" + order + ";'>";
        result += "<svg width='" + w + "' height='" + h + "' style='overflow:visible;'>";

        result += '<defs>';

        if (clrFillType == "GRADIENT_FILL") {
            result += svgGrdnt;
        } else if (clrFillType == "PIC_FILL") {
            result += svgBgImg;
        }

        if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
            (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
            var triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
                "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
            result += triangleMarker;
        }

        result += '</defs>';

        if (shapType !== undefined) {
            result += generateShapeByType(shapType, w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, sAdj1_val, sAdj2_val, sAdj3_val, sAdj4_val, sAdj5_val, sAdj6_val, sAdj7_val, sAdj8_val, headEndNodeAttrs, tailEndNodeAttrs);
        } else if (custShapType !== undefined) {
            var pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
            var pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);
            var maxX = parseInt(pathNodes["attrs"]["w"]);
            var maxY = parseInt(pathNodes["attrs"]["h"]);
            var cX = (1 / maxX) * w;
            var cY = (1 / maxY) * h;

            var moveToNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
            var total_shapes = moveToNode.length;

            var lnToNodes = pathNodes["a:lnTo"];
            var cubicBezToNodes = pathNodes["a:cubicBezTo"];
            var arcToNodes = pathNodes["a:arcTo"];
            var closeNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:close"]);

            if (!Array.isArray(moveToNode)) {
                moveToNode = [moveToNode];
            }

            var multiSapeAry = [];
            if (moveToNode.length > 0) {
                Object.keys(moveToNode).forEach(function (key) {
                    var moveToPtNode = moveToNode[key]["a:pt"];
                    if (moveToPtNode !== undefined) {
                        Object.keys(moveToPtNode).forEach(function (key2) {
                            var ptObj = {};
                            var moveToNoPt = moveToPtNode[key2];
                            var spX = moveToNoPt["attrs", "x"];
                            var spY = moveToNoPt["attrs", "y"];
                            var ptOrdr = moveToNoPt["attrs", "order"];
                            ptObj.type = "movto";
                            ptObj.order = ptOrdr;
                            ptObj.x = spX;
                            ptObj.y = spY;
                            multiSapeAry.push(ptObj);
                        });
                    }
                });

                if (lnToNodes !== undefined) {
                    Object.keys(lnToNodes).forEach(function (key) {
                        var lnToPtNode = lnToNodes[key]["a:pt"];
                        if (lnToPtNode !== undefined) {
                            Object.keys(lnToPtNode).forEach(function (key2) {
                                var ptObj = {};
                                var lnToNoPt = lnToPtNode[key2];
                                var ptX = lnToNoPt["attrs", "x"];
                                var ptY = lnToNoPt["attrs", "y"];
                                var ptOrdr = lnToNoPt["attrs", "order"];
                                ptObj.type = "lnto";
                                ptObj.order = ptOrdr;
                                ptObj.x = ptX;
                                ptObj.y = ptY;
                                multiSapeAry.push(ptObj);
                            });
                        }
                    });
                }

                if (cubicBezToNodes !== undefined) {
                    var cubicBezToPtNodesAry = [];
                    if (!Array.isArray(cubicBezToNodes)) {
                        cubicBezToNodes = [cubicBezToNodes];
                    }
                    Object.keys(cubicBezToNodes).forEach(function (key) {
                        cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
                    });

                    cubicBezToPtNodesAry.forEach(function (key2) {
                        var nodeObj = {};
                        nodeObj.type = "cubicBezTo";
                        nodeObj.order = key2[0]["attrs"]["order"];
                        var pts_ary = [];
                        key2.forEach(function (pt) {
                            var pt_obj = {
                                x: pt["attrs"]["x"],
                                y: pt["attrs"]["y"]
                            }
                            pts_ary.push(pt_obj)
                        })
                        nodeObj.cubBzPt = pts_ary;
                        multiSapeAry.push(nodeObj);
                    });
                }

                if (arcToNodes !== undefined) {
                    var arcToNodesAttrs = arcToNodes["attrs"];
                    var arcOrder = arcToNodesAttrs["order"];
                    var hR = arcToNodesAttrs["hR"];
                    var wR = arcToNodesAttrs["wR"];
                    var stAng = arcToNodesAttrs["stAng"];
                    var swAng = arcToNodesAttrs["swAng"];
                    var shftX = 0;
                    var shftY = 0;
                    var arcToPtNode = PPTXXmlUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
                    if (arcToPtNode !== undefined) {
                        shftX = arcToPtNode["x"];
                        shftY = arcToPtNode["y"];
                    }
                    var ptObj = {};
                    ptObj.type = "arcTo";
                    ptObj.order = arcOrder;
                    ptObj.hR = hR;
                    ptObj.wR = wR;
                    ptObj.stAng = stAng;
                    ptObj.swAng = swAng;
                    ptObj.shftX = shftX;
                    ptObj.shftY = shftY;
                    multiSapeAry.push(ptObj);
                }

                if (closeNode !== undefined) {
                    if (!Array.isArray(closeNode)) {
                        closeNode = [closeNode];
                    }
                    Object.keys(closeNode).forEach(function (key) {
                        var clsAttrs = closeNode[key]["attrs"];
                        var clsOrder = clsAttrs["order"];
                        var ptObj = {};
                        ptObj.type = "close";
                        ptObj.order = clsOrder;
                        multiSapeAry.push(ptObj);
                    });
                }

                multiSapeAry.sort(function (a, b) {
                    return a.order - b.order;
                });

                var k = 0;
                var d = "";
                while (k < multiSapeAry.length) {
                    if (multiSapeAry[k].type == "movto") {
                        var spX = parseInt(multiSapeAry[k].x) * cX;
                        var spY = parseInt(multiSapeAry[k].y) * cY;
                        d += " M" + spX + "," + spY;
                    } else if (multiSapeAry[k].type == "lnto") {
                        var Lx = parseInt(multiSapeAry[k].x) * cX;
                        var Ly = parseInt(multiSapeAry[k].y) * cY;
                        d += " L" + Lx + "," + Ly;
                    } else if (multiSapeAry[k].type == "cubicBezTo") {
                        var Cx1 = parseInt(multiSapeAry[k].cubBzPt[0].x) * cX;
                        var Cy1 = parseInt(multiSapeAry[k].cubBzPt[0].y) * cY;
                        var Cx2 = parseInt(multiSapeAry[k].cubBzPt[1].x) * cX;
                        var Cy2 = parseInt(multiSapeAry[k].cubBzPt[1].y) * cY;
                        var Cx3 = parseInt(multiSapeAry[k].cubBzPt[2].x) * cX;
                        var Cy3 = parseInt(multiSapeAry[k].cubBzPt[2].y) * cY;
                        d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
                    } else if (multiSapeAry[k].type == "arcTo") {
                        var hR = parseInt(multiSapeAry[k].hR) * cX;
                        var wR = parseInt(multiSapeAry[k].wR) * cY;
                        var stAng = parseInt(multiSapeAry[k].stAng) / 60000;
                        var swAng = parseInt(multiSapeAry[k].swAng) / 60000;
                        var endAng = stAng + swAng;
                        d += PPTXShapeUtils.shapeArc(wR, hR, wR, hR, stAng, endAng, false);
                    } else if (multiSapeAry[k].type == "quadBezTo") {
                        console.log("custShapType: quadBezTo - TODO");
                    } else if (multiSapeAry[k].type == "close") {
                        d += "z";
                    }
                    k++;
                }

                result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                result += "/>";
            }
        }

        result += "</svg>";

        result += "<div class='block " + PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
            " " + PPTXStyleUtils.getContentDir(node, type, warpObj) +
            "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
            "' style='" +
            PPTXXmlUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
            PPTXXmlUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
            " z-index: " + order + ";" +
            "transform: rotate(" + ((txtRotate !== undefined) ? txtRotate : 0) + "deg)" + flip + ";" +
            "'>";

        if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
            if (type != "diagram" && type != "textBox") {
                type = "shape";
            }
            result += PPTXTextUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
        }
        result += "</div>";
        result += "</div>";

        return result;
    }

    /**
     * 根据形状类型生成形状
     */
    function generateShapeByType(shapType, w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, sAdj1_val, sAdj2_val, sAdj3_val, sAdj4_val, sAdj5_val, sAdj6_val, sAdj7_val, sAdj8_val, headEndNodeAttrs, tailEndNodeAttrs) {
        var result = "";

        switch (shapType) {
            case "rect":
            case "flowChartProcess":
            case "flowChartPredefinedProcess":
            case "flowChartInternalStorage":
            case "actionButtonBlank":
                result += PPTXRectShapes.generateRect(w, h, shapType, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr);
                break;
            case "flowChartCollate":
                result += PPTXRectShapes.generateFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartDocument":
                result += PPTXRectShapes.generateFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartMultidocument":
                result += PPTXRectShapes.generateFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "ellipse":
            case "flowChartConnector":
            case "flowChartSummingJunction":
            case "flowChartOr":
                result += PPTXEllipseShapes.generateEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartTerminator":
                result += PPTXEllipseShapes.generateFlowChartTerminator(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "rtTriangle":
                result += PPTXPolygonShapes.generateRtTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "triangle":
            case "flowChartExtract":
            case "flowChartMerge":
                result += PPTXPolygonShapes.generateTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "diamond":
            case "flowChartDecision":
            case "flowChartSort":
                result += PPTXPolygonShapes.generateDiamond(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "trapezoid":
            case "flowChartManualOperation":
            case "flowChartManualInput":
                result += PPTXPolygonShapes.generateTrapezoid(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "parallelogram":
            case "flowChartInputOutput":
                result += PPTXPolygonShapes.generateParallelogram(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "pentagon":
                result += PPTXPolygonShapes.generatePentagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "hexagon":
            case "flowChartPreparation":
                result += PPTXPolygonShapes.generateHexagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "heptagon":
                result += PPTXPolygonShapes.generateHeptagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "octagon":
                result += PPTXPolygonShapes.generateOctagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "decagon":
                result += PPTXPolygonShapes.generateDecagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "dodecagon":
                result += PPTXPolygonShapes.generateDodecagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star4":
                result += PPTXStarShapes.generateStar4(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star5":
                result += PPTXStarShapes.generateStar5(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star6":
                result += PPTXStarShapes.generateStar6(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star7":
                result += PPTXStarShapes.generateStar7(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star8":
                result += PPTXStarShapes.generateStar8(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star10":
                result += PPTXStarShapes.generateStar10(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star12":
                result += PPTXStarShapes.generateStar12(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star16":
                result += PPTXStarShapes.generateStar16(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star24":
                result += PPTXStarShapes.generateStar24(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "star32":
                result += PPTXStarShapes.generateStar32(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "roundRect":
            case "flowChartAlternateProcess":
                result += PPTXRoundRectShapes.generateRoundRect(w, h, shapType, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "round1Rect":
                result += PPTXRoundRectShapes.generateRound1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "round2DiagRect":
                result += PPTXRoundRectShapes.generateRound2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "round2SameRect":
                result += PPTXRoundRectShapes.generateRound2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "snip1Rect":
            case "flowChartPunchedCard":
                result += PPTXRoundRectShapes.generateSnip1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "snip2DiagRect":
                result += PPTXRoundRectShapes.generateSnip2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "snip2SameRect":
                result += PPTXRoundRectShapes.generateSnip2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "snipRoundRect":
                result += PPTXRoundRectShapes.generateSnipRoundRect(w, h, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "line":
            case "straightConnector1":
            case "bentConnector2":
            case "bentConnector3":
            case "bentConnector4":
            case "bentConnector5":
            case "curvedConnector2":
            case "curvedConnector3":
            case "curvedConnector4":
            case "curvedConnector5": {
                var d = "M 0 0 L " + w + " " + h;
                result += "<path d='" + d + "' stroke='" + border.color +
                    "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                }
                if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                }
                result += "/>";
                break;
            }
            default:
                console.warn("Unsupported shape type: " + shapType);
                break;
        }

        return result;
    }

    return {
        shapeArc: PPTXBaseShapes.shapeArc,
        shapeArcAlt: PPTXBaseShapes.shapeArcAlt,
        shapePie: PPTXBaseShapes.shapePie,
        shapeGear: PPTXBaseShapes.shapeGear,
        shapeSnipRoundRect: PPTXBaseShapes.shapeSnipRoundRect,
        shapeSnipRoundRectAlt: PPTXBaseShapes.shapeSnipRoundRectAlt,
        polarToCartesian: PPTXBaseShapes.polarToCartesian,
        genShape: genShape
    };
})();

window.PPTXShapeUtils = PPTXShapeUtils;