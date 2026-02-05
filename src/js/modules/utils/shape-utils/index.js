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

        var svgCssName = "_svg_css_" + (Object.keys(warpObj.styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
        var effectsClassName = svgCssName + "_effects";
        result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "'" +
            "' style='" +
            PPTXXmlUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
            PPTXXmlUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
            " z-index: " + order + ";" +
            "transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + ";" +
            "'>";

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
            result += generateShapeByType(shapType, w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, sAdj1_val, sAdj2_val, sAdj3_val, sAdj4_val, sAdj5_val, sAdj6_val, sAdj7_val, sAdj8_val, headEndNodeAttrs, tailEndNodeAttrs, node);
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

        return result;
    }

    /**
     * 根据形状类型生成形状
     */
    function generateShapeByType(shapType, w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, sAdj1_val, sAdj2_val, sAdj3_val, sAdj4_val, sAdj5_val, sAdj6_val, sAdj7_val, sAdj8_val, headEndNodeAttrs, tailEndNodeAttrs, node) {
        var result = "";

        switch (shapType) {
            case "rect":
            case "flowChartProcess":
            case "flowChartPredefinedProcess":
            case "flowChartInternalStorage":
            case "actionButtonBlank":
                result += PPTXRectShapes.generateRect(w, h, shapType, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr);
                break;
            case "actionButtonBackPrevious": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    "M" + g11 + "," + vc +
                    " L" + g12 + "," + g9 +
                    " L" + g12 + "," + g10 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonBeginning": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 / 8;
                g15 = g13 / 4;
                g16 = g11 + g14;
                g17 = g11 + g15;
                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    "M" + g17 + "," + vc +
                    " L" + g12 + "," + g9 +
                    " L" + g12 + "," + g10 +
                    " z" +
                    "M" + g16 + "," + g9 +
                    " L" + g11 + "," + g9 +
                    " L" + g11 + "," + g10 +
                    " L" + g16 + "," + g10 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonDocument": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, dx1, g11, g12, g13, g14, g15;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                dx1 = ss * 9 / 32;
                g11 = hc - dx1;
                g12 = hc + dx1;
                g13 = ss * 3 / 16;
                g14 = g12 - g13;
                g15 = g9 + g13;
                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    "M" + g11 + "," + g9 +
                    " L" + g14 + "," + g9 +
                    " L" + g12 + "," + g15 +
                    " L" + g12 + "," + g10 +
                    " L" + g11 + "," + g10 +
                    " z" +
                    "M" + g14 + "," + g9 +
                    " L" + g14 + "," + g15 +
                    " L" + g12 + "," + g15 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonEnd": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 * 3 / 4;
                g15 = g13 * 7 / 8;
                g16 = g11 + g14;
                g17 = g11 + g15;
                var d = "M" + 0 + "," + h +
                    " L" + w + "," + h +
                    " L" + w + "," + 0 +
                    " L" + 0 + "," + 0 +
                    " z" +
                    " M" + g17 + "," + g9 +
                    " L" + g12 + "," + g9 +
                    " L" + g12 + "," + g10 +
                    " L" + g17 + "," + g10 +
                    " z" +
                    " M" + g16 + "," + vc +
                    " L" + g11 + "," + g9 +
                    " L" + g11 + "," + g10 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonForwardNext": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;

                var d = "M" + 0 + "," + h +
                    " L" + w + "," + h +
                    " L" + w + "," + 0 +
                    " L" + 0 + "," + 0 +
                    " z" +
                    " M" + g12 + "," + vc +
                    " L" + g11 + "," + g9 +
                    " L" + g11 + "," + g10 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonHelp": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g11, g13, g14, g15, g16, g19, g20, g21, g23, g24, g27, g29, g30, g31, g33, g36, g37, g41, g42;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g11 = hc - dx2;
                g13 = ss * 3 / 4;
                g14 = g13 / 7;
                g15 = g13 * 3 / 14;
                g16 = g13 * 2 / 7;
                g19 = g13 * 3 / 7;
                g20 = g13 * 4 / 7;
                g21 = g13 * 17 / 28;
                g23 = g13 * 21 / 28;
                g24 = g13 * 11 / 14;
                g27 = g9 + g16;
                g29 = g9 + g21;
                g30 = g9 + g23;
                g31 = g9 + g24;
                g33 = g11 + g15;
                g36 = g11 + g19;
                g37 = g11 + g20;
                g41 = g13 / 14;
                g42 = g13 * 3 / 28;
                var cX1 = g33 + g16;
                var cX2 = g36 + g14;
                var cY3 = g31 + g42;
                var cX4 = (g37 + g36 + g16) / 2;

                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    "M" + g33 + "," + g27 +
                    PPTXShapeUtils.shapeArcAlt(cX1, g27, g16, g16, 180, 360, false).replace("M", "L") +
                    PPTXShapeUtils.shapeArcAlt(cX4, g27, g14, g15, 0, 90, false).replace("M", "L") +
                    PPTXShapeUtils.shapeArcAlt(cX4, g29, g41, g42, 270, 180, false).replace("M", "L") +
                    " L" + g37 + "," + g30 +
                    " L" + g36 + "," + g30 +
                    " L" + g36 + "," + g29 +
                    PPTXShapeUtils.shapeArcAlt(cX2, g29, g14, g15, 180, 270, false).replace("M", "L") +
                    PPTXShapeUtils.shapeArcAlt(g37, g27, g41, g42, 90, 0, false).replace("M", "L") +
                    PPTXShapeUtils.shapeArcAlt(cX1, g27, g14, g14, 0, -180, false).replace("M", "L") +
                    " z" +
                    "M" + hc + "," + g31 +
                    PPTXShapeUtils.shapeArcAlt(hc, cY3, g42, g42, 270, 630, false).replace("M", "L") +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonHome": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, g27, g28, g29, g30, g31, g32, g33;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 / 16;
                g15 = g13 / 8;
                g16 = g13 * 3 / 16;
                g17 = g13 * 5 / 16;
                g18 = g13 * 7 / 16;
                g19 = g13 * 9 / 16;
                g20 = g13 * 11 / 16;
                g21 = g13 * 3 / 4;
                g22 = g13 * 13 / 16;
                g23 = g13 * 7 / 8;
                g24 = g9 + g14;
                g25 = g9 + g16;
                g26 = g9 + g17;
                g27 = g9 + g21;
                g28 = g11 + g15;
                g29 = g11 + g18;
                g30 = g11 + g19;
                g31 = g11 + g20;
                g32 = g11 + g22;
                g33 = g11 + g23;

                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    " M" + hc + "," + g9 +
                    " L" + g11 + "," + vc +
                    " L" + g28 + "," + vc +
                    " L" + g28 + "," + g10 +
                    " L" + g33 + "," + g10 +
                    " L" + g33 + "," + vc +
                    " L" + g12 + "," + vc +
                    " L" + g32 + "," + g26 +
                    " L" + g32 + "," + g24 +
                    " L" + g31 + "," + g24 +
                    " L" + g31 + "," + g25 +
                    " z" +
                    " M" + g29 + "," + g27 +
                    " L" + g30 + "," + g27 +
                    " L" + g30 + "," + g10 +
                    " L" + g29 + "," + g10 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonInformation": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g11, g13, g14, g17, g18, g19, g20, g22, g23, g24, g25, g28, g29, g30, g31, g32, g34, g35, g37, g38;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g11 = hc - dx2;
                g13 = ss * 3 / 4;
                g14 = g13 / 32;
                g17 = g13 * 5 / 16;
                g18 = g13 * 3 / 8;
                g19 = g13 * 13 / 32;
                g20 = g13 * 19 / 32;
                g22 = g13 * 11 / 16;
                g23 = g13 * 13 / 16;
                g24 = g13 * 7 / 8;
                g25 = g9 + g14;
                g28 = g9 + g17;
                g29 = g9 + g18;
                g30 = g9 + g23;
                g31 = g9 + g24;
                g32 = g11 + g17;
                g34 = g11 + g19;
                g35 = g11 + g20;
                g37 = g11 + g22;
                g38 = g13 * 3 / 32;
                var cY1 = g9 + dx2;
                var cY2 = g25 + g38;

                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    "M" + hc + "," + g9 +
                    PPTXShapeUtils.shapeArcAlt(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L") +
                    " z" +
                    "M" + hc + "," + g25 +
                    PPTXShapeUtils.shapeArcAlt(hc, cY2, g38, g38, 270, 630, false).replace("M", "L") +
                    "M" + g32 + "," + g28 +
                    " L" + g35 + "," + g28 +
                    " L" + g35 + "," + g30 +
                    " L" + g37 + "," + g30 +
                    " L" + g37 + "," + g31 +
                    " L" + g32 + "," + g31 +
                    " L" + g32 + "," + g30 +
                    " L" + g34 + "," + g30 +
                    " L" + g34 + "," + g29 +
                    " L" + g32 + "," + g29 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonMovie": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, g27,
                    g28, g29, g30, g31, g32, g33, g34, g35, g36, g37, g38, g39, g40, g41, g42, g43, g44, g45, g46, g47, g48;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 * 1455 / 21600;
                g15 = g13 * 1905 / 21600;
                g16 = g13 * 2325 / 21600;
                g17 = g13 * 16155 / 21600;
                g18 = g13 * 17010 / 21600;
                g19 = g13 * 19335 / 21600;
                g20 = g13 * 19725 / 21600;
                g21 = g13 * 20595 / 21600;
                g22 = g13 * 5280 / 21600;
                g23 = g13 * 5730 / 21600;
                g24 = g13 * 6630 / 21600;
                g25 = g13 * 7492 / 21600;
                g26 = g13 * 9067 / 21600;
                g27 = g13 * 9555 / 21600;
                g28 = g13 * 13342 / 21600;
                g29 = g13 * 14580 / 21600;
                g30 = g13 * 15592 / 21600;
                g31 = g11 + g14;
                g32 = g11 + g15;
                g33 = g11 + g16;
                g34 = g11 + g17;
                g35 = g11 + g18;
                g36 = g11 + g19;
                g37 = g11 + g20;
                g38 = g11 + g21;
                g39 = g9 + g22;
                g40 = g9 + g23;
                g41 = g9 + g24;
                g42 = g9 + g25;
                g43 = g9 + g26;
                g44 = g9 + g27;
                g45 = g9 + g28;
                g46 = g9 + g29;
                g47 = g9 + g30;
                g48 = g9 + g31;

                var d = "M" + 0 + "," + h +
                    " L" + w + "," + h +
                    " L" + w + "," + 0 +
                    " L" + 0 + "," + 0 +
                    " z" +
                    "M" + g11 + "," + g39 +
                    " L" + g11 + "," + g44 +
                    " L" + g31 + "," + g44 +
                    " L" + g32 + "," + g43 +
                    " L" + g33 + "," + g43 +
                    " L" + g33 + "," + g47 +
                    " L" + g35 + "," + g47 +
                    " L" + g35 + "," + g45 +
                    " L" + g36 + "," + g45 +
                    " L" + g38 + "," + g46 +
                    " L" + g12 + "," + g46 +
                    " L" + g12 + "," + g41 +
                    " L" + g38 + "," + g41 +
                    " L" + g37 + "," + g42 +
                    " L" + g35 + "," + g42 +
                    " L" + g35 + "," + g41 +
                    " L" + g34 + "," + g40 +
                    " L" + g32 + "," + g40 +
                    " L" + g31 + "," + g39 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonReturn": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, g27;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 * 7 / 8;
                g15 = g13 * 3 / 4;
                g16 = g13 * 5 / 8;
                g17 = g13 * 3 / 8;
                g18 = g13 / 4;
                g19 = g9 + g15;
                g20 = g9 + g16;
                g21 = g9 + g18;
                g22 = g11 + g14;
                g23 = g11 + g15;
                g24 = g11 + g16;
                g25 = g11 + g17;
                g26 = g11 + g18;
                g27 = g13 / 8;
                var cX1 = g24 - g27;
                var cY2 = g19 - g27;
                var cX3 = g11 + g17;
                var cY4 = g10 - g17;

                var d = "M" + 0 + "," + h +
                    " L" + w + "," + h +
                    " L" + w + "," + 0 +
                    " L" + 0 + "," + 0 +
                    " z" +
                    " M" + g12 + "," + g21 +
                    " L" + g23 + "," + g9 +
                    " L" + hc + "," + g21 +
                    " L" + g24 + "," + g21 +
                    " L" + g24 + "," + g20 +
                    PPTXShapeUtils.shapeArcAlt(cX1, g20, g27, g27, 0, 90, false).replace("M", "L") +
                    " L" + g25 + "," + g19 +
                    PPTXShapeUtils.shapeArcAlt(g25, cY2, g27, g27, 90, 180, false).replace("M", "L") +
                    " L" + g26 + "," + g21 +
                    " L" + g11 + "," + g21 +
                    " L" + g11 + "," + g20 +
                    PPTXShapeUtils.shapeArcAlt(cX3, g20, g17, g17, 180, 90, false).replace("M", "L") +
                    " L" + hc + "," + g10 +
                    PPTXShapeUtils.shapeArcAlt(hc, cY4, g17, g17, 90, 0, false).replace("M", "L") +
                    " L" + g22 + "," + g21 +
                    " z";

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
            case "actionButtonSound": {
                var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
                var dx2, g9, g10, g11, g12, g13, g14, g15, g16, g17, g18, g19, g20, g21, g22, g23, g24, g25, g26;

                dx2 = ss * 3 / 8;
                g9 = vc - dx2;
                g10 = vc + dx2;
                g11 = hc - dx2;
                g12 = hc + dx2;
                g13 = ss * 3 / 4;
                g14 = g13 / 8;
                g15 = g13 * 5 / 16;
                g16 = g13 * 5 / 8;
                g17 = g13 * 11 / 16;
                g18 = g13 * 3 / 4;
                g19 = g13 * 7 / 8;
                g20 = g9 + g14;
                g21 = g9 + g15;
                g22 = g9 + g17;
                g23 = g9 + g19;
                g24 = g11 + g15;
                g25 = g11 + g16;
                g26 = g11 + g18;

                var d = "M" + 0 + "," + 0 +
                    " L" + w + "," + 0 +
                    " L" + w + "," + h +
                    " L" + 0 + "," + h +
                    " z" +
                    " M" + g11 + "," + g21 +
                    " L" + g24 + "," + g21 +
                    " L" + g25 + "," + g9 +
                    " L" + g25 + "," + g10 +
                    " L" + g24 + "," + g22 +
                    " L" + g11 + "," + g22 +
                    " z" +
                    " M" + g26 + "," + g21 +
                    " L" + g12 + "," + g20 +
                    " M" + g26 + "," + vc +
                    " L" + g12 + "," + vc +
                    " M" + g26 + "," + g22 +
                    " L" + g12 + "," + g23;

                result += "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                    "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                break;
            }
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
                result += PPTXEllipseShapes.generateEllipse(w, h, shapType, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartTerminator":
                result += PPTXEllipseShapes.generateFlowChartTerminator(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartPunchedTape":
                result += PPTXEllipseShapes.generateFlowChartPunchedTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartOnlineStorage":
                result += PPTXEllipseShapes.generateFlowChartOnlineStorage(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartDisplay":
                result += PPTXEllipseShapes.generateFlowChartDisplay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartDelay":
                result += PPTXEllipseShapes.generateFlowChartDelay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "flowChartMagneticTape":
                result += PPTXEllipseShapes.generateFlowChartMagneticTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "rtTriangle":
                result += PPTXPolygonShapes.generateRtTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "triangle":
            case "flowChartExtract":
            case "flowChartMerge":
                var triangleAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXPolygonShapes.generateTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, triangleAdjst);
                break;
            case "diamond":
            case "flowChartDecision":
            case "flowChartSort":
                result += PPTXPolygonShapes.generateDiamond(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType);
                break;
            case "trapezoid":
            case "flowChartManualOperation":
            case "flowChartManualInput":
                var trapezoidAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXPolygonShapes.generateTrapezoid(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, trapezoidAdjst);
                break;
            case "parallelogram":
            case "flowChartInputOutput":
                var parallelogramAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXPolygonShapes.generateParallelogram(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, parallelogramAdjst);
                break;
            case "pentagon":
                result += PPTXPolygonShapes.generatePentagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                break;
            case "hexagon":
            case "flowChartPreparation":
                var hexagonAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXPolygonShapes.generateHexagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, hexagonAdjst);
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
            case "pie":
            case "pieWedge":
            case "arc": {
                var pieAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generatePie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, pieAdjst);
                break;
            }
            case "chord": {
                var chordAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateChord(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, chordAdjst);
                break;
            }
            case "frame": {
                var frameAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateFrame(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, frameAdjst);
                break;
            }
            case "donut": {
                var donutAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateDonut(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, donutAdjst);
                break;
            }
            case "noSmoking": {
                var noSmokingAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateNoSmoking(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, noSmokingAdjst);
                break;
            }
            case "halfFrame": {
                var halfFrameAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateHalfFrame(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, halfFrameAdjst);
                break;
            }
            case "blockArc": {
                var blockArcAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateBlockArc(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, blockArcAdjst);
                break;
            }
            case "bracePair": {
                var bracePairAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateBracePair(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, bracePairAdjst);
                break;
            }
            case "leftBrace": {
                var leftBraceAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateLeftBrace(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, leftBraceAdjst);
                break;
            }
            case "rightBrace": {
                var rightBraceAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateRightBrace(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, rightBraceAdjst);
                break;
            }
            case "bracketPair": {
                var bracketPairAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateBracketPair(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, bracketPairAdjst);
                break;
            }
            case "leftBracket": {
                var leftBracketAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateLeftBracket(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, leftBracketAdjst);
                break;
            }
            case "rightBracket": {
                var rightBracketAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateRightBracket(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, rightBracketAdjst);
                break;
            }
            case "moon": {
                var moonAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateMoon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, moonAdjst);
                break;
            }
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
            case "irregularSeal1":
            case "irregularSeal2":
                result += PPTXSpecialShapes.generateIrregularSeal(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType);
                break;
            case "corner":
                var cornerAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateCorner(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, cornerAdjst_ary);
                break;
            case "diagStripe":
                var diagStripeAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateDiagStripe(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, diagStripeAdjst);
                break;
            case "gear6":
            case "gear9":
                result += PPTXSpecialShapes.generateGear(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType);
                break;
            case "can":
            case "flowChartMagneticDisk":
            case "flowChartMagneticDrum":
                var canAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                result += PPTXSpecialShapes.generateCan(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, canAdjst);
                break;
            case "swooshArrow":
                var swooshArrowAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateSwooshArrow(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, swooshArrowAdjst_ary);
                break;
            case "circularArrow":
                var circularArrowAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                result += PPTXSpecialShapes.generateCircularArrow(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, circularArrowAdjst_ary);
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