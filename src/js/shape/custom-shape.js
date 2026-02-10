/**
 * 自定义形状 (custGeom) 渲染模块
 * 处理 PowerPoint 中的自定义几何形状
 * 参考: http://officeopenxml.com/drwSp-custGeom.php
 */

import { PPTXXmlUtils } from '../utils/xml.js';

/**
 * 渲染自定义形状
 * @param {Object} custShapType - 自定义形状数据
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否图片填充
 * @param {boolean} grndFillFlg - 是否渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框样式
 * @param {string} shpId - 形状ID
 * @param {Function} shapeArcFn - 圆弧路径生成函数
 * @returns {string} SVG路径元素
 */
export function renderCustomShape(custShapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcFn) {
    var pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
    var pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);

    // 验证 maxX 和 maxY 防止 NaN
    var maxX = 0;
    var maxY = 0;
    if (pathNodes && pathNodes["attrs"]) {
        maxX = parseInt(pathNodes["attrs"]["w"]) || 0;
        maxY = parseInt(pathNodes["attrs"]["h"]) || 0;
    }
    // 确保 maxX 和 maxY 为正数以避免除零
    if (maxX <= 0) maxX = 1;
    if (maxY <= 0) maxY = 1;
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
        // a:moveTo
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

        // a:lnTo
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

        // a:cubicBezTo
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
                    };
                    pts_ary.push(pt_obj);
                });
                nodeObj.cubBzPt = pts_ary;
                multiSapeAry.push(nodeObj);
            });
        }

        // a:arcTo
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

        // a:close
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

        // 按 order 排序
        multiSapeAry.sort(function (a, b) {
            return a.order - b.order;
        });

        // 生成路径字符串
        var k = 0;
        if (isNaN(cX)) cX = 0;
        if (isNaN(cY)) cY = 0;
        var d = "";
        while (k < multiSapeAry.length) {
            if (multiSapeAry[k].type == "movto") {
                var xVal = parseInt(multiSapeAry[k].x) || 0;
                var yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var spX = xVal * cX;
                var spY = yVal * cY;
                d += " M" + spX + "," + spY;
            } else if (multiSapeAry[k].type == "lnto") {
                var xVal = parseInt(multiSapeAry[k].x) || 0;
                var yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var Lx = xVal * cX;
                var Ly = yVal * cY;
                d += " L" + Lx + "," + Ly;
            } else if (multiSapeAry[k].type == "cubicBezTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var Cx1 = (parseInt(multiSapeAry[k].cubBzPt[0].x) || 0) * cX;
                var Cy1 = (parseInt(multiSapeAry[k].cubBzPt[0].y) || 0) * cY;
                var Cx2 = (parseInt(multiSapeAry[k].cubBzPt[1].x) || 0) * cX;
                var Cy2 = (parseInt(multiSapeAry[k].cubBzPt[1].y) || 0) * cY;
                var Cx3 = (parseInt(multiSapeAry[k].cubBzPt[2].x) || 0) * cX;
                var Cy3 = (parseInt(multiSapeAry[k].cubBzPt[2].y) || 0) * cY;
                d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
            } else if (multiSapeAry[k].type == "arcTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var hR = (parseInt(multiSapeAry[k].hR) || 0) * cX;
                var wR = (parseInt(multiSapeAry[k].wR) || 0) * cY;
                var stAng = (parseInt(multiSapeAry[k].stAng) || 0) / 60000;
                var swAng = (parseInt(multiSapeAry[k].swAng) || 0) / 60000;
                if (isNaN(stAng)) stAng = 0;
                if (isNaN(swAng)) swAng = 0;
                var endAng = stAng + swAng;

                if (isNaN(hR) || isNaN(wR) || isNaN(stAng) || isNaN(swAng)) {
                    console.warn("Invalid arc parameters detected");
                } else {
                    d += shapeArcFn(wR, hR, wR, hR, stAng, endAng, false);
                }
            } else if (multiSapeAry[k].type == "quadBezTo") {
                console.log("custShapType: quadBezTo - TODO");
            } else if (multiSapeAry[k].type == "close") {
                d += "z";
            }
            k++;
        }

        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' />";
    }

    return "";
}
