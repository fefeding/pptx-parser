/**
 * 自定义形状处理模块
 * 处理PPTX中的自定义几何形状
 */

var PPTXCustomShapes = (function() {
    /**
     * 处理自定义形状
     * @param {Object} custShapType - 自定义形状类型节点
     * @param {number} w - 宽度
     * @param {number} h - 高度
     * @param {boolean} imgFillFlg - 图片填充标志
     * @param {boolean} grndFillFlg - 渐变填充标志
     * @param {string} shpId - 形状ID
     * @param {string} fillColor - 填充颜色
     * @param {Object} border - 边框对象
     * @returns {string} SVG路径字符串
     */
    function processCustomShape(custShapType, w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
        var pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);
        
        if (pathNodes === undefined || pathNodes["attrs"] === undefined) {
            return "";
        }
        
        var maxX = parseInt(pathNodes["attrs"]["w"]) || 1;
        var maxY = parseInt(pathNodes["attrs"]["h"]) || 1;
        var cX = (1 / maxX) * w;
        var cY = (1 / maxY) * h;

        var moveToNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
        var lnToNodes = pathNodes["a:lnTo"];
        var cubicBezToNodes = pathNodes["a:cubicBezTo"];
        var arcToNodes = pathNodes["a:arcTo"];
        var closeNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:close"]);

        if (moveToNode === undefined) {
            moveToNode = [];
        }
        if (!Array.isArray(moveToNode)) {
            moveToNode = [moveToNode];
        }

        var multiSapeAry = parsePathNodes(moveToNode, lnToNodes, cubicBezToNodes, arcToNodes, closeNode);

        multiSapeAry.sort(function(a, b) {
            return a.order - b.order;
        });

        var d = generatePathFromNodes(multiSapeAry, cX, cY);

        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' />";
    }

    /**
     * 解析路径节点
     */
    function parsePathNodes(moveToNode, lnToNodes, cubicBezToNodes, arcToNodes, closeNode) {
        var multiSapeAry = [];

        if (moveToNode.length > 0) {
            parseMoveToNodes(moveToNode, multiSapeAry);
        }

        if (lnToNodes !== undefined) {
            parseLnToNodes(lnToNodes, multiSapeAry);
        }

        if (cubicBezToNodes !== undefined) {
            parseCubicBezToNodes(cubicBezToNodes, multiSapeAry);
        }

        if (arcToNodes !== undefined) {
            parseArcToNodes(arcToNodes, multiSapeAry);
        }

        if (closeNode !== undefined) {
            parseCloseNodes(closeNode, multiSapeAry);
        }

        return multiSapeAry;
    }

    /**
     * 解析moveTo节点
     */
    function parseMoveToNodes(moveToNode, multiSapeAry) {
        Object.keys(moveToNode).forEach(function(key) {
            var moveToPtNode = moveToNode[key]["a:pt"];
            if (moveToPtNode !== undefined) {
                Object.keys(moveToPtNode).forEach(function(key2) {
                    var ptObj = {};
                    var moveToNoPt = moveToPtNode[key2];
                    var attrs = moveToNoPt["attrs"];
                    if (attrs !== undefined) {
                        var spX = attrs["x"];
                        var spY = attrs["y"];
                        var ptOrdr = attrs["order"];
                        ptObj.type = "movto";
                        ptObj.order = ptOrdr;
                        ptObj.x = spX;
                        ptObj.y = spY;
                        multiSapeAry.push(ptObj);
                    }
                });
            }
        });
    }

    /**
     * 解析lnTo节点
     */
    function parseLnToNodes(lnToNodes, multiSapeAry) {
        Object.keys(lnToNodes).forEach(function(key) {
            var lnToPtNode = lnToNodes[key]["a:pt"];
            if (lnToPtNode !== undefined) {
                Object.keys(lnToPtNode).forEach(function(key2) {
                    var ptObj = {};
                    var lnToNoPt = lnToPtNode[key2];
                    var attrs = lnToNoPt["attrs"];
                    if (attrs !== undefined) {
                        var ptX = attrs["x"];
                        var ptY = attrs["y"];
                        var ptOrdr = attrs["order"];
                        ptObj.type = "lnto";
                        ptObj.order = ptOrdr;
                        ptObj.x = ptX;
                        ptObj.y = ptY;
                        multiSapeAry.push(ptObj);
                    }
                });
            }
        });
    }

    /**
     * 解析cubicBezTo节点
     */
    function parseCubicBezToNodes(cubicBezToNodes, multiSapeAry) {
        var cubicBezToPtNodesAry = [];
        if (!Array.isArray(cubicBezToNodes)) {
            cubicBezToNodes = [cubicBezToNodes];
        }
        Object.keys(cubicBezToNodes).forEach(function(key) {
            var ptNodes = cubicBezToNodes[key]["a:pt"];
            if (ptNodes !== undefined) {
                cubicBezToPtNodesAry.push(ptNodes);
            }
        });

        cubicBezToPtNodesAry.forEach(function(key2) {
            var nodeObj = {};
            nodeObj.type = "cubicBezTo";
            if (key2.length > 0 && key2[0] !== undefined && key2[0]["attrs"] !== undefined) {
                nodeObj.order = key2[0]["attrs"]["order"];
            }
            var pts_ary = [];
            key2.forEach(function(pt) {
                if (pt !== undefined && pt["attrs"] !== undefined) {
                    var pt_obj = {
                        x: pt["attrs"]["x"],
                        y: pt["attrs"]["y"]
                    }
                    pts_ary.push(pt_obj)
                }
            })
            nodeObj.cubBzPt = pts_ary;
            multiSapeAry.push(nodeObj);
        });
    }

    /**
     * 解析arcTo节点
     */
    function parseArcToNodes(arcToNodes, multiSapeAry) {
        var arcToNodesAttrs = arcToNodes["attrs"];
        if (arcToNodesAttrs === undefined) {
            return;
        }
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

    /**
     * 解析close节点
     */
    function parseCloseNodes(closeNode, multiSapeAry) {
        if (closeNode === undefined) {
            return;
        }
        if (!Array.isArray(closeNode)) {
            closeNode = [closeNode];
        }
        Object.keys(closeNode).forEach(function(key) {
            var clsAttrs = closeNode[key]["attrs"];
            if (clsAttrs !== undefined) {
                var clsOrder = clsAttrs["order"];
                var ptObj = {};
                ptObj.type = "close";
                ptObj.order = clsOrder;
                multiSapeAry.push(ptObj);
            }
        });
    }

    /**
     * 从节点生成路径
     */
    function generatePathFromNodes(multiSapeAry, cX, cY) {
        var k = 0;
        var d = "";
        while (k < multiSapeAry.length) {
            if (multiSapeAry[k].type == "movto") {
                var spX = (multiSapeAry[k].x !== undefined) ? parseInt(multiSapeAry[k].x) * cX : 0;
                var spY = (multiSapeAry[k].y !== undefined) ? parseInt(multiSapeAry[k].y) * cY : 0;
                d += " M" + spX + "," + spY;
            } else if (multiSapeAry[k].type == "lnto") {
                var Lx = (multiSapeAry[k].x !== undefined) ? parseInt(multiSapeAry[k].x) * cX : 0;
                var Ly = (multiSapeAry[k].y !== undefined) ? parseInt(multiSapeAry[k].y) * cY : 0;
                d += " L" + Lx + "," + Ly;
            } else if (multiSapeAry[k].type == "cubicBezTo") {
                var Cx1 = (multiSapeAry[k].cubBzPt[0] !== undefined && multiSapeAry[k].cubBzPt[0].x !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[0].x) * cX : 0;
                var Cy1 = (multiSapeAry[k].cubBzPt[0] !== undefined && multiSapeAry[k].cubBzPt[0].y !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[0].y) * cY : 0;
                var Cx2 = (multiSapeAry[k].cubBzPt[1] !== undefined && multiSapeAry[k].cubBzPt[1].x !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[1].x) * cX : 0;
                var Cy2 = (multiSapeAry[k].cubBzPt[1] !== undefined && multiSapeAry[k].cubBzPt[1].y !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[1].y) * cY : 0;
                var Cx3 = (multiSapeAry[k].cubBzPt[2] !== undefined && multiSapeAry[k].cubBzPt[2].x !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[2].x) * cX : 0;
                var Cy3 = (multiSapeAry[k].cubBzPt[2] !== undefined && multiSapeAry[k].cubBzPt[2].y !== undefined) ? parseInt(multiSapeAry[k].cubBzPt[2].y) * cY : 0;
                d += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
            } else if (multiSapeAry[k].type == "arcTo") {
                var hR = (multiSapeAry[k].hR !== undefined) ? parseInt(multiSapeAry[k].hR) * cX : 0;
                var wR = (multiSapeAry[k].wR !== undefined) ? parseInt(multiSapeAry[k].wR) * cY : 0;
                var stAng = (multiSapeAry[k].stAng !== undefined) ? parseInt(multiSapeAry[k].stAng) / 60000 : 0;
                var swAng = (multiSapeAry[k].swAng !== undefined) ? parseInt(multiSapeAry[k].swAng) / 60000 : 0;
                var endAng = stAng + swAng;
                d += PPTXShapeUtils.shapeArc(wR, hR, wR, hR, stAng, endAng, false);
            } else if (multiSapeAry[k].type == "quadBezTo") {
                console.log("custShapType: quadBezTo - TODO");
            } else if (multiSapeAry[k].type == "close") {
                d += "z";
            }
            k++;
        }
        return d;
    }

    return {
        processCustomShape: processCustomShape
    };
})();

window.PPTXCustomShapes = PPTXCustomShapes;