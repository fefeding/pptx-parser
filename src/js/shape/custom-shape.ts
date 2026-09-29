/**
 * 自定义形状 (custGeom) 渲染模块
 * 处理 PowerPoint 中的自定义几何形状
 * 参考: http://officeopenxml.com/drwSp-custGeom.php
 */

import { PPTXXmlUtils } from '../utils/xml';

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
export function renderCustomShape(custShapType: any, w: any, h: any, imgFillFlg: any, grndFillFlg: any, fillColor: any, border: any, shpId: any, shapeArcFn: any) {
    const pathLstNode = PPTXXmlUtils.getTextByPathList(custShapType, ["a:pathLst"]);
    const pathNodes = PPTXXmlUtils.getTextByPathList(pathLstNode, ["a:path"]);

    // 验证 maxX 和 maxY 防止 NaN
    let maxX = 0;
    let maxY = 0;
    if (pathNodes && pathNodes["attrs"]) {
        maxX = parseInt(pathNodes["attrs"]["w"]) || 0;
        maxY = parseInt(pathNodes["attrs"]["h"]) || 0;
    }
    // 确保 maxX 和 maxY 为正数以避免除零
    if (maxX <= 0) maxX = 1;
    if (maxY <= 0) maxY = 1;
    let cX = (1 / maxX) * w;
    let cY = (1 / maxY) * h;

    let moveToNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
    const total_shapes = moveToNode.length;

    const lnToNodes = pathNodes["a:lnTo"];
    let cubicBezToNodes = pathNodes["a:cubicBezTo"];
    const arcToNodes = pathNodes["a:arcTo"];
    let closeNode = PPTXXmlUtils.getTextByPathList(pathNodes, ["a:close"]);

    if (!Array.isArray(moveToNode)) {
        moveToNode = [moveToNode];
    }

    const multiSapeAry = [];
    if (moveToNode.length > 0) {
        // a:moveTo
        Object.keys(moveToNode).forEach((key) => {
    var moveToPtNode = moveToNode[key]["a:pt"];
    if (moveToPtNode !== undefined) {
        Object.keys(moveToPtNode).forEach((key2) => {
    var ptObj: any = {};
    var moveToNoPt = moveToPtNode[key2];
    
    // moveToNoPt 已是 a:pt 的 attrs 对象（由 Object.keys 迭代得到），直接取 x
    var spX = moveToNoPt["x"];
    
    var spY = moveToNoPt["y"];
    
    var ptOrdr = moveToNoPt["order"];
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
            Object.keys(lnToNodes).forEach((key) => {
    var lnToPtNode = lnToNodes[key]["a:pt"];
    if (lnToPtNode !== undefined) {
        Object.keys(lnToPtNode).forEach((key2) => {
    var ptObj: any = {};
    var lnToNoPt = lnToPtNode[key2];
    
    // lnToNoPt 已是 a:pt 的 attrs 对象，直接取 x
    var ptX = lnToNoPt["x"];
    
    var ptY = lnToNoPt["y"];
    
    var ptOrdr = lnToNoPt["order"];
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
            const cubicBezToPtNodesAry: any = [];
            if (!Array.isArray(cubicBezToNodes)) {
                cubicBezToNodes = [cubicBezToNodes];
            }
            Object.keys(cubicBezToNodes).forEach((key) => {
    cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
});

            cubicBezToPtNodesAry.forEach((key2: any) => {
    var nodeObj: any = {};
    nodeObj.type = "cubicBezTo";
    nodeObj.order = key2[0]["attrs"]["order"];
    var pts_ary: any = [];
    key2.forEach((pt: any) => {
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

        // a:quadBezTo
        let quadBezToNodes = pathNodes["a:quadBezTo"];
        if (quadBezToNodes !== undefined) {
            const quadBezToPtNodesAry: any = [];
            if (!Array.isArray(quadBezToNodes)) {
                quadBezToNodes = [quadBezToNodes];
            }
            Object.keys(quadBezToNodes).forEach((key) => {
    quadBezToPtNodesAry.push(quadBezToNodes[key]["a:pt"]);
});

            quadBezToPtNodesAry.forEach((key2: any) => {
    var nodeObj: any = {};
    nodeObj.type = "quadBezTo";
    nodeObj.order = key2[0]["attrs"]["order"];
    var pts_ary: any = [];
    key2.forEach((pt: any) => {
    var pt_obj = {
        x: pt["attrs"]["x"],
        y: pt["attrs"]["y"]
    };
    pts_ary.push(pt_obj);
});
    nodeObj.quadBzPt = pts_ary;
    multiSapeAry.push(nodeObj);
});
        }

        // a:arcTo
        if (arcToNodes !== undefined) {
            const arcToNodesAttrs = arcToNodes["attrs"];
            const arcOrder = arcToNodesAttrs["order"];
            const hR = arcToNodesAttrs["hR"];
            const wR = arcToNodesAttrs["wR"];
            let stAng = arcToNodesAttrs["stAng"];
            let swAng = arcToNodesAttrs["swAng"];
            let shftX = 0;
            let shftY = 0;
            const arcToPtNode = PPTXXmlUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
            if (arcToPtNode !== undefined) {
                shftX = arcToPtNode["x"];
                shftY = arcToPtNode["y"];
            }
            var ptObj: any = {};
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
            Object.keys(closeNode).forEach((key) => {
    var clsAttrs = closeNode[key]["attrs"];
    var clsOrder = clsAttrs["order"];
    var ptObj: any = {};
    ptObj.type = "close";
    ptObj.order = clsOrder;
    multiSapeAry.push(ptObj);
});
        }

        // 按 order 排序
        multiSapeAry.sort((a, b) => {
    return a.order - b.order;
});

        // 生成路径字符串
        let k = 0;
        if (isNaN(cX)) cX = 0;
        if (isNaN(cY)) cY = 0;
        let d = "";
        while (k < multiSapeAry.length) {
            if (multiSapeAry[k].type == "movto") {
                const xVal = parseInt(multiSapeAry[k].x) || 0;
                const yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                var spX = xVal * cX;
                var spY = yVal * cY;
                d += ` M${spX},${spY}`;
            } else if (multiSapeAry[k].type == "lnto") {
                const xVal = parseInt(multiSapeAry[k].x) || 0;
                const yVal = parseInt(multiSapeAry[k].y) || 0;
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                const Lx = xVal * cX;
                const Ly = yVal * cY;
                d += ` L${Lx},${Ly}`;
            } else if (multiSapeAry[k].type == "cubicBezTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                const Cx1 = (parseInt(multiSapeAry[k].cubBzPt[0].x) || 0) * cX;
                const Cy1 = (parseInt(multiSapeAry[k].cubBzPt[0].y) || 0) * cY;
                const Cx2 = (parseInt(multiSapeAry[k].cubBzPt[1].x) || 0) * cX;
                const Cy2 = (parseInt(multiSapeAry[k].cubBzPt[1].y) || 0) * cY;
                const Cx3 = (parseInt(multiSapeAry[k].cubBzPt[2].x) || 0) * cX;
                const Cy3 = (parseInt(multiSapeAry[k].cubBzPt[2].y) || 0) * cY;
                d += ` C${Cx1},${Cy1} ${Cx2},${Cy2} ${Cx3},${Cy3}`;
            } else if (multiSapeAry[k].type == "arcTo") {
                if (isNaN(cX)) cX = 0;
                if (isNaN(cY)) cY = 0;
                const hR = (parseInt(multiSapeAry[k].hR) || 0) * cX;
                const wR = (parseInt(multiSapeAry[k].wR) || 0) * cY;
                let stAng = (parseInt(multiSapeAry[k].stAng) || 0) / 60000;
                let swAng = (parseInt(multiSapeAry[k].swAng) || 0) / 60000;
                if (isNaN(stAng)) stAng = 0;
                if (isNaN(swAng)) swAng = 0;
                const endAng = stAng + swAng;

                if (!isNaN(hR) && !isNaN(wR) && !isNaN(stAng) && !isNaN(swAng)) {
                    d += shapeArcFn(wR, hR, wR, hR, stAng, endAng, false);
                }
            } else if (multiSapeAry[k].type == "quadBezTo") {
                // Quadratic Bezier curve: Q controlPoint endPoint
                // PPTX quadBezTo has 2 points: control point and end point
                // The start point is the previous point in the path
                var quadBzPt = multiSapeAry[k].quadBzPt;
                if (quadBzPt && quadBzPt.length >= 2) {
                    const ctrlX = quadBzPt[0].x * cX;
                    const ctrlY = quadBzPt[0].y * cY;
                    const endX = quadBzPt[1].x * cX;
                    const endY = quadBzPt[1].y * cY;
                    d += `Q${ctrlX},${ctrlY} ${endX},${endY}`;
                }
            } else if (multiSapeAry[k].type == "close") {
                d += "z";
            }
            k++;
        }

        return `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${((border === undefined) ? "" : border.color)}' stroke-width='${((border === undefined) ? "" : border.width)}' stroke-dasharray='${((border === undefined) ? "" : border.strokeDasharray)}' />`;
    }

    return "";
}
