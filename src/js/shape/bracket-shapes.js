/**
 * 括号形状渲染模块
 * 提供各种括号形状（大括号、方括号等）的生成和渲染功能
 */

import { PPTXXmlUtils } from '../utils/xml.js';
import { shapeArc } from './path-generators.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * 检查形状是否为括号形状
 */
export function isBracket(shapType) {
    const bracketShapes = [
        'bracePair', 'bracketPair', 'leftBrace', 'leftBracket', 'rightBrace', 'rightBracket'
    ];
    return bracketShapes.includes(shapType);
}

/**
 * 渲染括号形状
 */
export function renderBracket(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";

    if (shapType === "bracePair") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR;
        var cnstVal1 = 25000 * SLIDE_FACTOR;
        var cnstVal2 = 50000 * SLIDE_FACTOR;
        var cnstVal3 = 100000 * SLIDE_FACTOR;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
        }
        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, x3, x4, y2, y3, y4;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        var minWH = Math.min(w, h);
        x1 = minWH * a / cnstVal3;
        x2 = minWH * a / cnstVal2;
        x3 = w - x2;
        x4 = w - x1;
        y2 = vc - x1;
        y3 = vc + x1;
        y4 = h - x1;
        dVal = "M" + x2 + "," + h +
            shapeArc(x2, y4, x1, x1, cd4, cd2, false).replace("M", "L") +
            " L" + x1 + "," + y3 +
            shapeArc(0, y3, x1, x1, 0, (-cd4), false).replace("M", "L") +
            shapeArc(0, y2, x1, x1, cd4, 0, false).replace("M", "L") +
            " L" + x1 + "," + x1 +
            shapeArc(x2, x1, x1, x1, cd2, c3d4, false).replace("M", "L") +
            " M" + x3 + "," + 0 +
            shapeArc(x3, x1, x1, x1, c3d4, cd, false).replace("M", "L") +
            " L" + x4 + "," + y2 +
            shapeArc(w, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
            shapeArc(w, y3, x1, x1, c3d4, cd2, false).replace("M", "L") +
            " L" + x4 + "," + y4 +
            shapeArc(x3, y4, x1, x1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "leftBrace") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 8333 * SLIDE_FACTOR;
        var sAdj2, adj2 = 50000 * SLIDE_FACTOR;
        var cnstVal2 = 100000 * SLIDE_FACTOR;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                }
            }
        }
        var vc = h / 2, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0
        else if (adj2 > cnstVal2) a2 = cnstVal2
        else a2 = adj2
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2) q2 = q1
        else q2 = a2
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0
        else if (adj1 > maxAdj1) a1 = maxAdj1
        else a1 = adj1
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = y3 + y1;
        dVal = "M" + w + "," + h +
            shapeArc(w, h - y1, w / 2, y1, cd4, cd2, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            shapeArc(0, y4, w / 2, y1, 0, (-cd4), false).replace("M", "L") +
            shapeArc(0, y2, w / 2, y1, cd4, 0, false).replace("M", "L") +
            " L" + w / 2 + "," + y1 +
            shapeArc(w, y1, w / 2, y1, cd2, c3d4, false).replace("M", "L");
    }
    else if (shapType === "rightBrace") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 8333 * SLIDE_FACTOR;
        var sAdj2, adj2 = 50000 * SLIDE_FACTOR;
        var cnstVal2 = 100000 * SLIDE_FACTOR;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                }
            }
        }
        var vc = h / 2, cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0
        else if (adj2 > cnstVal2) a2 = cnstVal2
        else a2 = adj2
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        if (q1 < a2) q2 = q1
        else q2 = a2
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0
        else if (adj1 > maxAdj1) a1 = maxAdj1
        else a1 = adj1
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = h - y1;
        dVal = "M" + 0 + "," + 0 +
            shapeArc(0, y1, w / 2, y1, c3d4, cd, false).replace("M", "L") +
            " L" + w / 2 + "," + y2 +
            shapeArc(w, y2, w / 2, y1, cd2, cd4, false).replace("M", "L") +
            shapeArc(w, y3 + y1, w / 2, y1, c3d4, cd2, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            shapeArc(0, y4, w / 2, y1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "bracketPair") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 16667 * SLIDE_FACTOR;
        var cnstVal1 = 50000 * SLIDE_FACTOR;
        var cnstVal2 = 100000 * SLIDE_FACTOR;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
        }
        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, x1, x2, y2;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        x1 = Math.min(w, h) * a / cnstVal2;
        x2 = r - x1;
        y2 = b - x1;
        dVal = shapeArc(x1, x1, x1, x1, c3d4, cd2, false) +
            shapeArc(x1, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
            shapeArc(x2, x1, x1, x1, c3d4, (c3d4 + cd4), false) +
            shapeArc(x2, y2, x1, x1, 0, cd4, false).replace("M", "L");
    }
    else if (shapType === "leftBracket") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR;
        var cnstVal1 = 50000 * SLIDE_FACTOR;
        var cnstVal2 = 100000 * SLIDE_FACTOR;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
        }
        var r = w, b = h, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2;
        if (adj < 0) a = 0
        else if (adj > maxAdj) a = maxAdj
        else a = adj
        y1 = Math.min(w, h) * a / cnstVal2;
        if (y1 > w) y1 = w;
        y2 = b - y1;
        dVal = "M" + r + "," + b +
            shapeArc(y1, y2, y1, y1, cd4, cd2, false).replace("M", "L") +
            " L" + 0 + "," + y1 +
            shapeArc(y1, y1, y1, y1, cd2, c3d4, false).replace("M", "L") +
            " L" + r + "," + 0;
    }
    else if (shapType === "rightBracket") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var adj = 8333 * SLIDE_FACTOR;
        var cnstVal1 = 50000 * SLIDE_FACTOR;
        var cnstVal2 = 100000 * SLIDE_FACTOR;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
        }
        var cd = 360, cd2 = 180, cd4 = 90, c3d4 = 270, a, y1, y2, y3;
        if (adj < 0) a = 0
        else if (adj > maxAdj) a = maxAdj
        else a = adj
        y1 = Math.min(w, h) * a / cnstVal2;
        y2 = h - y1;
        y3 = w - y1;
        dVal = "M" + 0 + "," + h +
            shapeArc(y3, y2, y1, y1, cd4, 0, false).replace("M", "L") +
            " L" + w + "," + h / 2 +
            shapeArc(y3, y1, y1, y1, cd, c3d4, false).replace("M", "L") +
            " L" + 0 + "," + 0;
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}
