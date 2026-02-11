/**
 * 杂项形状渲染模块
 * 包含 smileyFace、scroll 等独立形状
 */

import { PPTXXmlUtils } from '../utils/xml.js';
import { shapeArc, shapeArcAlt } from './path-generators.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * 检查形状是否为杂项形状
 */
export function isMiscShape(shapType) {
    const miscShapes = [
        'smileyFace', 'verticalScroll', 'horizontalScroll'
    ];
    return miscShapes.includes(shapType);
}

/**
 * 渲染杂项形状
 */
export function renderMiscShape(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";
    let dVal = "";

    if (shapType === "smileyFace") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var refr = SLIDE_FACTOR;
        var adj = 4653 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        var cnstVal1 = 50000 * refr;
        var cnstVal2 = 100000 * refr;
        var cnstVal3 = 4653 * refr;
        var ss = Math.min(w, h);
        var a, x1, x2, x3, x4, y1, y3, dy2, y2, y4, dy3, y5, wR, hR, wd2, hd2;
        wd2 = w / 2;
        hd2 = h / 2;
        a = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
        x1 = w * 4969 / 21699;
        x2 = w * 6215 / 21600;
        x3 = w * 13135 / 21600;
        x4 = w * 16640 / 21600;
        y1 = h * 7570 / 21600;
        y3 = h * 16515 / 21600;
        dy2 = h * a / cnstVal2;
        y2 = y3 - dy2;
        y4 = y3 + dy2;
        dy3 = h * a / cnstVal1;
        y5 = y4 + dy3;
        wR = w * 1125 / 21600;
        hR = h * 1125 / 21600;
        var cX1 = x2 - wR * Math.cos(Math.PI);
        var cY1 = y1 - hR * Math.sin(Math.PI);
        var cX2 = x3 - wR * Math.cos(Math.PI);
        dVal = //eyes
            shapeArc(cX1, cY1, wR, hR, 180, 540, false) +
            shapeArc(cX2, cY1, wR, hR, 180, 540, false) +
            //mouth
            " M" + x1 + "," + y2 +
            " Q" + wd2 + "," + y5 + " " + x4 + "," + y2 +
            " Q" + wd2 + "," + y5 + " " + x1 + "," + y2 +
            //head
            " M" + 0 + "," + hd2 +
            shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L") +
            " z";
    }
    else if (shapType === "verticalScroll" || shapType === "horizontalScroll") {
        var shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        var refr = SLIDE_FACTOR;
        var adj = 12500 * refr;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * refr;
        }
        var cnstVal1 = 25000 * refr;
        var cnstVal2 = 100000 * refr;
        var ss = Math.min(w, h);
        var t = 0, l = 0, b = h, r = w;
        var a, ch, ch2, ch4;
        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
        ch = ss * a / cnstVal2;
        ch2 = ch / 2;
        ch4 = ch / 4;
        if (shapType === "verticalScroll") {
            var x3, x4, x6, x7, x5, y3, y4;
            x3 = ch + ch2;
            x4 = ch + ch;
            x6 = r - ch;
            x7 = r - ch2;
            x5 = x6 - ch2;
            y3 = b - ch;
            y4 = b - ch2;

            dVal = "M" + ch + "," + y3 +
                " L" + ch + "," + ch2 +
                shapeArc(x3, ch2, ch2, ch2, 180, 270, false).replace("M", "L") +
                " L" + x7 + "," + t +
                shapeArc(x7, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                " L" + x6 + "," + ch +
                " L" + x6 + "," + y4 +
                shapeArc(x5, y4, ch2, ch2, 0, 90, false).replace("M", "L") +
                " L" + ch2 + "," + b +
                shapeArc(ch2, y4, ch2, ch2, 90, 270, false).replace("M", "L") +
                " z" +
                " M" + x3 + "," + t +
                shapeArc(x3, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                shapeArc(x3, x3 / 2, ch4, ch4, 90, 270, false).replace("M", "L") +
                " L" + x4 + "," + ch2 +
                " M" + x6 + "," + ch +
                " L" + x3 + "," + ch +
                " M" + ch + "," + y4 +
                shapeArc(ch2, y4, ch2, ch2, 0, 270, false).replace("M", "L") +
                shapeArc(ch2, (y4 + y3) / 2, ch4, ch4, 270, 450, false).replace("M", "L") +
                " z" +
                " M" + ch + "," + y4 +
                " L" + ch + "," + y3;
        } else if (shapType === "horizontalScroll") {
            var y3, y4, y6, y7, y5, x3, x4;
            y3 = ch + ch2;
            y4 = ch + ch;
            y6 = b - ch;
            y7 = b - ch2;
            y5 = y6 - ch2;
            x3 = r - ch;
            x4 = r - ch2;

            dVal = "M" + l + "," + y3 +
                shapeArc(ch2, y3, ch2, ch2, 180, 270, false).replace("M", "L") +
                " L" + x3 + "," + ch +
                " L" + x3 + "," + ch2 +
                shapeArc(x4, ch2, ch2, ch2, 180, 360, false).replace("M", "L") +
                " L" + r + "," + y5 +
                shapeArc(x4, y5, ch2, ch2, 0, 90, false).replace("M", "L") +
                " L" + ch + "," + y6 +
                " L" + ch + "," + y7 +
                shapeArc(ch2, y7, ch2, ch2, 0, 180, false).replace("M", "L") +
                " z" +
                "M" + x4 + "," + ch +
                shapeArc(x4, ch2, ch2, ch2, 90, -180, false).replace("M", "L") +
                shapeArc((x3 + x4) / 2, ch2, ch4, ch4, 180, 0, false).replace("M", "L") +
                " z" +
                " M" + x4 + "," + ch +
                " L" + x3 + "," + ch +
                " M" + ch2 + "," + y4 +
                " L" + ch2 + "," + y3 +
                shapeArc(y3 / 2, y3, ch4, ch4, 180, 360, false).replace("M", "L") +
                shapeArc(ch2, y3, ch2, ch2, 0, 180, false).replace("M", "L") +
                " M" + ch + "," + y3 +
                " L" + ch + "," + y6;
        }
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}
