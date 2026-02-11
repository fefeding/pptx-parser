/**
 * 数学符号形状渲染模块
 * 提供数学符号（加减乘除等）的生成和渲染功能
 */

import { PPTXXmlUtils } from '../utils/xml.js';
import { shapeArc } from './path-generators.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * 检查形状是否为数学符号形状
 */
export function isMathSymbol(shapType) {
    const mathSymbols = [
        'mathDivide', 'mathEqual', 'mathMinus', 'mathMultiply', 'mathNotEqual', 'mathPlus'
    ];
    return mathSymbols.includes(shapType);
}

/**
 * 渲染数学符号形状
 */
export function renderMathSymbol(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";

    // 获取形状调整参数
    var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    var sAdj1, adj1;
    var sAdj2, adj2;
    var sAdj3, adj3;
    if (shapAdjst_ary !== undefined) {
        if (shapAdjst_ary.constructor === Array) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4));
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4));
                } else if (sAdj_name == "adj3") {
                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4));
                }
            }
        } else {
            sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            adj1 = parseInt(sAdj1.substr(4));
        }
    }
    var cnstVal1 = 50000 * SLIDE_FACTOR;
    var cnstVal2 = 100000 * SLIDE_FACTOR;
    var cnstVal3 = 200000 * SLIDE_FACTOR;
    var dVal;
    var hc = w / 2, vc = h / 2, hd2 = h / 2;

    // mathNotEqual (不等于符号)
    if (shapType == "mathNotEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
            adj2 = 110 * Math.PI / 180;
            adj3 = 11760 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
            adj2 = (adj2 / 60000) * Math.PI / 180;
            adj3 = adj3 * SLIDE_FACTOR;
        }
        var a1, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4,
            cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37,
            x3, dx27, x2, rx7, rx6, rx5, rx4, rx3, rx2, dx7, rxt, lxt, rx, lx,
            dy3, dy4, ry, ly, dlx, drx, dly, dry, xC1, xC2, yC1, yC2, yC3, yC4;
        var angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
        var cnstVal4 = 73490 * SLIDE_FACTOR;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        crAng = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
        a2a1 = a1 * 2;
        maxAdj3 = cnstVal2 - a2a1;
        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a3 / cnstVal3;
        dx1 = w * cnstVal4 / cnstVal3;
        x1 = hc - dx1;
        x8 = hc + dx1;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        cadj2 = crAng - Math.PI / 2;
        xadj2 = hd2 * Math.tan(cadj2);
        len = Math.sqrt(xadj2 * xadj2 + hd2 * hd2);
        bhw = len * dy1 / hd2;
        bhw2 = bhw / 2;
        x7 = hc + xadj2 - bhw2;
        dx67 = xadj2 * y1 / hd2;
        x6 = x7 - dx67;
        dx57 = xadj2 * y2 / hd2;
        x5 = x7 - dx57;
        dx47 = xadj2 * y3 / hd2;
        x4 = x7 - dx47;
        dx37 = xadj2 * y4 / hd2;
        x3 = x7 - dx37;
        dx27 = xadj2 * 2;
        x2 = x7 - dx27;
        rx7 = x7 + bhw;
        rx6 = x6 + bhw;
        rx5 = x5 + bhw;
        rx4 = x4 + bhw;
        rx3 = x3 + bhw;
        rx2 = x2 + bhw;
        dx7 = dy1 * hd2 / len;
        rxt = x7 + dx7;
        lxt = rx7 - dx7;
        rx = (cadj2 > 0) ? rxt : rx7;
        lx = (cadj2 > 0) ? x7 : lxt;
        dy3 = dy1 * xadj2 / len;
        dy4 = -dy3;
        ry = (cadj2 > 0) ? dy3 : 0;
        ly = (cadj2 > 0) ? 0 : dy4;
        dlx = w - rx;
        drx = w - lx;
        dly = h - ry;
        dry = h - ly;
        xC1 = (rx + lx) / 2;
        xC2 = (drx + dlx) / 2;
        yC1 = (ry + ly) / 2;
        yC2 = (y1 + y2) / 2;
        yC3 = (y3 + y4) / 2;
        yC4 = (dry + dly) / 2;

        dVal = "M" + x1 + "," + y1 +
            " L" + x6 + "," + y1 +
            " L" + lx + "," + ly +
            " L" + rx + "," + ry +
            " L" + rx6 + "," + y1 +
            " L" + x8 + "," + y1 +
            " L" + x8 + "," + y2 +
            " L" + rx5 + "," + y2 +
            " L" + rx4 + "," + y3 +
            " L" + x8 + "," + y3 +
            " L" + x8 + "," + y4 +
            " L" + rx3 + "," + y4 +
            " L" + drx + "," + dry +
            " L" + dlx + "," + dly +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " L" + x1 + "," + y3 +
            " L" + x4 + "," + y3 +
            " L" + x5 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";
    } 
    // mathDivide (除号)
    else if (shapType == "mathDivide") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
            adj2 = 5880 * SLIDE_FACTOR;
            adj3 = 11760 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
            adj2 = adj2 * SLIDE_FACTOR;
            adj3 = adj3 * SLIDE_FACTOR;
        }
        var a1, ma1, ma3h, ma3w, maxAdj3, a3, m4a3, maxAdj2, a2, dy1, yg, rad, dx1,
            y3, y4, a, y2, y1, y5, x1, x3, x2;
        var cnstVal4 = 1000 * SLIDE_FACTOR;
        var cnstVal5 = 36745 * SLIDE_FACTOR;
        var cnstVal6 = 73490 * SLIDE_FACTOR;
        a1 = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        ma1 = -a1;
        ma3h = (cnstVal6 + ma1) / 4;
        ma3w = cnstVal5 * w / h;
        maxAdj3 = (ma3h < ma3w) ? ma3h : ma3w;
        a3 = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        m4a3 = -4 * a3;
        maxAdj2 = cnstVal6 + m4a3 - a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
        dy1 = h * a1 / cnstVal3;
        yg = h * a2 / cnstVal2;
        rad = h * a3 / cnstVal2;
        dx1 = w * cnstVal6 / cnstVal3;
        y3 = vc - dy1;
        y4 = vc + dy1;
        a = yg + rad;
        y2 = y3 - a;
        y1 = y2 - rad;
        y5 = h - y1;
        x1 = hc - dx1;
        x3 = hc + dx1;
        x2 = hc - rad;
        var cd4 = 90, c3d4 = 270;
        var cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
        var cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
        var cX2 = hc - Math.cos(Math.PI / 2) * rad;
        var cY2 = y5 - Math.sin(Math.PI / 2) * rad;
            dVal = "M" + hc + "," + y1 +
                shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
                " z" +
                " M" + hc + "," + y5 +
                shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
                " z" +
            " M" + x1 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    } 
    // mathEqual (等号)
    else if (shapType == "mathEqual") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
            adj2 = 11760 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
            adj2 = adj2 * SLIDE_FACTOR;
        }
        var cnstVal5 = 36745 * SLIDE_FACTOR;
        var cnstVal6 = 73490 * SLIDE_FACTOR;
        var a1, a2a1, mAdj2, a2, dy1, dy2, dx1, y2, y3, y1, y4, x1, x2, yC1, yC2;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        a2a1 = a1 * 2;
        mAdj2 = cnstVal2 - a2a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > mAdj2) ? mAdj2 : adj2;
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a2 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y2 = vc - dy2;
        y3 = vc + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;
        yC1 = (y1 + y2) / 2;
        yC2 = (y3 + y4) / 2;
        dVal = "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z" +
            "M" + x1 + "," + y3 +
            " L" + x2 + "," + y3 +
            " L" + x2 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    } 
    // mathMinus (减号)
    else if (shapType == "mathMinus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
        }
        var cnstVal6 = 73490 * SLIDE_FACTOR;
        var a1, dy1, dx1, y1, y2, x1, x2;
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
        dy1 = h * a1 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y1 = vc - dy1;
        y2 = vc + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;

        dVal = "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";
    } 
    // mathMultiply (乘号)
    else if (shapType == "mathMultiply") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
        }
        var cnstVal6 = 51965 * SLIDE_FACTOR;
        var a1, th, a, sa, ca, ta, dl, rw, lM, xM, yM, dxAM, dyAM,
            xA, yA, xB, yB, xBC, yBC, yC, xD, xE, yFE, xFE, xF, xL, yG, yH, yI, xC2, yC3;
        var ss = Math.min(w, h);
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        th = ss * a1 / cnstVal2;
        a = Math.atan(h / w);
        sa = 1 * Math.sin(a);
        ca = 1 * Math.cos(a);
        ta = 1 * Math.tan(a);
        dl = Math.sqrt(w * w + h * h);
        rw = dl * cnstVal6 / cnstVal2;
        lM = dl - rw;
        xM = ca * lM / 2;
        yM = sa * lM / 2;
        dxAM = sa * th / 2;
        dyAM = ca * th / 2;
        xA = xM - dxAM;
        yA = yM + dyAM;
        xB = xM + dxAM;
        yB = yM - dyAM;
        xBC = hc - xB;
        yBC = xBC * ta;
        yC = yBC + yB;
        xD = w - xB;
        xE = w - xA;
        yFE = vc - yA;
        xFE = yFE / ta;
        xF = xE - xFE;
        xL = xA + xFE;
        yG = h - yA;
        yH = h - yB;
        yI = h - yC;
        xC2 = w - xM;
        yC3 = h - yM;

        dVal = "M" + xA + "," + yA +
            " L" + xB + "," + yB +
            " L" + hc + "," + yC +
            " L" + xD + "," + yB +
            " L" + xE + "," + yA +
            " L" + xF + "," + vc +
            " L" + xE + "," + yG +
            " L" + xD + "," + yH +
            " L" + hc + "," + yI +
            " L" + xB + "," + yH +
            " L" + xA + "," + yG +
            " L" + xL + "," + vc +
            " z";
    } 
    // mathPlus (加号)
    else if (shapType == "mathPlus") {
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * SLIDE_FACTOR;
        } else {
            adj1 = adj1 * SLIDE_FACTOR;
        }
        var cnstVal6 = 73490 * SLIDE_FACTOR;
        var ss = Math.min(w, h);
        var a1, dx1, dy1, dx2, x1, x2, x3, x4, y1, y2, y3, y4;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        dx1 = w * cnstVal6 / cnstVal3;
        dy1 = h * cnstVal6 / cnstVal3;
        dx2 = ss * a1 / cnstVal3;
        x1 = hc - dx1;
        x2 = hc - dx2;
        x3 = hc + dx2;
        x4 = hc + dx1;
        y1 = vc - dy1;
        y2 = vc - dx2;
        y3 = vc + dx2;
        y4 = vc + dy1;

        dVal = "M" + x1 + "," + y2 +
            " L" + x2 + "," + y2 +
            " L" + x2 + "," + y1 +
            " L" + x3 + "," + y1 +
            " L" + x3 + "," + y2 +
            " L" + x4 + "," + y2 +
            " L" + x4 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x2 + "," + y4 +
            " L" + x2 + "," + y3 +
            " L" + x1 + "," + y3 +
            " z";
    }

    result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

    return result;
}
