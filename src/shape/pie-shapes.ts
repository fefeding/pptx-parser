/**
 * 饼图/弧形形状渲染模块
 * 提供饼图、弧形、扇形、弦形等形状的生成和渲染功能
 */

import { PPTXXmlUtils } from '../utils/xml';
import { shapePie, shapeArc } from './path-generators';
import { SLIDE_FACTOR } from '../core/constants';
import type { XmlNode } from '../core/types';

/** 形状描边信息 */
export interface ShapeBorder {
    color: string;
    width: string;
    strokeDasharray: string;
}

/**
 * 检查形状是否为饼图/弧形形状
 */
export function isPieShape(shapType: string): boolean {
    const pieShapes = [
        'pie', 'pieWedge', 'arc', 'chord', 'blockArc'
    ];
    return pieShapes.includes(shapType);
}

/**
 * 渲染饼图/弧形形状
 */
export function renderPieShape(shapType: string, w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, fillColor: string, border: ShapeBorder, shpId: string, node: XmlNode, oShadowSvgUrlStr: string): string {
    let result = "";
    let dVal = "";

    if (shapType === "pie" || shapType === "pieWedge" || shapType === "arc") {
        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let adj1 = 0, adj2 = 0, H = 0, shapAdjst1: string | undefined, shapAdjst2: string | undefined, isClose = false;
        if (shapType === "pie") {
            adj1 = 0;
            adj2 = 270;
            H = h;
            isClose = true;
        } else if (shapType === "pieWedge") {
            adj1 = 180;
            adj2 = 270;
            H = 2 * h;
            isClose = true;
        } else if (shapType === "arc") {
            adj1 = 270;
            adj2 = 0;
            H = h;
            isClose = false;
        }
        if (shapAdjst !== undefined) {
            shapAdjst1 = PPTXXmlUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
            shapAdjst2 = shapAdjst1;
            if (shapAdjst1 === undefined) {
                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
            }
            if (shapAdjst1 !== undefined) {
                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
            }
            if (shapAdjst2 !== undefined) {
                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
            }
        }
        const pieVals = shapePie(H, w, adj1, adj2, isClose);
        result += `<path d='${pieVals[0]}' transform='${pieVals[1]}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }
    else if (shapType === "chord") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: string | undefined, sAdj1_val = 45;
        let sAdj2: string | undefined, sAdj2_val = 270;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    const fmla = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1 = fmla;
                    sAdj1_val = parseInt(fmla.substr(4)) / 60000;
                } else if (sAdj_name === "adj2") {
                    const fmla = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj2 = fmla;
                    sAdj2_val = parseInt(fmla.substr(4)) / 60000;
                }
            }
        }
        const hR = h / 2;
        const wR = w / 2;
        dVal = shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
        result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }
    else if (shapType === "blockArc") {
        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: string | undefined, adj1 = 180;
        let sAdj2: string | undefined, adj2 = 0;
        let sAdj3: string | undefined, adj3 = 25000 * SLIDE_FACTOR;
        const cnstVal1 = 50000 * SLIDE_FACTOR;
        const cnstVal2 = 100000 * SLIDE_FACTOR;
        if (shapAdjst_ary !== undefined) {
            for (const i of shapAdjst_ary.keys()) {
                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    const fmla = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1 = fmla;
                    adj1 = parseInt(fmla.substr(4)) / 60000;
                } else if (sAdj_name === "adj2") {
                    const fmla = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj2 = fmla;
                    adj2 = parseInt(fmla.substr(4)) / 60000;
                } else if (sAdj_name === "adj3") {
                    const fmla = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj3 = fmla;
                    adj3 = parseInt(fmla.substr(4)) * SLIDE_FACTOR;
                }
            }
        }

        let stAng = 0, istAng = 0, a3 = 0, sw11 = 0, sw12 = 0, swAng = 0, iswAng = 0;
        const cd1 = 360;
        if (adj1 < 0) stAng = 0;
        else if (adj1 > cd1) stAng = cd1;
        else stAng = adj1;

        if (adj2 < 0) istAng = 0;
        else if (adj2 > cd1) istAng = cd1;
        else istAng = adj2;

        if (adj3 < 0) a3 = 0;
        else if (adj3 > cnstVal1) a3 = cnstVal1;
        else a3 = adj3;

        sw11 = istAng - stAng;
        sw12 = sw11 + cd1;
        swAng = (sw11 > 0) ? sw11 : sw12;
        iswAng = -swAng;

        const endAng = stAng + swAng;
        const iendAng = istAng + iswAng;

        let wt1 = 0, ht1 = 0, dx1 = 0, dy1 = 0, x1 = 0, y1 = 0, stRd = 0, istRd = 0;
        let wd2 = w / 2, hd2 = h / 2, hc = w / 2, vc = h / 2;
        stRd = stAng * (Math.PI) / 180;
        istRd = istAng * (Math.PI) / 180;
        if (stAng > 90 && stAng < 270) {
            wt1 = wd2 * (Math.sin((Math.PI) / 2 - stRd));
            ht1 = hd2 * (Math.cos((Math.PI) / 2 - stRd));

            dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
            dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));

            x1 = hc - dx1;
            y1 = vc - dy1;
        } else {
            wt1 = wd2 * (Math.sin(stRd));
            ht1 = hd2 * (Math.cos(stRd));

            dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
            dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));

            x1 = hc + dx1;
            y1 = vc + dy1;
        }
        let dr = 0, iwd2 = 0, ihd2 = 0, wt2 = 0, ht2 = 0, dx2 = 0, dy2 = 0, x2 = 0, y2 = 0;
        dr = Math.min(w, h) * a3 / cnstVal2;
        iwd2 = wd2 - dr;
        ihd2 = hd2 - dr;
        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
            wt2 = iwd2 * (Math.sin(istRd));
            ht2 = ihd2 * (Math.cos(istRd));
            dx2 = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
            dy2 = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
            x2 = hc + dx2;
            y2 = vc + dy2;
        } else {
            wt2 = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
            ht2 = ihd2 * (Math.cos((Math.PI) / 2 - istRd));

            dx2 = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
            dy2 = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
            x2 = hc - dx2;
            y2 = vc - dy2;
        }
        dVal = `M${x1},${y1}${shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L")} L${x2},${y2}${shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L")} z`;
        result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${(oShadowSvgUrlStr || "")} />`;
    }

    return result;
}
