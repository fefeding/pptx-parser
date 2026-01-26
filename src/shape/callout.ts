    import { PPTXUtils } from '../core/utils.js';
    import { PPTXShapeUtils } from './shape.js';
    const PPTXCalloutShapes: any = {};

    // 楔形矩形标注
    PPTXCalloutShapes.genWedgeRectCallout = function(w, h, node, slideFactor) {
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        const refr = slideFactor;
        let sAdj1, adj1 = -20833 * refr;
        let sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        let d_val;
        const cnstVal1 = 100000 * slideFactor;
        const vc: any = h / 2, hc = w / 2;
        let dxPos: any, dyPos: any, xPos: any, yPos: any, dx: any, dy: any, dq: any, ady: any, adq: any, dz: any;
        dxPos = w * adj1 / cnstVal1;
        dyPos = h * adj2 / cnstVal1;
        xPos = hc + dxPos;
        yPos = vc + dyPos;
        dx = xPos - hc;
        dy = yPos - vc;
        dq = dxPos * h / w;
        ady = Math.abs(dyPos);
        adq = Math.abs(dq);
        dz = (ady > adq) ? 1 : 0;
        const xg1: any = hc + dxPos - (dxPos / 2);
        const xg2: any = hc - (dxPos / 2);
        const x1: any = (dz === 0) ? 0 : (dxPos > 0) ? xg1 : xg2;
        const x2: any = (dz === 0) ? 0 : (dxPos > 0) ? xg2 : xg1;
        const yg1: any = vc + dyPos - (dyPos / 2);
        const yg2: any = vc - (dyPos / 2);
        const y1: any = (dz === 1) ? 0 : (dyPos > 0) ? yg1 : yg2;
        const y2: any = (dz === 1) ? 0 : (dyPos > 0) ? yg2 : yg1;
        const t1: any = (dxPos >= 0) ? x1 : x2;
        const xl: any = (t1 < 0) ? 0 : t1;
        const t2: any = (dxPos >= 0) ? x2 : x1;
        const xt: any = (t2 < 0) ? 0 : t2;
        const t3: any = (dxPos >= 0) ? x1 : x2;
        const xr: any = (t3 > w) ? w : t3;
        const t4: any = (dxPos >= 0) ? x2 : x1;
        const xb: any = (t4 > w) ? w : t4;
        const t5: any = (dyPos >= 0) ? y1 : y2;
        const yl: any = (t5 < 0) ? 0 : t5;
        const t6: any = (dyPos >= 0) ? y2 : y1;
        const yt: any = (t6 < 0) ? 0 : t6;
        const t7: any = (dyPos >= 0) ? y1 : y2;
        const yr: any = (t7 > h) ? h : t7;
        const t8: any = (dyPos >= 0) ? y2 : y1;
        const yb: any = (t8 > h) ? h : t8;

        if (dz === 0) {
            d_val = "M" + xl + `,0 L` + xb + `,0 L` + xb + "," + yl +
                " L" + w + "," + yl +
                " L" + w + "," + yr +
                " L" + xb + "," + yr +
                " L" + xb + "," + h +
                " L" + xt + "," + h +
                " L" + xt + "," + yb +
                " L" + 0 + "," + yb +
                " L" + 0 + "," + yt +
                " L" + xt + "," + yt +
                " L" + xt + "," + 0 +
                ` z M` + xPos + "," + yPos +
                " L" + xt + "," + yl +
                " L" + xt + "," + yt +
                " z";
        } else {
            d_val = "M" + 0 + "," + yt +
                " L" + 0 + "," + yb +
                " L" + xl + "," + yb +
                " L" + xl + "," + h +
                " L" + xr + "," + h +
                " L" + xr + "," + yb +
                " L" + w + "," + yb +
                " L" + w + "," + yt +
                " L" + xr + "," + yt +
                " L" + xr + "," + 0 +
                " L" + xb + "," + 0 +
                " L" + xb + "," + yt +
                ` z M` + xPos + "," + yPos +
                " L" + xl + "," + yt +
                " L" + xb + "," + yt +
                " z";
        }
        return d_val;
    };

    // 楔形圆形标注
    PPTXCalloutShapes.genWedgeEllipseCallout = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let refr = slideFactor;
        let sAdj1, adj1 = -20833 * refr;
        let sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        let d_val;
        let cnstVal1 = 100000 * slideFactor;
        const angVal1 = 11 * Math.PI / 180;
        const ss = Math.min(w, h);
        const vc = h / 2, hc = w / 2;
        const dxPos: any = w * adj1 / cnstVal1;
        const dyPos: any = h * adj2 / cnstVal1;
        const xPos: any = hc + dxPos;
        const yPos: any = vc + dyPos;
        const sdx: any = dxPos * h;
        const sdy: any = dyPos * w;
        const pang: any = Math.atan(sdy / sdx);
        const stAng: any = pang - angVal1;
        const enAng: any = pang + angVal1;
        const dx1: any = dxPos * h / ss;
        const dy1: any = dyPos * w / ss;
        const x1: any = hc + dx1;
        const y1: any = vc + dy1;
        const dx2: any = (x1 - hc) * (ss / h);
        const dy2: any = (y1 - vc) * (ss / w);
        const x2: any = hc + dx2;
        const y2: any = vc + dy2;
        const stAng1: any = stAng * 180 / Math.PI;
        const enAng1: any = enAng * 180 / Math.PI;
        const swAng1: any = enAng1 - stAng1;
        const swAng2: any = swAng1 - 360;
        const swAng: any = (swAng1 > 180) ? swAng2 : swAng1;

        const wd2 = ss / 2, hd2 = ss / 2;
        d_val = PPTXShapeUtils.shapeArc(hc, vc, wd2, hd2, stAng1, stAng1 + swAng, false) +
            " L" + xPos + "," + yPos +
            " z";
        return d_val;
    };

    // 云形标注
    PPTXCalloutShapes.genCloudCallout = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let refr = slideFactor;
        let sAdj1, adj1 = -20833 * refr;
        let sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                    let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        let cnstVal1 = 100000 * slideFactor;
        const cnstVal2 = 200000 * slideFactor;
        let vc = h / 2, hc = w / 2;
        let d_val;
        let dxPos = w * adj1 / cnstVal1;
        const dyPos = h * adj2 / cnstVal1;
        const xPos = hc + dxPos;
        const yPos = vc + dyPos;

        // 简化版云形 - 使用多个圆形组合
        const wd4 = w / 4;
        const hd4 = h / 4;
        d_val = "M" + wd4 + `,0 Q` + w / 2 + "," + (hd4 * 0.5) + " " + (w - wd4) + `,0 Q` + w + "," + hd4 + " " + w + "," + (h / 2) +
            " Q" + w + "," + (h - hd4) + " " + (w - wd4) + "," + h +
            " Q" + (w / 2) + "," + h + " " + wd4 + "," + h +
            " Q0," + (h - hd4) + " 0," + (h / 2) +
            " Q0," + hd4 + " " + wd4 + `,0 M` + xPos + "," + yPos +
            " L" + (xPos - 20) + "," + (yPos - 30) +
            " L" + (xPos + 20) + "," + (yPos - 30) +
            " z";
        return d_val;
    };

    // 气泡标注
    PPTXCalloutShapes.genBorderCallout = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let refr = slideFactor;
        let sAdj1, adj1 = -20833 * refr;
        let sAdj2, adj2 = 62500 * refr;
        let sAdj3, adj3 = 12500 * refr;
        let d_val;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                } else if (sAdj_name == "adj3") {
                    sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                }
            }
        }
        let cnstVal1 = 100000 * slideFactor;
        let cnstVal2 = 200000 * slideFactor;
        let vc = h / 2, hc = w / 2;
        let dxPos = w * adj1 / cnstVal1;
        let dyPos = h * adj2 / cnstVal1;
        let xPos = hc + dxPos;
        let yPos = vc + dyPos;
        const r = (adj3 * Math.min(w, h)) / cnstVal2;

        // 带边框的气泡
        d_val = "M" + r + `,0 L` + (w - r) + ",0" +
            PPTXShapeUtils.shapeArc(w - r, r, r, r, 270, 360, false).replace("M", "L") +
            " L" + w + "," + (h - r) +
            PPTXShapeUtils.shapeArc(w - r, h - r, r, r, 0, 90, false).replace("M", "L") +
            " L" + r + "," + h +
            PPTXShapeUtils.shapeArc(r, h - r, r, r, 90, 180, false).replace("M", "L") +
            " L" + 0 + "," + r +
            PPTXShapeUtils.shapeArc(r, r, r, r, 180, 270, false).replace("M", "L") +
            " M" + xPos + "," + yPos +
            " L" + (xPos - 15) + "," + (yPos - 25) +
            " L" + (xPos + 15) + "," + (yPos - 25) +
            " z";
        return d_val;
    };

export { PPTXCalloutShapes };

// Also export to global scope for backward compatibility
// window.PPTXCalloutShapes = PPTXCalloutShapes; // Removed for ES modules
