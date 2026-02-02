import { PPTXUtils } from '../core/utils.js';
import { PPTXShapeUtils } from './shape.js';

interface PPTXMathShapesType {
    genMathPlus: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathMinus: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathMultiply: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathEqual: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathDivide: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathNotEqual: (w: number, h: number, node: any, slideFactor: number) => string;
}

const PPTXMathShapes = {} as PPTXMathShapesType;

    // 加号
    PPTXMathShapes.genMathPlus = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 23520 * slideFactor;
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key: string) {
                const name: string = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        const cnstVal6: number = 73490 * slideFactor;
        const ss: number = Math.min(w, h);
        let a1: number, dx1: number, dy1: number, dx2: number, x1: number, x2: number, x3: number, x4: number, y1: number, y2: number, y3: number, y4: number;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        const hc: number = w / 2, vc: number = h / 2;
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

        return "M" + x1 + "," + y2 +
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
    };

    // 减号
    PPTXMathShapes.genMathMinus = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 23520 * slideFactor;
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key: string) {
                const name: string = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        const cnstVal6: number = 73490 * slideFactor;
        let a1: number, dy1: number, dx1: number, y1: number, y2: number, x1: number, x2: number;
        const hc: number = w / 2, vc: number = h / 2;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
        dy1 = h * a1 / cnstVal3;
        dx1 = w * cnstVal6 / cnstVal3;
        y1 = vc - dy1;
        y2 = vc + dy1;
        x1 = hc - dx1;
        x2 = hc + dx1;

        return "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";
    };

    // 乘号
    PPTXMathShapes.genMathMultiply = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 23520 * slideFactor;
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key: string) {
                const name: string = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        const cnstVal6: number = 51965 * slideFactor;
        const ss: number = Math.min(w, h);
        let a1: number, th: number, a: number, sa: number, ca: number, ta: number, dl: number, rw: number, lM: number, xM: number, yM: number, dxAM: number, dyAM: number,
            xA: number, yA: number, xB: number, yB: number, xBC: number, yBC: number, yC: number, xD: number, xE: number, yFE: number, xFE: number, xF: number, xL: number, yG: number, yH: number, yI: number, xC2: number, yC3: number;

        const hc: number = w / 2, vc: number = h / 2;
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

        return "M" + xA + "," + yA +
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
    };

    // 等号
    PPTXMathShapes.genMathEqual = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 23520 * slideFactor;
        let sAdj2: any, adj2: number = 11760 * slideFactor;
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key: string) {
                const name: string = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                } else if (name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                }
            });
        }
        const cnstVal5: number = 36745 * slideFactor;
        const cnstVal6: number = 73490 * slideFactor;
        let a1: number, a2a1: number, mAdj2: number, a2: number, dy1: number, dy2: number, dx1: number, y2: number, y3: number, y1: number, y4: number, x1: number, x2: number, yC1: number, yC2: number;
        const hc: number = w / 2, vc: number = h / 2;

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

        return "M" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            ` zM` + x1 + "," + y3 +
            " L" + x2 + "," + y3 +
            " L" + x2 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    };

    // 除号
    PPTXMathShapes.genMathDivide = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 23520 * slideFactor;
        let sAdj2: any, adj2: number = 5880 * slideFactor;
        let sAdj3: any, adj3: number = 11760 * slideFactor;
        
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        
        if (shapAdjst_ary !== undefined) {
            if (Array.isArray(shapAdjst_ary)) {
                for (let i = 0; i < shapAdjst_ary.length; i++) {
                    const sAdj_name: string = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                    if (sAdj_name == "adj1") {
                        sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                    } else if (sAdj_name == "adj2") {
                        sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                    } else if (sAdj_name == "adj3") {
                        sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                    }
                }
            } else {
                sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
            }
        }
        
        const cnstVal4: number = 1000 * slideFactor;
        const cnstVal5: number = 36745 * slideFactor;
        const cnstVal6: number = 73490 * slideFactor;
        
        let a1: number, ma1: number, ma3h: number, ma3w: number, maxAdj3: number, a3: number, m4a3: number, maxAdj2: number, a2: number, dy1: number, yg: number, rad: number, dx1: number;
        let y3: number, y4: number, a: number, y2: number, y1: number, y5: number, x1: number, x3: number, x2: number;
        
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
        
        const hc: number = w / 2, vc: number = h / 2;
        y3 = vc - dy1;
        y4 = vc + dy1;
        a = yg + rad;
        y2 = y3 - a;
        y1 = y2 - rad;
        y5 = h - y1;
        x1 = hc - dx1;
        x3 = hc + dx1;
        x2 = hc - rad;
        
        const cd4 = 90, c3d4 = 270;
        const cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
        const cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
        const cX2 = hc - Math.cos(Math.PI / 2) * rad;
        const cY2 = y5 - Math.sin(Math.PI / 2) * rad;
        
        return "M" + hc + "," + y1 +
            PPTXShapeUtils.shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
            ` z M` + hc + "," + y5 +
            PPTXShapeUtils.shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
            ` z M` + x1 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    };
    
    // 不等于号
    PPTXMathShapes.genMathNotEqual = function(w: number, h: number, node: any, slideFactor: number): string {
        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1: any, adj1: number = 0;
        let sAdj2: any, adj2: number = 0;
        let sAdj3: any, adj3: number = 0;
        
        const cnstVal1: number = 50000 * slideFactor;
        const cnstVal2: number = 100000 * slideFactor;
        const cnstVal3: number = 200000 * slideFactor;
        
        if (shapAdjst_ary !== undefined) {
            if (Array.isArray(shapAdjst_ary)) {
                for (let i = 0; i < shapAdjst_ary.length; i++) {
                    const sAdj_name: string = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                    if (sAdj_name == "adj1") {
                        sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj1 = parseInt(sAdj1.substr(4));
                    } else if (sAdj_name == "adj2") {
                        sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj2 = parseInt(sAdj2.substr(4));
                    } else if (sAdj_name == "adj3") {
                        sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                        adj3 = parseInt(sAdj3.substr(4));
                    }
                }
            } else {
                sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                adj1 = parseInt(sAdj1.substr(4));
            }
        }
        
        const cnstVal4: number = 73490 * slideFactor;
        const angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
        
        let a1_not_equal: number, crAng_not_equal: number, a2a1_not_equal: number, maxAdj3_not_equal: number, a3_not_equal: number, dy1_not_equal: number, dy2_not_equal: number, dx1_not_equal: number, x1_not_equal: number, x8_not_equal: number, y2_not_equal: number, y3_not_equal: number, y1_not_equal: number, y4_not_equal: number;
        let cadj2_not_equal: number, xadj2_not_equal: number, len_not_equal: number, bhw_not_equal: number, bhw2_not_equal: number, x7_not_equal: number, dx67_not_equal: number, x6_not_equal: number, dx57_not_equal: number, x5_not_equal: number, dx47_not_equal: number, x4_not_equal: number, dx37_not_equal: number;
        let x3_not_equal: number, dx27_not_equal: number, x2_not_equal: number, rx7_not_equal: number, rx6_not_equal: number, rx5_not_equal: number, rx4_not_equal: number, rx3_not_equal: number, rx2_not_equal: number, dx7_not_equal: number, rxt_not_equal: number, lxt_not_equal: number, rx_not_equal: number, lx_not_equal: number;
        let dy3_not_equal: number, dy4_not_equal: number, ry_not_equal: number, ly_not_equal: number, dlx_not_equal: number, drx_not_equal: number, dly_not_equal: number, dry_not_equal: number, xC1_not_equal: number, xC2_not_equal: number, yC1_not_equal: number, yC2_not_equal: number, yC3_not_equal: number, yC4_not_equal: number;
        
        if (shapAdjst_ary === undefined) {
            adj1 = 23520 * slideFactor;
            adj2 = 110 * Math.PI / 180;
            adj3 = 11760 * slideFactor;
        } else {
            adj1 = adj1 * slideFactor;
            adj2 = (adj2 / 60000) * Math.PI / 180;
            adj3 = adj3 * slideFactor;
        }
        
        const hc: number = w / 2, vc: number = h / 2, hd2: number = h / 2;
        
        a1_not_equal = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        crAng_not_equal = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
        a2a1_not_equal = a1_not_equal * 2;
        maxAdj3_not_equal = cnstVal2 - a2a1_not_equal;
        a3_not_equal = (adj3 < 0) ? 0 : (adj3 > maxAdj3_not_equal) ? maxAdj3_not_equal : adj3;
        
        dy1_not_equal = h * a1_not_equal / cnstVal2;
        dy2_not_equal = h * a3_not_equal / cnstVal3;
        dx1_not_equal = w * cnstVal4 / cnstVal3;
        
        x1_not_equal = hc - dx1_not_equal;
        x8_not_equal = hc + dx1_not_equal;
        y2_not_equal = vc - dy2_not_equal;
        y3_not_equal = vc + dy2_not_equal;
        y1_not_equal = y2_not_equal - dy1_not_equal;
        y4_not_equal = y3_not_equal + dy1_not_equal;
        
        cadj2_not_equal = crAng_not_equal - Math.PI / 2;
        xadj2_not_equal = hd2 * Math.tan(cadj2_not_equal);
        len_not_equal = Math.sqrt(xadj2_not_equal * xadj2_not_equal + hd2 * hd2);
        bhw_not_equal = len_not_equal * dy1_not_equal / hd2;
        bhw2_not_equal = bhw_not_equal / 2;
        
        x7_not_equal = hc + xadj2_not_equal - bhw2_not_equal;
        dx67_not_equal = xadj2_not_equal * y1_not_equal / hd2;
        x6_not_equal = x7_not_equal - dx67_not_equal;
        dx57_not_equal = xadj2_not_equal * y2_not_equal / hd2;
        x5_not_equal = x7_not_equal - dx57_not_equal;
        dx47_not_equal = xadj2_not_equal * y3_not_equal / hd2;
        x4_not_equal = x7_not_equal - dx47_not_equal;
        dx37_not_equal = xadj2_not_equal * y4_not_equal / hd2;
        x3_not_equal = x7_not_equal - dx37_not_equal;
        dx27_not_equal = xadj2_not_equal * 2;
        x2_not_equal = x7_not_equal - dx27_not_equal;
        
        rx7_not_equal = x7_not_equal + bhw_not_equal;
        rx6_not_equal = x6_not_equal + bhw_not_equal;
        rx5_not_equal = x5_not_equal + bhw_not_equal;
        rx4_not_equal = x4_not_equal + bhw_not_equal;
        rx3_not_equal = x3_not_equal + bhw_not_equal;
        rx2_not_equal = x2_not_equal + bhw_not_equal;
        
        dx7_not_equal = dy1_not_equal * hd2 / len_not_equal;
        rxt_not_equal = x7_not_equal + dx7_not_equal;
        lxt_not_equal = rx7_not_equal - dx7_not_equal;
        
        rx_not_equal = (cadj2_not_equal > 0) ? rxt_not_equal : rx7_not_equal;
        lx_not_equal = (cadj2_not_equal > 0) ? x7_not_equal : lxt_not_equal;
        
        dy3_not_equal = dy1_not_equal * xadj2_not_equal / len_not_equal;
        dy4_not_equal = -dy3_not_equal;
        
        ry_not_equal = (cadj2_not_equal > 0) ? dy3_not_equal : 0;
        ly_not_equal = (cadj2_not_equal > 0) ? 0 : dy4_not_equal;
        
        dlx_not_equal = w - rx_not_equal;
        drx_not_equal = w - lx_not_equal;
        dly_not_equal = h - ry_not_equal;
        dry_not_equal = h - ly_not_equal;
        
        xC1_not_equal = (rx_not_equal + lx_not_equal) / 2;
        xC2_not_equal = (drx_not_equal + dlx_not_equal) / 2;
        yC1_not_equal = (ry_not_equal + ly_not_equal) / 2;
        yC2_not_equal = (y1_not_equal + y2_not_equal) / 2;
        yC3_not_equal = (y3_not_equal + y4_not_equal) / 2;
        yC4_not_equal = (dry_not_equal + dly_not_equal) / 2;
        
        return "M" + x1_not_equal + "," + y1_not_equal +
            " L" + x6_not_equal + "," + y1_not_equal +
            " L" + lx_not_equal + "," + ly_not_equal +
            " L" + rx_not_equal + "," + ry_not_equal +
            " L" + rx6_not_equal + "," + y1_not_equal +
            " L" + x8_not_equal + "," + y1_not_equal +
            " L" + x8_not_equal + "," + y2_not_equal +
            " L" + rx5_not_equal + "," + y2_not_equal +
            " L" + rx4_not_equal + "," + y3_not_equal +
            " L" + x8_not_equal + "," + y3_not_equal +
            " L" + x8_not_equal + "," + y4_not_equal +
            " L" + rx3_not_equal + "," + y4_not_equal +
            " L" + drx_not_equal + "," + dry_not_equal +
            " L" + dlx_not_equal + "," + dly_not_equal +
            " L" + x3_not_equal + "," + y4_not_equal +
            " L" + x1_not_equal + "," + y4_not_equal +
            " L" + x1_not_equal + "," + y3_not_equal +
            " L" + x4_not_equal + "," + y3_not_equal +
            " L" + x5_not_equal + "," + y2_not_equal +
            " L" + x1_not_equal + "," + y2_not_equal +
            " z";
    };

export { PPTXMathShapes };