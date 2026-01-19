import { PPTXUtils } from '../core/utils.js';

interface PPTXMathShapesType {
    genMathPlus: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathMinus: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathMultiply: (w: number, h: number, node: any, slideFactor: number) => string;
    genMathEqual: (w: number, h: number, node: any, slideFactor: number) => string;
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

export { PPTXMathShapes };