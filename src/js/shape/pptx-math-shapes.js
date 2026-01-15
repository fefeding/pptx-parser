 const PPTXMathShapes = {};

    // 加号
    PPTXMathShapes.genMathPlus = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 23520 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var cnstVal3 = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key) {
                var name = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        var cnstVal6 = 73490 * slideFactor;
        var ss = Math.min(w, h);
        var a1, dx1, dy1, dx2, x1, x2, x3, x4, y1, y2, y3, y4;

        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal6) ? cnstVal6 : adj1;
        var hc = w / 2, vc = h / 2;
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
    PPTXMathShapes.genMathMinus = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 23520 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var cnstVal3 = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key) {
                var name = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        var cnstVal6 = 73490 * slideFactor;
        var a1, dy1, dx1, y1, y2, x1, x2;
        var hc = w / 2, vc = h / 2;

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
    PPTXMathShapes.genMathMultiply = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 23520 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var cnstVal3 = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key) {
                var name = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                }
            });
        }
        var cnstVal6 = 51965 * slideFactor;
        var ss = Math.min(w, h);
        var a1, th, a, sa, ca, ta, dl, rw, lM, xM, yM, dxAM, dyAM,
            xA, yA, xB, yB, xBC, yBC, yC, xD, xE, yFE, xFE, xF, xL, yG, yH, yI, xC2, yC3;

        var hc = w / 2, vc = h / 2;
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
    PPTXMathShapes.genMathEqual = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, adj1 = 23520 * slideFactor;
        var sAdj2, adj2 = 11760 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var cnstVal3 = 200000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            Object.keys(shapAdjst_ary).forEach(function (key) {
                var name = shapAdjst_ary[key]["attrs"]["name"];
                if (name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                } else if (name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[key], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                }
            });
        }
        var cnstVal5 = 36745 * slideFactor;
        var cnstVal6 = 73490 * slideFactor;
        var a1, a2a1, mAdj2, a2, dy1, dy2, dx1, y2, y3, y1, y4, x1, x2, yC1, yC2;
        var hc = w / 2, vc = h / 2;

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
            " z" +
            "M" + x1 + "," + y3 +
            " L" + x2 + "," + y3 +
            " L" + x2 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
    };

export { PPTXMathShapes };

// Also export to global scope for backward compatibility
window.PPTXMathShapes = PPTXMathShapes;
