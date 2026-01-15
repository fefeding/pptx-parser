    const PPTXCalloutShapes = {};

    // 楔形矩形标注
    PPTXCalloutShapes.genWedgeRectCallout = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var refr = slideFactor;
        var sAdj1, adj1 = -20833 * refr;
        var sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        var d_val;
        var cnstVal1 = 100000 * slideFactor;
        var dxPos, dyPos, xPos, yPos, dx, dy, dq, ady, adq, dz, xg1, xg2, x1, x2,
            yg1, yg2, y1, y2, t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb,
            vc = h / 2, hc = w / 2;
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
        xg1 = hc + dxPos - (dxPos / 2);
        xg2 = hc - (dxPos / 2);
        x1 = (dz === 0) ? 0 : (dxPos > 0) ? xg1 : xg2;
        x2 = (dz === 0) ? 0 : (dxPos > 0) ? xg2 : xg1;
        yg1 = vc + dyPos - (dyPos / 2);
        yg2 = vc - (dyPos / 2);
        y1 = (dz === 1) ? 0 : (dyPos > 0) ? yg1 : yg2;
        y2 = (dz === 1) ? 0 : (dyPos > 0) ? yg2 : yg1;
        t1 = (dxPos >= 0) ? x1 : x2;
        xl = (t1 < 0) ? 0 : t1;
        t2 = (dxPos >= 0) ? x2 : x1;
        xt = (t2 < 0) ? 0 : t2;
        t3 = (dxPos >= 0) ? x1 : x2;
        xr = (t3 > w) ? w : t3;
        t4 = (dxPos >= 0) ? x2 : x1;
        xb = (t4 > w) ? w : t4;
        t5 = (dyPos >= 0) ? y1 : y2;
        yl = (t5 < 0) ? 0 : t5;
        t6 = (dyPos >= 0) ? y2 : y1;
        yt = (t6 < 0) ? 0 : t6;
        t7 = (dyPos >= 0) ? y1 : y2;
        yr = (t7 > h) ? h : t7;
        t8 = (dyPos >= 0) ? y2 : y1;
        yb = (t8 > h) ? h : t8;

        if (dz === 0) {
            d_val = "M" + xl + ",0" +
                " L" + xb + ",0" +
                " L" + xb + "," + yl +
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
                " z" +
                " M" + xPos + "," + yPos +
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
                " z" +
                " M" + xPos + "," + yPos +
                " L" + xl + "," + yt +
                " L" + xb + "," + yt +
                " z";
        }
        return d_val;
    };

    // 楔形圆形标注
    PPTXCalloutShapes.genWedgeEllipseCallout = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var refr = slideFactor;
        var sAdj1, adj1 = -20833 * refr;
        var sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        var d_val;
        var cnstVal1 = 100000 * slideFactor;
        var angVal1 = 11 * Math.PI / 180;
        var ss = Math.min(w, h);
        var dxPos, dyPos, xPos, yPos, sdx, sdy, pang, stAng, enAng, dx1, dy1, x1, y1, dx2, dy2,
            x2, y2, stAng1, enAng1, swAng1, swAng2, swAng,
            vc = h / 2, hc = w / 2;
        dxPos = w * adj1 / cnstVal1;
        dyPos = h * adj2 / cnstVal1;
        xPos = hc + dxPos;
        yPos = vc + dyPos;
        sdx = dxPos * h;
        sdy = dyPos * w;
        pang = Math.atan(sdy / sdx);
        stAng = pang - angVal1;
        enAng = pang + angVal1;
        dx1 = dxPos * h / ss;
        dy1 = dyPos * w / ss;
        x1 = hc + dx1;
        y1 = vc + dy1;
        dx2 = (x1 - hc) * (ss / h);
        dy2 = (y1 - vc) * (ss / w);
        x2 = hc + dx2;
        y2 = vc + dy2;
        stAng1 = stAng * 180 / Math.PI;
        enAng1 = enAng * 180 / Math.PI;
        swAng1 = enAng1 - stAng1;
        swAng2 = swAng1 - 360;
        swAng = (swAng1 > 180) ? swAng2 : swAng1;

        d_val = window.PPTXShapeUtils.shapeArc(hc, vc, wd2, hd2, stAng1, stAng1 + swAng, false) +
            " L" + xPos + "," + yPos +
            " z";
        return d_val;
    };

    // 云形标注
    PPTXCalloutShapes.genCloudCallout = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var refr = slideFactor;
        var sAdj1, adj1 = -20833 * refr;
        var sAdj2, adj2 = 62500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        var d_val;
        var cnstVal1 = 100000 * slideFactor;
        var cnstVal2 = 200000 * slideFactor;
        var vc = h / 2, hc = w / 2;
        var dxPos = w * adj1 / cnstVal1;
        var dyPos = h * adj2 / cnstVal1;
        var xPos = hc + dxPos;
        var yPos = vc + dyPos;

        // 简化版云形 - 使用多个圆形组合
        var wd4 = w / 4;
        var hd4 = h / 4;
        var d_val = "M" + wd4 + ",0" +
            " Q" + w / 2 + "," + (hd4 * 0.5) + " " + (w - wd4) + ",0" +
            " Q" + w + "," + hd4 + " " + w + "," + (h / 2) +
            " Q" + w + "," + (h - hd4) + " " + (w - wd4) + "," + h +
            " Q" + (w / 2) + "," + h + " " + wd4 + "," + h +
            " Q0," + (h - hd4) + " 0," + (h / 2) +
            " Q0," + hd4 + " " + wd4 + ",0" +
            " M" + xPos + "," + yPos +
            " L" + (xPos - 20) + "," + (yPos - 30) +
            " L" + (xPos + 20) + "," + (yPos - 30) +
            " z";
        return d_val;
    };

    // 气泡标注
    PPTXCalloutShapes.genBorderCallout = function(w, h, node, slideFactor) {
        var shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var refr = slideFactor;
        var sAdj1, adj1 = -20833 * refr;
        var sAdj2, adj2 = 62500 * refr;
        var sAdj3, adj3 = 12500 * refr;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
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
        var d_val;
        var cnstVal1 = 100000 * slideFactor;
        var cnstVal2 = 200000 * slideFactor;
        var vc = h / 2, hc = w / 2;
        var dxPos = w * adj1 / cnstVal1;
        var dyPos = h * adj2 / cnstVal1;
        var xPos = hc + dxPos;
        var yPos = vc + dyPos;
        var r = (adj3 * Math.min(w, h)) / cnstVal2;

        // 带边框的气泡
        d_val = "M" + r + ",0" +
            " L" + (w - r) + ",0" +
            window.PPTXShapeUtils.shapeArc(w - r, r, r, r, 270, 360, false).replace("M", "L") +
            " L" + w + "," + (h - r) +
            window.PPTXShapeUtils.shapeArc(w - r, h - r, r, r, 0, 90, false).replace("M", "L") +
            " L" + r + "," + h +
            window.PPTXShapeUtils.shapeArc(r, h - r, r, r, 90, 180, false).replace("M", "L") +
            " L" + 0 + "," + r +
            window.PPTXShapeUtils.shapeArc(r, r, r, r, 180, 270, false).replace("M", "L") +
            " M" + xPos + "," + yPos +
            " L" + (xPos - 15) + "," + (yPos - 25) +
            " L" + (xPos + 15) + "," + (yPos - 25) +
            " z";
        return d_val;
    };

export { PPTXCalloutShapes };

// Also export to global scope for backward compatibility
// window.PPTXCalloutShapes = PPTXCalloutShapes; // Removed for ES modules
