/**
 * 特殊形状生成器
 */

var PPTXSpecialShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成饼图、扇形和圆弧
     */
    function generatePie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, shapAdjst) {
        border = ensureBorder(border);
        var adj1, adj2, H, isClose;
        if (shapType == "pie") {
            adj1 = 0;
            adj2 = 270;
            H = h;
            isClose = true;
        } else if (shapType == "pieWedge") {
            adj1 = 180;
            adj2 = 270;
            H = 2 * h;
            isClose = true;
        } else if (shapType == "arc") {
            adj1 = 270;
            adj2 = 0;
            H = h;
            isClose = false;
        }
        if (shapAdjst !== undefined) {
            var shapAdjst1 = undefined;
            var shapAdjst2 = undefined;
            if (shapAdjst["attrs"] !== undefined) {
                shapAdjst1 = shapAdjst["attrs"]["fmla"];
                shapAdjst2 = shapAdjst1;
            }
            if (shapAdjst1 === undefined && shapAdjst[0] !== undefined && shapAdjst[0]["attrs"] !== undefined) {
                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                if (shapAdjst[1] !== undefined && shapAdjst[1]["attrs"] !== undefined) {
                    shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
                }
            }
            if (shapAdjst1 !== undefined) {
                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
            }
            if (shapAdjst2 !== undefined) {
                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
            }
        }
        var pieVals = PPTXBaseShapes.shapePie(H, w, adj1, adj2, isClose);
        return "<path   d='" + pieVals[0] + "' transform='" + pieVals[1] + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成弦形
     */
    function generateChord(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var sAdj1_val = 45;
        var sAdj2_val = 270;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            sAdj1_val = parseInt(sAdj1.substr(4)) / 60000;
                        }
                    } else if (sAdj_name == "adj2") {
                        var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            sAdj2_val = parseInt(sAdj2.substr(4)) / 60000;
                        }
                    }
                }
            }
        }
        var hR = h / 2;
        var wR = w / 2;
        var d_val = PPTXShapeUtils.shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
        return "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成框架形状
     */
    function generateFrame(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj1 = 12500 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj1 = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a1, x1, x4, y4;
        if (adj1 < 0) a1 = 0
        else if (adj1 > cnstVal1) a1 = cnstVal1
        else a1 = adj1
        x1 = Math.min(w, h) * a1 / cnstVal2;
        x4 = w - x1;
        y4 = h - x1;
        var d = "M" + 0 + "," + 0 +
            " L" + w + "," + 0 +
            " L" + w + "," + h +
            " L" + 0 + "," + h +
            " z" +
            "M" + x1 + "," + x1 +
            " L" + x1 + "," + y4 +
            " L" + x4 + "," + y4 +
            " L" + x4 + "," + x1 +
            " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成环形
     */
    function generateDonut(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 25000 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, dr, iwd2, ihd2;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        var d = "M" + 0 + "," + h / 2 +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
            " z" +
            "M" + dr + "," + h / 2 +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L") +
            " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成禁止吸烟标志
     */
    function generateNoSmoking(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 18750 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, dr, iwd2, ihd2, ang, ct, st, m, n, drd2, dang, dang2, swAng, stAng1, stAng2;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        ang = Math.atan(h / w);
        ct = ihd2 * Math.cos(ang);
        st = iwd2 * Math.sin(ang);
        m = Math.sqrt(ct * ct + st * st);
        n = iwd2 * ihd2 / m;
        drd2 = dr / 2;
        dang = Math.atan(drd2 / n);
        dang2 = dang * 2;
        swAng = -Math.PI + dang2;
        stAng1 = ang - dang;
        stAng2 = stAng1 - Math.PI;
        var ct1, st1, m1, n1, dx1, dy1, x1, y1, x2, y2;
        ct1 = ihd2 * Math.cos(stAng1);
        st1 = iwd2 * Math.sin(stAng1);
        m1 = Math.sqrt(ct1 * ct1 + st1 * st1);
        n1 = iwd2 * ihd2 / m1;
        dx1 = n1 * Math.cos(stAng1);
        dy1 = n1 * Math.sin(stAng1);
        x1 = w / 2 + dx1;
        y1 = h / 2 + dy1;
        x2 = w / 2 - dx1;
        y2 = h / 2 - dy1;
        var stAng1deg = stAng1 * 180 / Math.PI;
        var stAng2deg = stAng2 * 180 / Math.PI;
        var swAng2deg = swAng * 180 / Math.PI;
        var d = "M" + 0 + "," + h / 2 +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
            " z" +
            "M" + x1 + "," + y1 +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
            " z" +
            "M" + x2 + "," + y2 +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
            " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成半框架
     */
    function generateHalfFrame(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var sAdj1_val = 3.5;
        var sAdj2_val = 3.5;
        var cnsVal = 100000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            sAdj1_val = parseInt(sAdj1.substr(4)) * slideFactor;
                        }
                    } else if (sAdj_name == "adj2") {
                        var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            sAdj2_val = parseInt(sAdj2.substr(4)) * slideFactor;
                        }
                    }
                }
            }
        }
        var minWH = Math.min(w, h);
        var maxAdj2 = (cnsVal * w) / minWH;
        var a1, a2;
        if (sAdj2_val < 0) a2 = 0
        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
        else a2 = sAdj2_val
        var x1 = (minWH * a2) / cnsVal;
        var g1 = h * x1 / w;
        var g2 = h - g1;
        var maxAdj1 = (cnsVal * g2) / minWH;
        if (sAdj1_val < 0) a1 = 0
        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
        else a1 = sAdj1_val
        var y1 = minWH * a1 / cnsVal;
        var dx2 = y1 * w / h;
        var x2 = w - dx2;
        var dy2 = x1 * h / w;
        var y2 = h - dy2;
        var d = "M0,0" +
            " L" + w + "," + 0 +
            " L" + x2 + "," + y1 +
            " L" + x1 + "," + y1 +
            " L" + x1 + "," + y2 +
            " L0," + h + " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成块状圆弧
     */
    function generateBlockArc(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj1 = 180;
        var adj2 = 0;
        var adj3 = 25000 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            adj1 = parseInt(sAdj1.substr(4)) / 60000;
                        }
                    } else if (sAdj_name == "adj2") {
                        var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            adj2 = parseInt(sAdj2.substr(4)) / 60000;
                        }
                    } else if (sAdj_name == "adj3") {
                        var sAdj3 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj3 !== undefined) {
                            adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                        }
                    }
                }
            }
        }
        var cd1 = 360;
        var stAng = (adj1 < 0) ? 0 : (adj1 > cd1) ? cd1 : adj1;
        var istAng = (adj2 < 0) ? 0 : (adj2 > cd1) ? cd1 : adj2;
        var a3 = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
        var sw11 = istAng - stAng;
        var sw12 = sw11 + cd1;
        var swAng = (sw11 > 0) ? sw11 : sw12;
        var iswAng = -swAng;
        var endAng = stAng + swAng;
        var iendAng = istAng + iswAng;
        var stRd = stAng * (Math.PI) / 180;
        var istRd = istAng * (Math.PI) / 180;
        var wd2 = w / 2;
        var hd2 = h / 2;
        var hc = w / 2;
        var vc = h / 2;
        var x1, y1;
        if (stAng > 90 && stAng < 270) {
            var wt1 = wd2 * (Math.sin((Math.PI) / 2 - stRd));
            var ht1 = hd2 * (Math.cos((Math.PI) / 2 - stRd));
            var dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
            var dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));
            x1 = hc - dx1;
            y1 = vc - dy1;
        } else {
            var wt1 = wd2 * (Math.sin(stRd));
            var ht1 = hd2 * (Math.cos(stRd));
            var dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
            var dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));
            x1 = hc + dx1;
            y1 = vc + dy1;
        }
        var dr = Math.min(w, h) * a3 / cnstVal2;
        var iwd2 = wd2 - dr;
        var ihd2 = hd2 - dr;
        var x2, y2;
        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
            var wt2 = iwd2 * (Math.sin(istRd));
            var ht2 = ihd2 * (Math.cos(istRd));
            var dx2 = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
            var dy2 = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
            x2 = hc + dx2;
            y2 = vc + dy2;
        } else {
            var wt2 = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
            var ht2 = ihd2 * (Math.cos((Math.PI) / 2 - istRd));
            var dx2 = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
            var dy2 = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
            x2 = hc - dx2;
            y2 = vc - dy2;
        }
        var d = "M" + x1 + "," + y1 +
            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L") +
            " L" + x2 + "," + y2 +
            PPTXShapeUtils.shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L") +
            " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成大括号对
     */
    function generateBracePair(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 8333 * slideFactor;
        var cnstVal1 = 25000 * slideFactor;
        var cnstVal2 = 50000 * slideFactor;
        var cnstVal3 = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, x1, x2, x3, x4, y2, y3, y4;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        var minWH = Math.min(w, h);
        x1 = minWH * a / cnstVal3;
        x2 = minWH * a / cnstVal2;
        x3 = w - x2;
        x4 = w - x1;
        var vc = h / 2;
        y2 = vc - x1;
        y3 = vc + x1;
        y4 = h - x1;
        var d = "M" + x2 + "," + h +
            PPTXShapeUtils.shapeArc(x2, y4, x1, x1, 90, 180, false).replace("M", "L") +
            " L" + x1 + "," + y3 +
            PPTXShapeUtils.shapeArc(0, y3, x1, x1, 0, (-90), false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(0, y2, x1, x1, 90, 0, false).replace("M", "L") +
            " L" + x1 + "," + x1 +
            PPTXShapeUtils.shapeArc(x2, x1, x1, x1, 180, 270, false).replace("M", "L") +
            " M" + x3 + "," + 0 +
            PPTXShapeUtils.shapeArc(x3, x1, x1, x1, 270, 360, false).replace("M", "L") +
            " L" + x4 + "," + y2 +
            PPTXShapeUtils.shapeArc(w, y2, x1, x1, 180, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w, y3, x1, x1, 270, 180, false).replace("M", "L") +
            " L" + x4 + "," + y4 +
            PPTXShapeUtils.shapeArc(x3, y4, x1, x1, 0, 90, false).replace("M", "L");
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成左大括号
     */
    function generateLeftBrace(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj1 = 8333 * slideFactor;
        var adj2 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                        }
                    } else if (sAdj_name == "adj2") {
                        var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                        }
                    }
                }
            }
        }
        var a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0
        else if (adj2 > cnstVal2) a2 = cnstVal2
        else a2 = adj2
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        q2 = (q1 < a2) ? q1 : a2;
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0
        else if (adj1 > maxAdj1) a1 = maxAdj1
        else a1 = adj1
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = y3 + y1;
        var d = "M" + w + "," + h +
            PPTXShapeUtils.shapeArc(w, h - y1, w / 2, y1, 90, 180, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, (-90), false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(0, y2, w / 2, y1, 90, 0, false).replace("M", "L") +
            " L" + w / 2 + "," + y1 +
            PPTXShapeUtils.shapeArc(w, y1, w / 2, y1, 180, 270, false).replace("M", "L");
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成右大括号
     */
    function generateRightBrace(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj1 = 8333 * slideFactor;
        var adj2 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                        }
                    } else if (sAdj_name == "adj2") {
                        var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                        }
                    }
                }
            }
        }
        var a1, a2, q1, q2, q3, y1, y2, y3, y4;
        if (adj2 < 0) a2 = 0
        else if (adj2 > cnstVal2) a2 = cnstVal2
        else a2 = adj2
        var minWH = Math.min(w, h);
        q1 = cnstVal2 - a2;
        q2 = (q1 < a2) ? q1 : a2;
        q3 = q2 / 2;
        var maxAdj1 = q3 * h / minWH;
        if (adj1 < 0) a1 = 0
        else if (adj1 > maxAdj1) a1 = maxAdj1
        else a1 = adj1
        y1 = minWH * a1 / cnstVal2;
        y3 = h * a2 / cnstVal2;
        y2 = y3 - y1;
        y4 = h - y1;
        var d = "M" + 0 + "," + 0 +
            PPTXShapeUtils.shapeArc(0, y1, w / 2, y1, 270, 360, false).replace("M", "L") +
            " L" + w / 2 + "," + y2 +
            PPTXShapeUtils.shapeArc(w, y2, w / 2, y1, 180, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w, y3 + y1, w / 2, y1, 270, 180, false).replace("M", "L") +
            " L" + w / 2 + "," + y4 +
            PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, 90, false).replace("M", "L");
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成方括号对
     */
    function generateBracketPair(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 16667 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, x1, x2, y2;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        x1 = Math.min(w, h) * a / cnstVal2;
        x2 = w - x1;
        y2 = h - x1;
        var d = PPTXShapeUtils.shapeArc(x1, x1, x1, x1, 270, 180, false) +
            PPTXShapeUtils.shapeArc(x1, y2, x1, x1, 180, 90, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(x2, x1, x1, x1, 270, (270 + 90), false) +
            PPTXShapeUtils.shapeArc(x2, y2, x1, x1, 0, 90, false).replace("M", "L");
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成左方括号
     */
    function generateLeftBracket(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 8333 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, y1, y2;
        if (adj < 0) a = 0
        else if (adj > maxAdj) a = maxAdj
        else a = adj
        y1 = Math.min(w, h) * a / cnstVal2;
        if (y1 > w) y1 = w;
        y2 = h - y1;
        var d = "M" + w + "," + h +
            PPTXShapeUtils.shapeArc(y1, y2, y1, y1, 90, 180, false).replace("M", "L") +
            " L" + 0 + "," + y1 +
            PPTXShapeUtils.shapeArc(y1, y1, y1, y1, 180, 270, false).replace("M", "L") +
            " L" + w + "," + 0;
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成右方括号
     */
    function generateRightBracket(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 8333 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var maxAdj = cnstVal1 * h / Math.min(w, h);
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) * slideFactor;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        var a, y1, y2, y3;
        if (adj < 0) a = 0
        else if (adj > maxAdj) a = maxAdj
        else a = adj
        y1 = Math.min(w, h) * a / cnstVal2;
        y2 = h - y1;
        y3 = w - y1;
        var d = "M" + 0 + "," + h +
            PPTXShapeUtils.shapeArc(y3, y2, y1, y1, 90, 0, false).replace("M", "L") +
            " L" + w + "," + h / 2 +
            PPTXShapeUtils.shapeArc(y3, y1, y1, y1, 360, 270, false).replace("M", "L") +
            " L" + 0 + "," + 0;
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成月亮形状
     */
    function generateMoon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var adj = 0.5;
        if (shapAdjst !== undefined) {
            if (typeof shapAdjst === "string") {
                adj = parseInt(shapAdjst.substr(4)) / 100000;
            } else if (shapAdjst["attrs"] !== undefined && shapAdjst["attrs"]["fmla"] !== undefined) {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) / 100000;
            }
        }
        var hd2 = h / 2;
        var adj2 = (1 - adj) * w;
        var d = "M" + w + "," + h +
            PPTXShapeUtils.shapeArc(w, hd2, w, hd2, 90, (90 + 180), false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (90 + 180), 90, false).replace("M", "L") +
            " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成不规则印章形状
     */
    function generateIrregularSeal(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType) {
        border = ensureBorder(border);
        var d;
        if (shapType == "irregularSeal1") {
            d = "M" + w * 10800 / 21600 + "," + h * 5800 / 21600 +
                " L" + w * 14522 / 21600 + "," + 0 +
                " L" + w * 14155 / 21600 + "," + h * 5325 / 21600 +
                " L" + w * 18380 / 21600 + "," + h * 4457 / 21600 +
                " L" + w * 16702 / 21600 + "," + h * 7315 / 21600 +
                " L" + w * 21097 / 21600 + "," + h * 8137 / 21600 +
                " L" + w * 17607 / 21600 + "," + h * 10475 / 21600 +
                " L" + w + "," + h * 13290 / 21600 +
                " L" + w * 16837 / 21600 + "," + h * 12942 / 21600 +
                " L" + w * 18145 / 21600 + "," + h * 18095 / 21600 +
                " L" + w * 14020 / 21600 + "," + h * 14457 / 21600 +
                " L" + w * 13247 / 21600 + "," + h * 19737 / 21600 +
                " L" + w * 10532 / 21600 + "," + h * 14935 / 21600 +
                " L" + w * 8485 / 21600 + "," + h +
                " L" + w * 7715 / 21600 + "," + h * 15627 / 21600 +
                " L" + w * 4762 / 21600 + "," + h * 17617 / 21600 +
                " L" + w * 5667 / 21600 + "," + h * 13937 / 21600 +
                " L" + w * 135 / 21600 + "," + h * 14587 / 21600 +
                " L" + w * 3722 / 21600 + "," + h * 11775 / 21600 +
                " L" + 0 + "," + h * 8615 / 21600 +
                " L" + w * 4627 / 21600 + "," + h * 7617 / 21600 +
                " L" + w * 370 / 21600 + "," + h * 2295 / 21600 +
                " L" + w * 7312 / 21600 + "," + h * 6320 / 21600 +
                " L" + w * 8352 / 21600 + "," + h * 2295 / 21600 +
                " z";
        } else if (shapType == "irregularSeal2") {
            d = "M" + w * 11462 / 21600 + "," + h * 4342 / 21600 +
                " L" + w * 14790 / 21600 + "," + 0 +
                " L" + w * 14525 / 21600 + "," + h * 5777 / 21600 +
                " L" + w * 18007 / 21600 + "," + h * 3172 / 21600 +
                " L" + w * 16380 / 21600 + "," + h * 6532 / 21600 +
                " L" + w + "," + h * 6645 / 21600 +
                " L" + w * 16985 / 21600 + "," + h * 9402 / 21600 +
                " L" + w * 18270 / 21600 + "," + h * 11290 / 21600 +
                " L" + w * 16380 / 21600 + "," + h * 12310 / 21600 +
                " L" + w * 18877 / 21600 + "," + h * 15632 / 21600 +
                " L" + w * 14640 / 21600 + "," + h * 14350 / 21600 +
                " L" + w * 14942 / 21600 + "," + h * 17370 / 21600 +
                " L" + w * 12180 / 21600 + "," + h * 15935 / 21600 +
                " L" + w * 11612 / 21600 + "," + h * 18842 / 21600 +
                " L" + w * 9872 / 21600 + "," + h * 17370 / 21600 +
                " L" + w * 8700 / 21600 + "," + h * 19712 / 21600 +
                " L" + w * 7527 / 21600 + "," + h * 18125 / 21600 +
                " L" + w * 4917 / 21600 + "," + h +
                " L" + w * 4805 / 21600 + "," + h * 18240 / 21600 +
                " L" + w * 1285 / 21600 + "," + h * 17825 / 21600 +
                " L" + w * 2190 / 21600 + "," + h * 15712 / 21600 +
                " L" + 0 + "," + h * 15825 / 21600 +
                " L" + w * 3177 / 21600 + "," + h * 12562 / 21600 +
                " L" + w * 2340 / 21600 + "," + h * 10952 / 21600 +
                " L" + w * 1727 / 21600 + "," + h * 9737 / 21600 +
                " L" + 0 + "," + h * 9850 / 21600 +
                " L" + w * 3177 / 21600 + "," + h * 6587 / 21600 +
                " L" + w * 2190 / 21600 + "," + h * 4475 / 21600 +
                " L" + w * 5667 / 21600 + "," + h * 7237 / 21600 +
                " L" + w * 7692 / 21600 + "," + h * 3720 / 21600 +
                " L" + w * 9872 / 21600 + "," + h * 7155 / 21600 +
                " z";
        }
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成角落形状
     */
    function generateCorner(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var sAdj1_val = 50000 * slideFactor;
        var sAdj2_val = 50000 * slideFactor;
        var cnsVal = 100000 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                if (sAdj_name == "adj1") {
                    var sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                    sAdj1_val = parseInt(sAdj1.substr(4)) * slideFactor;
                } else if (sAdj_name == "adj2") {
                    var sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                    sAdj2_val = parseInt(sAdj2.substr(4)) * slideFactor;
                }
            }
        }
        var minWH = Math.min(w, h);
        var maxAdj1 = cnsVal * h / minWH;
        var maxAdj2 = cnsVal * w / minWH;
        var a1, a2, x1, dy1, y1;
        if (sAdj1_val < 0) a1 = 0
        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
        else a1 = sAdj1_val

        if (sAdj2_val < 0) a2 = 0
        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
        else a2 = sAdj2_val
        x1 = minWH * a2 / cnsVal;
        dy1 = minWH * a1 / cnsVal;
        y1 = h - dy1;

        var d = "M0,0" +
            " L" + x1 + "," + 0 +
            " L" + x1 + "," + y1 +
            " L" + w + "," + y1 +
            " L" + w + "," + h +
            " L0," + h + " z";

        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成对角条纹形状
     */
    function generateDiagStripe(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var sAdj1_val = 50000 * slideFactor;
        var cnsVal = 100000 * slideFactor;
        if (shapAdjst !== undefined) {
            sAdj1_val = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        var a1, x2, y2;
        if (sAdj1_val < 0) a1 = 0
        else if (sAdj1_val > cnsVal) a1 = cnsVal
        else a1 = sAdj1_val
        x2 = w * a1 / cnsVal;
        y2 = h * a1 / cnsVal;
        var d = "M" + 0 + "," + y2 +
            " L" + x2 + "," + 0 +
            " L" + w + "," + 0 +
            " L" + 0 + "," + h + " z";
        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成齿轮形状
     */
    function generateGear(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType) {
        border = ensureBorder(border);
        var gearNum = shapType.substr(4);
        var d = shapeGear(w, h / 3.5, parseInt(gearNum));
        return "<path   d='" + d + "' transform='rotate(20 " + (3 / 7) * h + "," + (3 / 7) * h + ")' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成齿轮形状的辅助函数
     */
    function shapeGear(w, h, teeth) {
        var r1 = w / 2;
        var r2 = r1 * 0.7;
        var r3 = r1 * 0.6;
        var r4 = r1 * 0.3;
        var angleStep = Math.PI * 2 / teeth;
        var points = [];
        
        for (var i = 0; i < teeth; i++) {
            var angle1 = i * angleStep;
            var angle2 = (i + 0.3) * angleStep;
            var angle3 = (i + 0.5) * angleStep;
            var angle4 = (i + 0.7) * angleStep;
            var angle5 = (i + 1) * angleStep;
            
            points.push({
                x: r1 * Math.cos(angle1) + r1,
                y: r1 * Math.sin(angle1) + r1
            });
            points.push({
                x: r2 * Math.cos(angle2) + r1,
                y: r2 * Math.sin(angle2) + r1
            });
            points.push({
                x: r3 * Math.cos(angle3) + r1,
                y: r3 * Math.sin(angle3) + r1
            });
            points.push({
                x: r2 * Math.cos(angle4) + r1,
                y: r2 * Math.sin(angle4) + r1
            });
            points.push({
                x: r1 * Math.cos(angle5) + r1,
                y: r1 * Math.sin(angle5) + r1
            });
        }
        
        var d = "M" + points[0].x + "," + points[0].y;
        for (var j = 1; j < points.length; j++) {
            d += " L" + points[j].x + "," + points[j].y;
        }
        d += " z";
        
        // 添加中心孔
        d += " M" + (r1 - r4) + "," + r1;
        d += PPTXShapeUtils.shapeArc(r1, r1, r4, r4, 0, 360, false).replace("M", "L");
        
        return d;
    }

    /**
     * 生成罐形状和流程图磁盘/磁鼓形状
     */
    function generateCan(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var adj = 25000 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 200000 * slideFactor;
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        var ss = Math.min(w, h);
        var maxAdj, a, y1, y2, y3, dVal;
        if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
            adj = 50000 * slideFactor;
        }
        maxAdj = cnstVal1 * h / ss;
        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        y1 = ss * a / cnstVal2;
        y2 = y1 + y1;
        y3 = h - y1;
        var cd2 = 180, wd2 = w / 2;

        var tranglRott = "";
        if (shapType == "flowChartMagneticDrum") {
            tranglRott = "transform='rotate(90 " + w / 2 + "," + h / 2 + ")'";
        }
        dVal = PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, 0, cd2, false) +
            PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace("M", "L") +
            " L" + w + "," + y3 +
            PPTXShapeUtils.shapeArc(wd2, y3, wd2, y1, 0, cd2, false).replace("M", "L") +
            " L" + 0 + "," + y1;

        return "<path " + tranglRott + " d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成 swoosh 箭头形状
     */
    function generateSwooshArrow(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var refr = slideFactor;
        var sAdj1, adj1 = 25000 * refr;
        var sAdj2, adj2 = 16667 * refr;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                if (sAdj_name == "adj1") {
                    sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                }
            }
        }
        var cnstVal1 = 50000 * refr;
        var cnstVal2 = 100000 * refr;
        var a1, a2, q1, q2, q3, y1, x1, x2, y2, y3, y4;
        if (adj1 < 0) a1 = 0
        else if (adj1 > cnstVal1) a1 = cnstVal1
        else a1 = adj1
        if (adj2 < 0) a2 = 0
        else if (adj2 > cnstVal1) a2 = cnstVal1
        else a2 = adj2
        var minWH = Math.min(w, h);
        q1 = cnstVal1 * h / minWH;
        var maxAdj1 = (q1 < cnstVal1) ? q1 : cnstVal1;
        var a11 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
        y1 = minWH * a11 / cnstVal2;
        q2 = (w - x1) / 2;
        q3 = (h - y1) / 2;
        x1 = minWH * a2 / cnstVal2;
        x2 = w - x1;
        y2 = h - y1;
        y3 = y1 + q3;
        y4 = y2 - q3;
        var dVal = "M" + 0 + "," + y3 +
            " L" + x1 + "," + y3 +
            PPTXShapeUtils.shapeArc(x1, y3, q2, q3, 180, 90, false).replace("M", "L") +
            " L" + x2 + "," + y4 +
            PPTXShapeUtils.shapeArc(x2, y4, q2, q3, 270, 360, false).replace("M", "L") +
            " L" + w + "," + y4 +
            " L" + w + "," + y3 +
            " L" + x2 + "," + y3 +
            PPTXShapeUtils.shapeArc(x2, y3, q2, q3, 0, 90, false).replace("M", "L") +
            " L" + x1 + "," + y4 +
            PPTXShapeUtils.shapeArc(x1, y4, q2, q3, 90, 180, false).replace("M", "L") +
            " L" + 0 + "," + y4 + " z";
        return "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成圆形箭头形状
     */
    function generateCircularArrow(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst_ary) {
        border = ensureBorder(border);
        var slideFactor = PPTXBaseShapes.slideFactor;
        var sAdj1, adj1 = 12500 * slideFactor;
        var sAdj2, adj2 = (1142319 / 60000) * Math.PI / 180;
        var sAdj3, adj3 = (20457681 / 60000) * Math.PI / 180;
        var sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
        var sAdj5, adj5 = 12500 * slideFactor;
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                if (shapAdjst_ary[i]["attrs"] !== undefined) {
                    var sAdj_name = shapAdjst_ary[i]["attrs"]["name"];
                    if (sAdj_name == "adj1") {
                        sAdj1 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj1 !== undefined) {
                            adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                        }
                    } else if (sAdj_name == "adj2") {
                        sAdj2 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj2 !== undefined) {
                            adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                        }
                    } else if (sAdj_name == "adj3") {
                        sAdj3 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj3 !== undefined) {
                            adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                        }
                    } else if (sAdj_name == "adj4") {
                        sAdj4 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj4 !== undefined) {
                            adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                        }
                    } else if (sAdj_name == "adj5") {
                        sAdj5 = shapAdjst_ary[i]["attrs"]["fmla"];
                        if (sAdj5 !== undefined) {
                            adj5 = parseInt(sAdj5.substr(4)) * slideFactor;
                        }
                    }
                }
            }
        }
        var vc = h / 2, hc = w / 2, r = w, b = h, l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
        var ss = Math.min(w, h);
        var cnstVal1 = 25000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var rdAngVal1 = (1 / 60000) * Math.PI / 180;
        var rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
        var rdAngVal3 = 2 * Math.PI;
        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, maxAng, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO, q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng, isw1, isw2, iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng, stAng0;

        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
        maxAdj1 = a5 * 2;
        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
        th = ss * a1 / cnstVal2;
        thh = ss * a5 / cnstVal2;
        th2 = th / 2;
        rw1 = wd2 + th2 - thh;
        rh1 = hd2 + th2 - thh;
        rw2 = rw1 - th;
        rh2 = rh1 - th;
        rw3 = rw2 + th2;
        rh3 = rh2 + th2;
        wtH = rw3 * Math.sin(enAng);
        htH = rh3 * Math.cos(enAng);
        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
        xH = hc + dxH;
        yH = vc + dyH;
        rI = (rw2 < rh2) ? rw2 : rh2;
        u1 = dxH * dxH;
        u2 = dyH * dyH;
        u3 = rI * rI;
        u4 = u1 - u3;
        u5 = u2 - u3;
        u6 = u4 * u5 / u1;
        u7 = u6 / u2;
        u8 = 1 - u7;
        u9 = Math.sqrt(u8);
        u10 = u4 / dxH;
        u11 = u10 / dyH;
        u12 = (1 + u9) / u11;
        u13 = Math.atan2(u12, 1);
        u14 = u13 + rdAngVal3;
        u15 = (u13 > 0) ? u13 : u14;
        u16 = (u15 > Math.PI) ? u15 - Math.PI : u15;
        u17 = Math.PI + u16;
        u18 = u17 - enAng;
        u19 = -u18;
        u20 = u19 - rdAngVal3;
        u21 = (u19 > 0) ? u20 : u19;
        maxAng = Math.abs(u21);
        aAng = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
        ptAng = enAng + aAng;
        wtA = rw3 * Math.sin(ptAng);
        htA = rh3 * Math.cos(ptAng);
        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
        xA = hc + dxA;
        yA = vc + dyA;
        wtE = rw1 * Math.sin(stAng);
        htE = rh1 * Math.cos(stAng);
        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
        xE = hc + dxE;
        yE = vc + dyE;
        dxG = thh * Math.cos(ptAng);
        dyG = thh * Math.sin(ptAng);
        xG = xH + dxG;
        yG = yH + dyG;
        dxB = thh * Math.cos(ptAng);
        dyB = thh * Math.sin(ptAng);
        xB = xH - dxB;
        yB = yH - dyB;
        sx1 = xB - hc;
        sy1 = yB - vc;
        sx2 = xG - hc;
        sy2 = yG - vc;
        rO = (rw1 < rh1) ? rw1 : rh1;
        x1O = sx1 * rO / rw1;
        y1O = sy1 * rO / rh1;
        x2O = sx2 * rO / rw1;
        y2O = sy2 * rO / rh1;
        dxO = x2O - x1O;
        dyO = y2O - y1O;
        dO = Math.sqrt(dxO * dxO + dyO * dyO);
        q1 = x1O * y2O;
        q2 = x2O * y1O;
        DO = q1 - q2;
        q3 = rO * rO;
        q4 = dO * dO;
        q5 = q3 * q4;
        q6 = DO * DO;
        q7 = q5 - q6;
        q8 = (q7 > 0) ? q7 : 0;
        sdelO = Math.sqrt(q8);
        ndyO = dyO * -1;
        sdyO = (ndyO > 0) ? -1 : 1;
        q9 = sdyO * dxO;
        q10 = q9 * sdelO;
        q11 = DO * dyO;
        dxF1 = (q11 + q10) / q4;
        q12 = q11 - q10;
        dxF2 = q12 / q4;
        adyO = Math.abs(dyO);
        q13 = adyO * sdelO;
        q14 = DO * dxO / -1;
        dyF1 = (q14 + q13) / q4;
        q15 = q14 - q13;
        dyF2 = q15 / q4;
        q16 = x2O - dxF1;
        q17 = x2O - dxF2;
        q18 = y2O - dyF1;
        q19 = y2O - dyF2;
        q20 = Math.sqrt(q16 * q16 + q18 * q18);
        q21 = Math.sqrt(q17 * q17 + q19 * q19);
        q22 = q21 - q20;
        dxF = (q22 > 0) ? dxF1 : dxF2;
        dyF = (q22 > 0) ? dyF1 : dyF2;
        sdxF = dxF * rw1 / rO;
        sdyF = dyF * rh1 / rO;
        xF = hc + sdxF;
        yF = vc + sdyF;
        x1I = sx1 * rI / rw2;
        y1I = sy1 * rI / rh2;
        x2I = sx2 * rI / rw2;
        y2I = sy2 * rI / rh2;
        dxI = x2I - x1I;
        dyI = y2I - y1I;
        dI = Math.sqrt(dxI * dxI + dyI * dyI);
        v1 = x1I * y2I;
        v2 = x2I * y1I;
        DI = v1 - v2;
        v3 = rI * rI;
        v4 = dI * dI;
        v5 = v3 * v4;
        v6 = DI * DI;
        v7 = v5 - v6;
        v8 = (v7 > 0) ? v7 : 0;
        sdelI = Math.sqrt(v8);
        var ndyI = dyI * -1;
        var sdyI = (ndyI > 0) ? -1 : 1;
        var v9 = sdyI * dxI;
        var v10 = v9 * sdelI;
        var v11 = DI * dyI;
        var dxC1 = (v11 + v10) / v4;
        var v12 = v11 - v10;
        var dxC2 = v12 / v4;
        var adyI = Math.abs(dyI);
        var v13 = adyI * sdelI;
        var v14 = DI * dxI / -1;
        var dyC1 = (v14 + v13) / v4;
        var v15 = v14 - v13;
        var dyC2 = v15 / v4;
        var v16 = x1I - dxC1;
        var v17 = x1I - dxC2;
        var v18 = y1I - dyC1;
        var v19 = y1I - dyC2;
        var v20 = Math.sqrt(v16 * v16 + v18 * v18);
        var v21 = Math.sqrt(v17 * v17 + v19 * v19);
        var v22 = v21 - v20;
        dxC = (v22 > 0) ? dxC1 : dxC2;
        dyC = (v22 > 0) ? dyC1 : dyC2;
        sdxC = dxC * rw2 / rI;
        sdyC = dyC * rh2 / rI;
        xC = hc + sdxC;
        yC = vc + sdyC;
        ist0 = Math.atan2(sdyC, sdxC);
        ist1 = ist0 + rdAngVal3;
        istAng = (ist0 > 0) ? ist0 : ist1;
        isw1 = stAng - istAng;
        isw2 = isw1 - rdAngVal3;
        iswAng = (isw1 > 0) ? isw2 : isw1;
        p1 = xF - xC;
        p2 = yF - yC;
        p3 = Math.sqrt(p1 * p1 + p2 * p2);
        p4 = p3 / 2;
        p5 = p4 - thh;
        xGp = (p5 > 0) ? xF : xG;
        yGp = (p5 > 0) ? yF : yG;
        xBp = (p5 > 0) ? xC : xB;
        yBp = (p5 > 0) ? yC : yB;
        en0 = Math.atan2(sdyF, sdxF);
        en1 = en0 + rdAngVal3;
        en2 = (en0 > 0) ? en0 : en1;
        sw0 = en2 - stAng;
        sw1 = sw0 + rdAngVal3;
        swAng = (sw0 > 0) ? sw1 : sw0;
        stAng0 = stAng + swAng;

        var strtAng = stAng0 * 180 / Math.PI;
        var endAng = stAng * 180 / Math.PI;
        var stiAng = istAng * 180 / Math.PI;
        var swiAng = iswAng * 180 / Math.PI;
        var ediAng = stiAng + swiAng;

        var d_val = "M" + xE + "," + yE +
            " L" + xE + "," + yE +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
            " L" + xBp + "," + yBp +
            " L" + xA + "," + yA +
            " L" + xGp + "," + yGp +
            " L" + xF + "," + yF +
            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L") +
            " z";
        return "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generatePie: generatePie,
        generateChord: generateChord,
        generateFrame: generateFrame,
        generateDonut: generateDonut,
        generateNoSmoking: generateNoSmoking,
        generateHalfFrame: generateHalfFrame,
        generateBlockArc: generateBlockArc,
        generateBracePair: generateBracePair,
        generateLeftBrace: generateLeftBrace,
        generateRightBrace: generateRightBrace,
        generateBracketPair: generateBracketPair,
        generateLeftBracket: generateLeftBracket,
        generateRightBracket: generateRightBracket,
        generateMoon: generateMoon,
        generateIrregularSeal: generateIrregularSeal,
        generateCorner: generateCorner,
        generateDiagStripe: generateDiagStripe,
        generateGear: generateGear,
        generateCan: generateCan,
        generateSwooshArrow: generateSwooshArrow,
        generateCircularArrow: generateCircularArrow
    };
})();

window.PPTXSpecialShapes = PPTXSpecialShapes;