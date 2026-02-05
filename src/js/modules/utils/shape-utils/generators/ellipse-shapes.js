/**
 * 椭圆和圆形形状生成器
 */

var PPTXEllipseShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成椭圆形状
     */
    function generateEllipse(w, h, shapType, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var rx = w / 2;
        var ry = h / 2;
        var result = "<ellipse cx='" + cx + "' cy='" + cy + "' rx='" + rx + "' ry='" + ry + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

        if (shapType == "flowChartOr") {
            result += " <polyline points='" + w / 2 + " " + 0 + "," + w / 2 + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            result += " <polyline points='" + 0 + " " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        } else if (shapType == "flowChartSummingJunction") {
            var iDx, idy, il, ir, it, ib, hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
            var angVal = Math.PI / 4;
            iDx = wd2 * Math.cos(angVal);
            idy = hd2 * Math.sin(angVal);
            il = hc - iDx;
            ir = hc + iDx;
            it = vc - idy;
            ib = vc + idy;
            result += " <polyline points='" + il + " " + it + "," + ir + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            result += " <polyline points='" + ir + " " + it + "," + il + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        }
        return result;
    }

    /**
     * 生成流程图终结符形状
     */
    function generateFlowChartTerminator(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var x1, x2, y1, cd2 = 180, cd4 = 90, c3d4 = 270;
        x1 = w * 3475 / 21600;
        x2 = w * 18125 / 21600;
        y1 = h * 10800 / 21600;
        var d = "M" + x1 + "," + 0 +
            " L" + x2 + "," + 0 +
            PPTXShapeUtils.shapeArcAlt(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
            " L" + x1 + "," + h +
            PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图穿孔纸带形状
     */
    function generateFlowChartPunchedTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var x1, y1, y2, cd2 = 180;
        x1 = w * 5 / 20;
        y1 = h * 2 / 20;
        y2 = h * 18 / 20;
        var d = "M" + 0 + "," + y1 +
            PPTXShapeUtils.shapeArcAlt(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
            PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
            " L" + w + "," + y2 +
            PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
            PPTXShapeUtils.shapeArcAlt(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图在线存储形状
     */
    function generateFlowChartOnlineStorage(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var x1, y1, c3d4 = 270, cd4 = 90;
        x1 = w * 1 / 6;
        y1 = h * 3 / 6;
        var d = "M" + x1 + "," + 0 +
            " L" + w + "," + 0 +
            PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
            " L" + x1 + "," + h +
            PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图显示形状
     */
    function generateFlowChartDisplay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var x1, x2, y1, c3d4 = 270, cd2 = 180;
        x1 = w * 1 / 6;
        x2 = w * 5 / 6;
        y1 = h * 3 / 6;
        var d = "M" + 0 + "," + y1 +
            " L" + x1 + "," + 0 +
            " L" + x2 + "," + 0 +
            PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
            " L" + x1 + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图延迟形状
     */
    function generateFlowChartDelay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
        var d = "M" + 0 + "," + 0 +
            " L" + wd2 + "," + 0 +
            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L") +
            " L" + 0 + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图磁带形状
     */
    function generateFlowChartMagneticTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
        var idy, ib, ang1;
        idy = hd2 * Math.sin(Math.PI / 4);
        ib = hd2 + idy;
        ang1 = Math.atan(h / w);
        var ang1Dg = ang1 * 180 / Math.PI;
        var d = "M" + wd2 + "," + h +
            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L") +
            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L") +
            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L") +
            PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L") +
            " L" + w + "," + ib +
            " L" + w + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generateEllipse: generateEllipse,
        generateFlowChartTerminator: generateFlowChartTerminator,
        generateFlowChartPunchedTape: generateFlowChartPunchedTape,
        generateFlowChartOnlineStorage: generateFlowChartOnlineStorage,
        generateFlowChartDisplay: generateFlowChartDisplay,
        generateFlowChartDelay: generateFlowChartDelay,
        generateFlowChartMagneticTape: generateFlowChartMagneticTape
    };
})();

window.PPTXEllipseShapes = PPTXEllipseShapes;