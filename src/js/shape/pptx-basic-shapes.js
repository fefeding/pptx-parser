/**
 * PPTX Basic Shapes Utils
 * 生成基本形状（矩形、椭圆、圆形、三角形等）的 SVG 元素
 */

(function() {
    "use strict";

    // 生成矩形
    function genRect(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr) {
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        return "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + " />";
    }

    // 生成椭圆
    function genEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        return "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成菱形
    function genDiamond(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        return "<polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成三角形
    function genTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        return "<polygon points='" + (w / 2) + " 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成圆形
    function genCircle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var r = Math.min(w, h) / 2;
        var cx = w / 2;
        var cy = h / 2;
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        return "<circle cx='" + cx + "' cy='" + cy + "' r='" + r + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成五边形
    function genPentagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var hc = w / 2, vc = h / 2, r = Math.min(w, h) / 2;
        var dx2 = r * Math.sin(Math.PI * 2 / 5);
        var dy2 = r * Math.cos(Math.PI * 2 / 5);
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        var d = "M" + hc + "," + (vc - r) +
            " L" + (hc + dx2) + "," + (vc - dy2) +
            " L" + (hc + dx2) + "," + (vc + dy2) +
            " L" + (hc - dx2) + "," + (vc + dy2) +
            " L" + (hc - dx2) + "," + (vc - dy2) +
            " z";
        return "<path d='" + d + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成六边形
    function genHexagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var hc = w / 2, vc = h / 2, r = Math.min(w, h) / 2;
        var dx2 = r * Math.sin(Math.PI / 3);
        var dy2 = r * Math.cos(Math.PI / 3);
        var fillAttr;
        if (imgFillFlg) {
            fillAttr = "url(#imgPtrn_" + shpId + ")";
        } else if (grndFillFlg) {
            fillAttr = "url(#linGrd_" + shpId + ")";
        } else {
            fillAttr = fillColor;
        }
        var d = "M" + hc + "," + (vc - r) +
            " L" + (hc + dx2) + "," + (vc - dy2) +
            " L" + (hc + dx2) + "," + (vc + dy2) +
            " L" + hc + "," + (vc + r) +
            " L" + (hc - dx2) + "," + (vc + dy2) +
            " L" + (hc - dx2) + "," + (vc - dy2) +
            " z";
        return "<path d='" + d + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 生成带特殊标记的矩形（用于流程图）
    function genRectWithDecoration(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, shapType) {
        var result = genRect(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr);

        if (shapType == "flowChartPredefinedProcess") {
            result += "<rect x='" + w * (1 / 8) + "' y='0' width='" + w * (6 / 8) + "' height='" + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        } else if (shapType == "flowChartInternalStorage") {
            result += " <polyline points='" + w * (1 / 8) + " 0," + w * (1 / 8) + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            result += " <polyline points='0 " + h * (1 / 8) + "," + w + " " + h * (1 / 8) + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        }

        return result;
    }

    // 导出函数到全局对象
    if (!window.PPTXBasicShapes) {
        window.PPTXBasicShapes = {};
    }

    window.PPTXBasicShapes.genRect = genRect;
    window.PPTXBasicShapes.genEllipse = genEllipse;
    window.PPTXBasicShapes.genDiamond = genDiamond;
    window.PPTXBasicShapes.genTriangle = genTriangle;
    window.PPTXBasicShapes.genCircle = genCircle;
    window.PPTXBasicShapes.genPentagon = genPentagon;
    window.PPTXBasicShapes.genHexagon = genHexagon;
    window.PPTXBasicShapes.genRectWithDecoration = genRectWithDecoration;

})();
