/**
 * 矩形和流程图形状生成器
 */

var PPTXRectShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成矩形形状
     */
    function generateRect(w, h, shapType, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr) {
        border = ensureBorder(border);
        var result = "";
        result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + "  />";

        if (shapType == "flowChartPredefinedProcess") {
            result += "<rect x='" + w * (1 / 8) + "' y='0' width='" + w * (6 / 8) + "' height='" + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        } else if (shapType == "flowChartInternalStorage") {
            result += " <polyline points='" + w * (1 / 8) + " 0," + w * (1 / 8) + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
            result += " <polyline points='0 " + h * (1 / 8) + "," + w + " " + h * (1 / 8) + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        }
        return result;
    }

    /**
     * 生成流程图整理形状
     */
    function generateFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M 0,0" +
            " L" + w + "," + 0 +
            " L" + 0 + "," + h +
            " L" + w + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图文档形状
     */
    function generateFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var y1, y2, y3, x1;
        x1 = w * 10800 / 21600;
        y1 = h * 17322 / 21600;
        y2 = h * 20172 / 21600;
        y3 = h * 23922 / 21600;
        var d = "M" + 0 + "," + 0 +
            " L" + w + "," + 0 +
            " L" + w + "," + y1 +
            " C" + x1 + "," + y1 + " " + x1 + "," + y3 + " " + 0 + "," + y2 +
            " z";
        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图多文档形状
     */
    function generateFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var y1, y2, y3, y4, y5, y6, y7, y8, y9, x1, x2, x3, x4, x5, x6, x7;
        y1 = h * 18022 / 21600;
        y2 = h * 3675 / 21600;
        y3 = h * 23542 / 21600;
        y4 = h * 1815 / 21600;
        y5 = h * 16252 / 21600;
        y6 = h * 16352 / 21600;
        y7 = h * 14392 / 21600;
        y8 = h * 20782 / 21600;
        y9 = h * 14467 / 21600;
        x1 = w * 1532 / 21600;
        x2 = w * 20000 / 21600;
        x3 = w * 9298 / 21600;
        x4 = w * 19298 / 21600;
        x5 = w * 18595 / 21600;
        x6 = w * 2972 / 21600;
        x7 = w * 20800 / 21600;
        var d = "M" + 0 + "," + y2 +
            " L" + x5 + "," + y2 +
            " L" + x5 + "," + y1 +
            " C" + x3 + "," + y1 + " " + x3 + "," + y3 + " " + 0 + "," + y8 +
            " z" +
            "M" + x1 + "," + y2 +
            " L" + x1 + "," + y4 +
            " L" + x2 + "," + y4 +
            " L" + x2 + "," + y5 +
            " C" + x4 + "," + y5 + " " + x5 + "," + y6 + " " + x5 + "," + y6 +
            "M" + x6 + "," + y4 +
            " L" + x6 + "," + 0 +
            " L" + w + "," + 0 +
            " L" + w + "," + y7 +
            " C" + x7 + "," + y7 + " " + x2 + "," + y9 + " " + x2 + "," + y9;

        return "<path d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generateRect: generateRect,
        generateFlowChartCollate: generateFlowChartCollate,
        generateFlowChartDocument: generateFlowChartDocument,
        generateFlowChartMultidocument: generateFlowChartMultidocument
    };
})();

window.PPTXRectShapes = PPTXRectShapes;