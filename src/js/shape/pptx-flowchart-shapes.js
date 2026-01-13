/**
 * PPTX Flowchart Shapes Utils
 * 生成流程图形状的 SVG 元素
 */

(function() {
    "use strict";

    // 流程图：收集
    function genFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var d = "M 0,0" +
            " L" + w + "," + 0 +
            " L" + 0 + "," + h +
            " L" + w + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：文档
    function genFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var y1 = h * 17322 / 21600;
        var y2 = h * 20172 / 21600;
        var y3 = h * 23922 / 21600;
        var x1 = w * 10800 / 21600;
        var d = "M" + 0 + "," + 0 +
            " L" + w + "," + 0 +
            " L" + w + "," + y1 +
            " C" + x1 + "," + y1 + " " + x1 + "," + y3 + " " + 0 + "," + y2 +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：多文档
    function genFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var y1 = h * 18022 / 21600;
        var y2 = h * 3675 / 21600;
        var y3 = h * 23542 / 21600;
        var y4 = h * 1815 / 21600;
        var y5 = h * 16252 / 21600;
        var y6 = h * 16352 / 21600;
        var y7 = h * 14392 / 21600;
        var y8 = h * 20782 / 21600;
        var y9 = h * 14467 / 21600;
        var x1 = w * 1532 / 21600;
        var x2 = w * 20000 / 21600;
        var x3 = w * 9298 / 21600;
        var x4 = w * 19298 / 21600;
        var x5 = w * 18595 / 21600;
        var x6 = w * 2972 / 21600;
        var x7 = w * 20800 / 21600;
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

        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：终止符
    function genFlowChartTerminator(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var x1 = w * 3475 / 21600;
        var x2 = w * 18125 / 21600;
        var y1 = h * 10800 / 21600;
        var cd2 = 180, cd4 = 90, c3d4 = 270;
        var d = "M" + x1 + "," + 0 +
            " L" + x2 + "," + 0 +
            window.PPTXShapeUtils.shapeArc(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
            " L" + x1 + "," + h +
            window.PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：穿孔纸带
    function genFlowChartPunchedTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var x1 = w * 5 / 20;
        var y1 = h * 2 / 20;
        var y2 = h * 18 / 20;
        var cd2 = 180;
        var d = "M" + 0 + "," + y1 +
            window.PPTXShapeUtils.shapeArc(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
            window.PPTXShapeUtils.shapeArc(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
            " L" + w + "," + y2 +
            window.PPTXShapeUtils.shapeArc(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
            window.PPTXShapeUtils.shapeArc(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：在线存储
    function genFlowChartOnlineStorage(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var x1 = w * 1 / 6;
        var y1 = h * 3 / 6;
        var c3d4 = 270, cd4 = 90;
        var d = "M" + x1 + "," + 0 +
            " L" + w + "," + 0 +
            window.PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
            " L" + x1 + "," + h +
            window.PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：显示
    function genFlowChartDisplay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var x1 = w * 1 / 6;
        var x2 = w * 5 / 6;
        var y1 = h * 3 / 6;
        var c3d4 = 270, cd2 = 180;
        var d = "M" + x1 + "," + 0 +
            " L" + x2 + "," + 0 +
            " L" + x2 + "," + h +
            " L" + x1 + "," + h +
            " z" +
            "M" + 0 + "," + (h / 2) +
            " L" + x1 + "," + 0 +
            " L" + x1 + "," + h +
            " z" +
            "M" + w + "," + (h / 2) +
            " L" + x2 + "," + 0 +
            " L" + x2 + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：延迟
    function genFlowChartDelay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        var wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
        var d = "M" + 0 + "," + 0 +
            " L" + w + "," + 0 +
            window.PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 0, cd2, false).replace("M", "L") +
            " L" + 0 + "," + h +
            " z";
        return "<path d='" + d + "'  fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 流程图：决策
    function genFlowChartDecision(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
        return "<polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + fillAttr +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    // 导出函数到全局对象
    if (!window.PPTXFlowchartShapes) {
        window.PPTXFlowchartShapes = {};
    }

    window.PPTXFlowchartShapes.genFlowChartCollate = genFlowChartCollate;
    window.PPTXFlowchartShapes.genFlowChartDocument = genFlowChartDocument;
    window.PPTXFlowchartShapes.genFlowChartMultidocument = genFlowChartMultidocument;
    window.PPTXFlowchartShapes.genFlowChartTerminator = genFlowChartTerminator;
    window.PPTXFlowchartShapes.genFlowChartPunchedTape = genFlowChartPunchedTape;
    window.PPTXFlowchartShapes.genFlowChartOnlineStorage = genFlowChartOnlineStorage;
    window.PPTXFlowchartShapes.genFlowChartDisplay = genFlowChartDisplay;
    window.PPTXFlowchartShapes.genFlowChartDelay = genFlowChartDelay;
    window.PPTXFlowchartShapes.genFlowChartDecision = genFlowChartDecision;

})();
