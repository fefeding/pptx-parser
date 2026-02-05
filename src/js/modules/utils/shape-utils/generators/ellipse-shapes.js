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
    function generateEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var rx = w / 2;
        var ry = h / 2;
        return "<ellipse cx='" + cx + "' cy='" + cy + "' rx='" + rx + "' ry='" + ry + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图终结符形状
     */
    function generateFlowChartTerminator(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图穿孔纸带形状
     */
    function generateFlowChartPunchedTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图在线存储形状
     */
    function generateFlowChartOnlineStorage(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图显示形状
     */
    function generateFlowChartDisplay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图延迟形状
     */
    function generateFlowChartDelay(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成流程图磁带形状
     */
    function generateFlowChartMagneticTape(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var rx = w / 2;
        var ry = h / 2;
        return "<path d='M" + 0 + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + w + "," + ry +
            " A" + rx + "," + ry + " 0 0 1 " + 0 + "," + ry +
            " z' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
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