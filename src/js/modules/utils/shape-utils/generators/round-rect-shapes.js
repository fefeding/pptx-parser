/**
 * 圆角矩形和裁剪矩形形状生成器
 */

var PPTXRoundRectShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成圆角矩形
     */
    function generateRoundRect(w, h, shapType, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, sAdj2_val, "round", "cornrAll");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成单角圆角矩形
     */
    function generateRound1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "round", "cornr1");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成对角圆角矩形
     */
    function generateRound2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "round", "diag");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成双角圆角矩形
     */
    function generateRound2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "round", "cornr2");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成单角裁剪矩形
     */
    function generateSnip1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "snip", "cornr1");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成对角裁剪矩形
     */
    function generateSnip2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "snip", "diag");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成双角裁剪矩形
     */
    function generateSnip2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, 0, "snip", "cornr2");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成裁剪圆角矩形
     */
    function generateSnipRoundRect(w, h, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val, sAdj2_val, "round", "cornrAll");
        return "<path d='" + d_val + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generateRoundRect: generateRoundRect,
        generateRound1Rect: generateRound1Rect,
        generateRound2DiagRect: generateRound2DiagRect,
        generateRound2SameRect: generateRound2SameRect,
        generateSnip1Rect: generateSnip1Rect,
        generateSnip2DiagRect: generateSnip2DiagRect,
        generateSnip2SameRect: generateSnip2SameRect,
        generateSnipRoundRect: generateSnipRoundRect
    };
})();

window.PPTXRoundRectShapes = PPTXRoundRectShapes;