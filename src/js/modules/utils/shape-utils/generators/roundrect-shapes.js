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
     * @param {number} w - 宽度
     * @param {number} h - 高度
     * @param {string} shapType - 形状类型
     * @param {number} sAdj1_val - 调整值1
     * @param {number} sAdj2_val - 调整值2
     * @param {boolean} imgFillFlg - 图片填充标志
     * @param {boolean} grndFillFlg - 渐变填充标志
     * @param {string} shpId - 形状ID
     * @param {string} fillColor - 填充颜色
     * @param {Object} border - 边框对象
     * @returns {string} SVG字符串
     */
    function generateRoundRect(w, h, shapType, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var result = "";
        var adjTyp = shapType;
        var sAdj1 = sAdj1_val || 0.25;
        var sAdj2 = sAdj2_val || 0.25;
        
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj2, shapType, adjTyp);
        
        result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        
        return result;
    }

    /**
     * 生成裁剪圆角矩形
     */
    function generateSnipRoundRect(w, h, sAdj1_val, sAdj2_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var sAdj2 = sAdj2_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj2, "snipRoundRect", "snipRoundRect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成单角圆角矩形
     */
    function generateRound1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, 0, "round1Rect", "round1Rect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成对角双角圆角矩形
     */
    function generateRound2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj1, "round2DiagRect", "round2DiagRect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成同侧双角圆角矩形
     */
    function generateRound2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj1, "round2SameRect", "round2SameRect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成单角裁剪矩形
     */
    function generateSnip1Rect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, 0, "snip1Rect", "snip1Rect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成对角双角裁剪矩形
     */
    function generateSnip2DiagRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj1, "snip2DiagRect", "snip2DiagRect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成同侧双角裁剪矩形
     */
    function generateSnip2SameRect(w, h, sAdj1_val, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var sAdj1 = sAdj1_val || 0.25;
        var d = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1, sAdj1, "snip2SameRect", "snip2SameRect");
        
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generateRoundRect: generateRoundRect,
        generateSnipRoundRect: generateSnipRoundRect,
        generateRound1Rect: generateRound1Rect,
        generateRound2DiagRect: generateRound2DiagRect,
        generateRound2SameRect: generateRound2SameRect,
        generateSnip1Rect: generateSnip1Rect,
        generateSnip2DiagRect: generateSnip2DiagRect,
        generateSnip2SameRect: generateSnip2SameRect
    };
})();

window.PPTXRoundRectShapes = PPTXRoundRectShapes;