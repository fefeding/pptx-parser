/**
 * 星形和圆形形状生成器
 */

var PPTXStarShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成星形
     * @param {number} w - 宽度
     * @param {number} h - 高度
     * @param {number} points - 星形的点数
     * @param {number} innerRadiusRatio - 内半径比例
     * @param {boolean} imgFillFlg - 图片填充标志
     * @param {boolean} grndFillFlg - 渐变填充标志
     * @param {string} shpId - 形状ID
     * @param {string} fillColor - 填充颜色
     * @param {Object} border - 边框对象
     * @returns {string} SVG字符串
     */
    function generateStar(w, h, points, innerRadiusRatio, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var outerRadius = Math.min(w, h) / 2;
        var innerRadius = outerRadius * innerRadiusRatio;
        var pointsArray = [];
        var angleStep = Math.PI / points;

        for (var i = 0; i < 2 * points; i++) {
            var radius = (i % 2 === 0) ? outerRadius : innerRadius;
            var angle = i * angleStep - Math.PI / 2;
            var x = cx + radius * Math.cos(angle);
            var y = cy + radius * Math.sin(angle);
            pointsArray.push(x + "," + y);
        }

        var d = "M" + pointsArray.join(" L") + " z";
        return "<polygon points='" + pointsArray.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成四角星
     */
    function generateStar4(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 4, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成五角星
     */
    function generateStar5(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 5, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成六角星
     */
    function generateStar6(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 6, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成七角星
     */
    function generateStar7(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 7, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成八角星
     */
    function generateStar8(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 8, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成十角星
     */
    function generateStar10(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 10, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成十二角星
     */
    function generateStar12(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 12, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成十六角星
     */
    function generateStar16(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 16, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成二十四角星
     */
    function generateStar24(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 24, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    /**
     * 生成三十二角星
     */
    function generateStar32(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        return generateStar(w, h, 32, 0.5, imgFillFlg, grndFillFlg, shpId, fillColor, border);
    }

    return {
        generateStar: generateStar,
        generateStar4: generateStar4,
        generateStar5: generateStar5,
        generateStar6: generateStar6,
        generateStar7: generateStar7,
        generateStar8: generateStar8,
        generateStar10: generateStar10,
        generateStar12: generateStar12,
        generateStar16: generateStar16,
        generateStar24: generateStar24,
        generateStar32: generateStar32
    };
})();

window.PPTXStarShapes = PPTXStarShapes;