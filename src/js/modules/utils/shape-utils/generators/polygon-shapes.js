/**
 * 三角形和多边形形状生成器
 */

var PPTXPolygonShapes = (function() {
    function ensureBorder(border) {
        if (border === undefined) {
            border = { color: "#000000", width: 1, strokeDasharray: "none" };
        }
        return border;
    }

    /**
     * 生成直角三角形
     */
    function generateRtTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M0,0 L" + w + "," + h + " L0," + h + " z";
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成三角形
     */
    function generateTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M" + w / 2 + ",0 L" + w + "," + h + " L0," + h + " z";
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成菱形
     */
    function generateDiamond(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M" + w / 2 + ",0 L" + w + "," + h / 2 + " L" + w / 2 + "," + h + " L0," + h / 2 + " z";
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成梯形
     */
    function generateTrapezoid(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M" + w * 0.25 + ",0 L" + w * 0.75 + ",0 L" + w + "," + h + " L0," + h + " z";
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成平行四边形
     */
    function generateParallelogram(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var d = "M" + w * 0.25 + ",0 L" + w + ",0 L" + w * 0.75 + "," + h + " L0," + h + " z";
        return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成五边形
     */
    function generatePentagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 5; i++) {
            var angle = (i * 72 - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成六边形
     */
    function generateHexagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 6; i++) {
            var angle = (i * 60 - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成七边形
     */
    function generateHeptagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 7; i++) {
            var angle = (i * (360 / 7) - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成八边形
     */
    function generateOctagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 8; i++) {
            var angle = (i * 45 - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成十边形
     */
    function generateDecagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 10; i++) {
            var angle = (i * 36 - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成十二边形
     */
    function generateDodecagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        var cx = w / 2;
        var cy = h / 2;
        var r = Math.min(w, h) / 2;
        var points = [];
        for (var i = 0; i < 12; i++) {
            var angle = (i * 30 - 90) * Math.PI / 180;
            var x = cx + r * Math.cos(angle);
            var y = cy + r * Math.sin(angle);
            points.push(x + "," + y);
        }
        var d = "M" + points.join(" L") + " z";
        return "<polygon points='" + points.join(" ") + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return {
        generateRtTriangle: generateRtTriangle,
        generateTriangle: generateTriangle,
        generateDiamond: generateDiamond,
        generateTrapezoid: generateTrapezoid,
        generateParallelogram: generateParallelogram,
        generatePentagon: generatePentagon,
        generateHexagon: generateHexagon,
        generateHeptagon: generateHeptagon,
        generateOctagon: generateOctagon,
        generateDecagon: generateDecagon,
        generateDodecagon: generateDodecagon
    };
})();

window.PPTXPolygonShapes = PPTXPolygonShapes;