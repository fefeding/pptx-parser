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
    function generateTriangle(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = 96 / 914400;
        var shapAdjst_val = 0.5;
        var tranglRott = "";
        
        if (shapAdjst !== undefined) {
            shapAdjst_val = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        if (shapType == "flowChartMerge") {
            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
        }
        
        return " <polygon " + tranglRott + " points='" + (w * shapAdjst_val) + " 0,0 " + h + "," + w + " " + h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成菱形
     */
    function generateDiamond(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType) {
        border = ensureBorder(border);
        var result = " <polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        
        if (shapType == "flowChartSort") {
            result += " <polyline points='0 " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        }
        
        return result;
    }

    /**
     * 生成梯形
     */
    function generateTrapezoid(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = 96 / 914400;
        var adjst_val = 0.2;
        var max_adj_const = 0.7407;
        var cnstVal = 0;
        var tranglRott = "";
        
        if (shapAdjst !== undefined) {
            var adjst = parseInt(shapAdjst.substr(4)) * slideFactor;
            adjst_val = (adjst * 0.5) / max_adj_const;
        }
        
        if (shapType == "flowChartManualOperation") {
            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
        }
        
        if (shapType == "flowChartManualInput") {
            adjst_val = 0;
            cnstVal = h / 5;
        }
        
        var d = "M" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0";
        return " <polygon " + tranglRott + " points='" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成平行四边形
     */
    function generateParallelogram(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapAdjst) {
        border = ensureBorder(border);
        var adjst_val = 0.25;
        var max_adj_const;
        
        if (w > h) {
            max_adj_const = w / h;
        } else {
            max_adj_const = h / w;
        }
        
        if (shapAdjst !== undefined) {
            var adjst = parseInt(shapAdjst.substr(4)) / 100000;
            adjst_val = adjst / max_adj_const;
        }
        
        var d = "M" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0";
        return " <polygon points='" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成五边形
     */
    function generatePentagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
        border = ensureBorder(border);
        return " <polygon points='" + (0.5 * w) + " 0,0 " + (0.375 * h) + "," + (0.15 * w) + " " + h + "," + 0.85 * w + " " + h + "," + w + " " + 0.375 * h + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    /**
     * 生成六边形
     */
    function generateHexagon(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, shapType, shapAdjst) {
        border = ensureBorder(border);
        var slideFactor = 96 / 914400;
        var adj = 25000 * slideFactor;
        var vf = 115470 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var cnstVal2 = 100000 * slideFactor;
        var angVal1 = 60 * Math.PI / 180;
        
        if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        var maxAdj, a, shd2, x1, x2, dy1, y1, y2, vc = h / 2, hd2 = h / 2;
        var ss = Math.min(w, h);
        maxAdj = cnstVal1 * w / ss;
        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        shd2 = hd2 * vf / cnstVal2;
        x1 = ss * a / cnstVal2;
        x2 = w - x1;
        dy1 = shd2 * Math.sin(angVal1);
        y1 = vc - dy1;
        y2 = vc + dy1;

        var d = "M" + 0 + "," + vc +
            " L" + x1 + "," + y1 +
            " L" + x2 + "," + y1 +
            " L" + w + "," + vc +
            " L" + x2 + "," + y2 +
            " L" + x1 + "," + y2 +
            " z";

        return "<path   d='" + d + "'  fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
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