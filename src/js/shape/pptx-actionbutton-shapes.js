    // 辅助函数：生成操作按钮的基础矩形
function getActionButtonRect(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
    return "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

    // 辅助函数：生成路径元素
function genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
    return "<path d='" + d + "'  fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

    // 操作按钮：后退/上一个
function genActionButtonBackPrevious(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g11 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：开始
function genActionButtonBeginning(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 / 8;
    var g15 = g13 / 4;
    var g16 = g11 + g14;
    var g17 = g11 + g15;
    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g17 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z" +
        "M" + g16 + "," + g9 +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " L" + g16 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：文档
function genActionButtonDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var dx1 = ss * 9 / 32;
    var g11 = hc - dx1;
    var g12 = hc + dx1;
    var g13 = ss * 3 / 16;
    var g14 = g12 - g13;
    var g15 = g9 + g13;
    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g11 + "," + g9 +
        " L" + g14 + "," + g9 +
        " L" + g12 + "," + g15 +
        " L" + g12 + "," + g10 +
        " L" + g11 + "," + g10 +
        " z" +
        "M" + g14 + "," + g9 +
        " L" + g14 + "," + g15 +
        " L" + g12 + "," + g15 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：结束
function genActionButtonEnd(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 * 3 / 4;
    var g15 = g13 * 7 / 8;
    var g16 = g11 + g14;
    var g17 = g11 + g15;
    var d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g17 + "," + g9 +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " L" + g17 + "," + g10 +
        " z" +
        " M" + g16 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：前进/下一个
function genActionButtonForwardNext(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：帮助
function genActionButtonHelp(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g11 = hc - dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 / 7;
    var g15 = g13 * 3 / 14;
    var g16 = g13 * 2 / 7;
    var g19 = g13 * 3 / 7;
    var g20 = g13 * 4 / 7;
    var g21 = g13 * 17 / 28;
    var g23 = g13 * 21 / 28;
    var g24 = g13 * 11 / 14;
    var g27 = g9 + g16;
    var g29 = g9 + g21;
    var g30 = g9 + g23;
    var g31 = g9 + g24;
    var g33 = g11 + g15;
    var g36 = g11 + g19;
    var g37 = g11 + g20;
    var g41 = g13 / 14;
    var g42 = g13 * 3 / 28;
    var cX1 = g33 + g16;
    var cX2 = g36 + g14;
    var cY3 = g31 + g42;
    var cX4 = (g37 + g36 + g16) / 2;

    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g33 + "," + g27 +
        window.PPTXShapeUtils.shapeArc(cX1, g27, g16, g16, 180, 360, false).replace("M", "L") +
        window.PPTXShapeUtils.shapeArc(cX4, g27, g14, g15, 0, 90, false).replace("M", "L") +
        window.PPTXShapeUtils.shapeArc(cX4, g29, g41, g42, 270, 180, false).replace("M", "L") +
        " L" + g37 + "," + g30 +
        " L" + g36 + "," + g30 +
        " L" + g36 + "," + g29 +
        window.PPTXShapeUtils.shapeArc(cX2, g29, g14, g15, 180, 270, false).replace("M", "L") +
        window.PPTXShapeUtils.shapeArc(g37, g27, g41, g42, 90, 0, false).replace("M", "L") +
        window.PPTXShapeUtils.shapeArc(cX1, g27, g14, g14, 0, -180, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g31 +
        window.PPTXShapeUtils.shapeArc(hc, cY3, g42, g42, 270, 630, false).replace("M", "L") +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：主页
function genActionButtonHome(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 / 16;
    var g15 = g13 / 8;
    var g16 = g13 * 3 / 16;
    var g17 = g13 * 5 / 16;
    var g18 = g13 * 7 / 16;
    var g19 = g13 * 9 / 16;
    var g20 = g13 * 11 / 16;
    var g21 = g13 * 3 / 4;
    var g22 = g13 * 13 / 16;
    var g23 = g13 * 7 / 8;
    var g24 = g9 + g14;
    var g25 = g9 + g16;
    var g26 = g9 + g17;
    var g27 = g9 + g21;
    var g28 = g11 + g15;
    var g29 = g11 + g18;
    var g30 = g11 + g19;
    var g31 = g11 + g20;
    var g32 = g11 + g22;
    var g33 = g11 + g23;

    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        " M" + hc + "," + g9 +
        " L" + g11 + "," + vc +
        " L" + g28 + "," + vc +
        " L" + g28 + "," + g10 +
        " L" + g33 + "," + g10 +
        " L" + g33 + "," + vc +
        " L" + g12 + "," + vc +
        " L" + g32 + "," + g26 +
        " L" + g32 + "," + g24 +
        " L" + g31 + "," + g24 +
        " L" + g31 + "," + g25 +
        " z" +
        " M" + g29 + "," + g27 +
        " L" + g30 + "," + g27 +
        " L" + g30 + "," + g10 +
        " L" + g29 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：信息
function genActionButtonInformation(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g11 = hc - dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 / 32;
    var g17 = g13 * 5 / 16;
    var g18 = g13 * 3 / 8;
    var g19 = g13 * 13 / 32;
    var g20 = g13 * 19 / 32;
    var g22 = g13 * 11 / 16;
    var g23 = g13 * 13 / 16;
    var g24 = g13 * 7 / 8;
    var g25 = g9 + g14;
    var g28 = g9 + g17;
    var g29 = g9 + g18;
    var g30 = g9 + g23;
    var g31 = g9 + g24;
    var g32 = g11 + g17;
    var g34 = g11 + g19;
    var g35 = g11 + g20;
    var g37 = g11 + g22;
    var g38 = g13 * 3 / 32;
    var cY1 = g9 + dx2;
    var cY2 = g25 + g38;

    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + hc + "," + g9 +
        window.PPTXShapeUtils.shapeArc(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g25 +
        window.PPTXShapeUtils.shapeArc(hc, cY2, g38, g38, 270, 630, false).replace("M", "L") +
        "M" + g32 + "," + g28 +
        " L" + g35 + "," + g28 +
        " L" + g35 + "," + g30 +
        " L" + g37 + "," + g30 +
        " L" + g37 + "," + g31 +
        " L" + g32 + "," + g31 +
        " L" + g32 + "," + g30 +
        " L" + g34 + "," + g30 +
        " L" + g34 + "," + g29 +
        " L" + g32 + "," + g29 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：电影
function genActionButtonMovie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 * 1455 / 21600;
    var g15 = g13 * 1905 / 21600;
    var g16 = g13 * 2325 / 21600;
    var g17 = g13 * 16155 / 21600;
    var g18 = g13 * 17010 / 21600;
    var g19 = g13 * 19335 / 21600;
    var g20 = g13 * 19725 / 21600;
    var g21 = g13 * 20595 / 21600;
    var g22 = g13 * 5280 / 21600;
    var g23 = g13 * 5730 / 21600;
    var g24 = g13 * 6630 / 21600;
    var g25 = g13 * 7492 / 21600;
    var g26 = g13 * 9067 / 21600;
    var g27 = g13 * 9555 / 21600;
    var g28 = g13 * 13342 / 21600;
    var g29 = g13 * 14580 / 21600;
    var g30 = g13 * 15592 / 21600;
    var g31 = g11 + g14;
    var g32 = g11 + g15;
    var g33 = g11 + g16;
    var g34 = g11 + g17;
    var g35 = g11 + g18;
    var g36 = g11 + g19;
    var g37 = g11 + g20;
    var g38 = g11 + g21;
    var g39 = g9 + g22;
    var g40 = g9 + g23;
    var g41 = g9 + g24;
    var g42 = g9 + g25;
    var g43 = g9 + g26;
    var g44 = g9 + g27;
    var g45 = g9 + g28;
    var g46 = g9 + g29;
    var g47 = g9 + g30;
    var g48 = g9 + g31;

    var d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        "M" + g11 + "," + g39 +
        " L" + g11 + "," + g44 +
        " L" + g31 + "," + g44 +
        " L" + g32 + "," + g43 +
        " L" + g33 + "," + g43 +
        " L" + g33 + "," + g47 +
        " L" + g35 + "," + g47 +
        " L" + g35 + "," + g45 +
        " L" + g36 + "," + g45 +
        " L" + g38 + "," + g46 +
        " L" + g12 + "," + g46 +
        " L" + g12 + "," + g41 +
        " L" + g38 + "," + g41 +
        " L" + g37 + "," + g42 +
        " L" + g35 + "," + g42 +
        " L" + g35 + "," + g41 +
        " L" + g34 + "," + g40 +
        " L" + g32 + "," + g40 +
        " L" + g31 + "," + g39 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：返回
function genActionButtonReturn(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 * 7 / 8;
    var g15 = g13 * 3 / 4;
    var g16 = g13 * 5 / 8;
    var g17 = g13 * 3 / 8;
    var g18 = g13 / 4;
    var g19 = g9 + g15;
    var g20 = g9 + g16;
    var g21 = g9 + g18;
    var g22 = g11 + g14;
    var g23 = g11 + g15;
    var g24 = g11 + g16;
    var g25 = g11 + g17;
    var g26 = g11 + g18;
    var g27 = g13 / 8;
    var cX1 = g24 - g27;
    var cY2 = g19 - g27;
    var cX3 = g11 + g17;
    var cY4 = g10 - g17;

    var d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + g21 +
        " L" + g23 + "," + g9 +
        " L" + hc + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g24 + "," + g20 +
        window.PPTXShapeUtils.shapeArc(cX1, g20, g27, g27, 0, 90, false).replace("M", "L") +
        " L" + g25 + "," + g19 +
        window.PPTXShapeUtils.shapeArc(g25, cY2, g27, g27, 90, 180, false).replace("M", "L") +
        " L" + g26 + "," + g21 +
        " L" + g11 + "," + g21 +
        " L" + g11 + "," + g20 +
        window.PPTXShapeUtils.shapeArc(cX3, g20, g17, g17, 180, 90, false).replace("M", "L") +
        " L" + hc + "," + g10 +
        window.PPTXShapeUtils.shapeArc(hc, cY4, g17, g17, 90, 0, false).replace("M", "L") +
        " L" + g22 + "," + g21 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

// 操作按钮：声音
function genActionButtonSound(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    var hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    var dx2 = ss * 3 / 8;
    var g9 = vc - dx2;
    var g10 = vc + dx2;
    var g11 = hc - dx2;
    var g12 = hc + dx2;
    var g13 = ss * 3 / 4;
    var g14 = g13 / 8;
    var g15 = g13 * 5 / 16;
    var g16 = g13 * 5 / 8;
    var g17 = g13 * 11 / 16;
    var g18 = g13 * 3 / 4;
    var g19 = g13 * 7 / 8;
    var g20 = g9 + g14;
    var g21 = g9 + g15;
    var g22 = g9 + g17;
    var g23 = g9 + g19;
    var g24 = g11 + g15;
    var g25 = g11 + g16;
    var g26 = g11 + g18;

    var d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        " M" + g11 + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g25 + "," + g9 +
        " L" + g25 + "," + g10 +
        " L" + g24 + "," + g22 +
        " L" + g11 + "," + g22 +
        " z" +
        " M" + g26 + "," + g21 +
        " L" + g12 + "," + g20 +
        " M" + g26 + "," + vc +
        " L" + g12 + "," + vc +
        " M" + g26 + "," + g22 +
        " L" + g12 + "," + g23;
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

const PPTXActionButtonShapes = {
    getActionButtonRect,
    genPath,
    genActionButtonBackPrevious,
    genActionButtonBeginning,
    genActionButtonDocument,
    genActionButtonEnd,
    genActionButtonForwardNext,
    genActionButtonHelp,
    genActionButtonHome,
    genActionButtonInformation,
    genActionButtonMovie,
    genActionButtonReturn,
    genActionButtonSound
}

export { PPTXActionButtonShapes };

// Also export to global scope for backward compatibility
// window.PPTXActionButtonShapes = PPTXActionButtonShapes; // Removed for ES modules
