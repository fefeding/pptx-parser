/**
 * 按钮类形状渲染模块
 * 处理所有 actionButton 类型的形状渲染
 */

/**
 * 渲染 actionButtonBackPrevious 形状
 * 返回按钮（左箭头）
 */
function renderBackPrevious(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g11 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonBeginning 形状
 * 开始按钮（双竖线+左箭头）
 */
function renderBeginning(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 / 4;
    const g16 = g11 + g14;
    const g17 = g11 + g15;

    const d = "M" + 0 + "," + 0 +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonDocument 形状
 * 文档按钮
 */
function renderDocument(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const dx1 = ss * 9 / 32;
    const g11 = hc - dx1;
    const g12 = hc + dx1;
    const g13 = ss * 3 / 16;
    const g14 = g12 - g13;
    const g15 = g9 + g13;

    const d = "M" + 0 + "," + 0 +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonEnd 形状
 * 结束按钮（双竖线+右箭头）
 */
function renderEnd(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 3 / 4;
    const g15 = g13 * 7 / 8;
    const g16 = g11 + g14;
    const g17 = g11 + g15;

    const d = "M" + 0 + "," + h +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonForwardNext 形状
 * 前进按钮（右箭头）
 */
function renderForwardNext(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonHelp 形状
 * 帮助按钮（问号）
 */
function renderHelp(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 7;
    const g15 = g13 * 3 / 14;
    const g16 = g13 * 2 / 7;
    const g19 = g13 * 3 / 7;
    const g20 = g13 * 4 / 7;
    const g21 = g13 * 17 / 28;
    const g23 = g13 * 21 / 28;
    const g24 = g13 * 11 / 14;
    const g27 = g9 + g16;
    const g29 = g9 + g21;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g33 = g11 + g15;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g41 = g13 / 14;
    const g42 = g13 * 3 / 28;
    const cX1 = g33 + g16;
    const cX2 = g36 + g14;
    const cY3 = g31 + g42;
    const cX4 = (g37 + g36 + g16) / 2;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + g33 + "," + g27 +
        shapeArcAlt(cX1, g27, g16, g16, 180, 360, false).replace("M", "L") +
        shapeArcAlt(cX4, g27, g14, g15, 0, 90, false).replace("M", "L") +
        shapeArcAlt(cX4, g29, g41, g42, 270, 180, false).replace("M", "L") +
        " L" + g37 + "," + g30 +
        " L" + g36 + "," + g30 +
        " L" + g36 + "," + g29 +
        shapeArcAlt(cX2, g29, g14, g15, 180, 270, false).replace("M", "L") +
        shapeArcAlt(g37, g27, g41, g42, 90, 0, false).replace("M", "L") +
        shapeArcAlt(cX1, g27, g14, g14, 0, -180, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g31 +
        shapeArcAlt(hc, cY3, g42, g42, 270, 630, false).replace("M", "L") +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonHome 形状
 * 主页按钮（房子图标）
 */
function renderHome(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 16;
    const g15 = g13 / 8;
    const g16 = g13 * 3 / 16;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 7 / 16;
    const g19 = g13 * 9 / 16;
    const g20 = g13 * 11 / 16;
    const g21 = g13 * 3 / 4;
    const g22 = g13 * 13 / 16;
    const g23 = g13 * 7 / 8;
    const g24 = g9 + g14;
    const g25 = g9 + g16;
    const g26 = g9 + g17;
    const g27 = g9 + g21;
    const g28 = g11 + g15;
    const g29 = g11 + g18;
    const g30 = g11 + g19;
    const g31 = g11 + g20;
    const g32 = g11 + g22;
    const g33 = g11 + g23;

    const d = "M" + 0 + "," + 0 +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonInformation 形状
 * 信息按钮（感叹号+圆点）
 */
function renderInformation(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 32;
    const g17 = g13 * 5 / 16;
    const g18 = g13 * 3 / 8;
    const g19 = g13 * 13 / 32;
    const g20 = g13 * 19 / 32;
    const g22 = g13 * 11 / 16;
    const g23 = g13 * 13 / 16;
    const g24 = g13 * 7 / 8;
    const g25 = g9 + g14;
    const g28 = g9 + g17;
    const g29 = g9 + g18;
    const g30 = g9 + g23;
    const g31 = g9 + g24;
    const g32 = g11 + g17;
    const g34 = g11 + g19;
    const g35 = g11 + g20;
    const g37 = g11 + g22;
    const g38 = g13 * 3 / 32;
    const cY1 = g9 + dx2;
    const cY2 = g25 + g38;

    const d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        " z" +
        "M" + hc + "," + g9 +
        shapeArcAlt(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L") +
        " z" +
        "M" + hc + "," + g25 +
        shapeArcAlt(hc, cY2, g38, g38, 270, 630, false).replace("M", "L") +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonMovie 形状
 * 影片按钮（胶片图标）
 */
function renderMovie(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 1455 / 21600;
    const g15 = g13 * 1905 / 21600;
    const g16 = g13 * 2325 / 21600;
    const g17 = g13 * 16155 / 21600;
    const g18 = g13 * 17010 / 21600;
    const g19 = g13 * 19335 / 21600;
    const g20 = g13 * 19725 / 21600;
    const g21 = g13 * 20595 / 21600;
    const g22 = g13 * 5280 / 21600;
    const g23 = g13 * 5730 / 21600;
    const g24 = g13 * 6630 / 21600;
    const g25 = g13 * 7492 / 21600;
    const g26 = g13 * 9067 / 21600;
    const g27 = g13 * 9555 / 21600;
    const g28 = g13 * 13342 / 21600;
    const g29 = g13 * 14580 / 21600;
    const g30 = g13 * 15592 / 21600;
    const g31 = g11 + g14;
    const g32 = g11 + g15;
    const g33 = g11 + g16;
    const g34 = g11 + g17;
    const g35 = g11 + g18;
    const g36 = g11 + g19;
    const g37 = g11 + g20;
    const g38 = g11 + g21;
    const g39 = g9 + g22;
    const g40 = g9 + g23;
    const g41 = g9 + g24;
    const g42 = g9 + g25;
    const g43 = g9 + g26;
    const g44 = g9 + g27;
    const g45 = g9 + g28;
    const g46 = g9 + g29;
    const g47 = g9 + g30;
    const g48 = g9 + g31;

    const d = "M" + 0 + "," + h +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonReturn 形状
 * 返回按钮（折返箭头）
 */
function renderReturn(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 * 7 / 8;
    const g15 = g13 * 3 / 4;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 3 / 8;
    const g18 = g13 / 4;
    const g19 = g9 + g15;
    const g20 = g9 + g16;
    const g21 = g9 + g18;
    const g22 = g11 + g14;
    const g23 = g11 + g15;
    const g24 = g11 + g16;
    const g25 = g11 + g17;
    const g26 = g11 + g18;
    const g27 = g13 / 8;
    const cX1 = g24 - g27;
    const cY2 = g19 - g27;
    const cX3 = g11 + g17;
    const cY4 = g10 - g17;

    const d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        " z" +
        " M" + g12 + "," + g21 +
        " L" + g23 + "," + g9 +
        " L" + hc + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g24 + "," + g20 +
        shapeArcAlt(cX1, g20, g27, g27, 0, 90, false).replace("M", "L") +
        " L" + g25 + "," + g19 +
        shapeArcAlt(g25, cY2, g27, g27, 90, 180, false).replace("M", "L") +
        " L" + g26 + "," + g21 +
        " L" + g11 + "," + g21 +
        " L" + g11 + "," + g20 +
        shapeArcAlt(cX3, g20, g17, g17, 180, 90, false).replace("M", "L") +
        " L" + hc + "," + g10 +
        shapeArcAlt(hc, cY4, g17, g17, 90, 0, false).replace("M", "L") +
        " L" + g22 + "," + g21 +
        " z";

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染 actionButtonSound 形状
 * 声音按钮（喇叭图标）
 */
function renderSound(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    const dx2 = ss * 3 / 8;
    const g9 = vc - dx2;
    const g10 = vc + dx2;
    const g11 = hc - dx2;
    const g12 = hc + dx2;
    const g13 = ss * 3 / 4;
    const g14 = g13 / 8;
    const g15 = g13 * 5 / 16;
    const g16 = g13 * 5 / 8;
    const g17 = g13 * 11 / 16;
    const g18 = g13 * 3 / 4;
    const g19 = g13 * 7 / 8;
    const g20 = g9 + g14;
    const g21 = g9 + g15;
    const g22 = g9 + g17;
    const g23 = g9 + g19;
    const g24 = g11 + g15;
    const g25 = g11 + g16;
    const g26 = g11 + g18;

    const d = "M" + 0 + "," + 0 +
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

    return "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 按钮形状渲染器映射表
 */
const BUTTON_RENDERERS = {
    'actionButtonBackPrevious': renderBackPrevious,
    'actionButtonBeginning': renderBeginning,
    'actionButtonDocument': renderDocument,
    'actionButtonEnd': renderEnd,
    'actionButtonForwardNext': renderForwardNext,
    'actionButtonHelp': renderHelp,
    'actionButtonHome': renderHome,
    'actionButtonInformation': renderInformation,
    'actionButtonMovie': renderMovie,
    'actionButtonReturn': renderReturn,
    'actionButtonSound': renderSound
};

/**
 * 渲染按钮类形状
 * @param {string} shapeType - 按钮类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 图片填充标志
 * @param {boolean} grndFillFlg - 渐变填充标志
 * @param {string} fillColor - 填充颜色
 * @param {object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Function} shapeArcAlt - 弧形生成函数
 * @returns {string} SVG 路径字符串
 */
export function renderActionButton(shapeType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt) {
    const renderer = BUTTON_RENDERERS[shapeType];
    if (!renderer) {
        return '';
    }
    return renderer(w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
}

/**
 * 检查是否为按钮类形状
 * @param {string} shapeType - 形状类型
 * @returns {boolean}
 */
export function isActionButton(shapeType) {
    return BUTTON_RENDERERS.hasOwnProperty(shapeType);
}
