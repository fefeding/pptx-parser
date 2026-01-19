    import { PPTXShapeUtils } from './shape.js';
    // 辅助函数：生成操作按钮的基础矩形
function getActionButtonRect(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    const fillAttr = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
    return "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

    // 辅助函数：生成路径元素
function genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const fillAttr: any = imgFillFlg ? "url(#imgPtrn_" + shpId + ")" : (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor);
    return "<path d='" + d + "'  fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

    // 操作按钮：后退/上一个
function genActionButtonBackPrevious(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    const hc = w / 2, vc = h / 2, ss = Math.min(w, h);
    let dx2 = ss * 3 / 8;
    let g9 = vc - dx2;
    let g10 = vc + dx2;
    let g11 = hc - dx2;
    let g12 = hc + dx2;
    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` zM` + g11 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：开始
function genActionButtonBeginning(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
    let g13 = ss * 3 / 4;
    let g14 = g13 / 8;
    let g15 = g13 / 4;
    let g16 = g11 + g14;
    let g17 = g11 + g15;
    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` zM` + g17 + "," + vc +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        ` zM` + g16 + "," + g9 +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " L" + g16 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：文档
function genActionButtonDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
    const g9: any = vc - dx2;
    const g10: any = vc + dx2;
    let dx1 = ss * 9 / 32;
    const g11: any = hc - dx1;
    const g12: any = hc + dx1;
    const g13: any = ss * 3 / 16;
    const g14: any = g12 - g13;
    const g15: any = g9 + g13;
    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` zM` + g11 + "," + g9 +
        " L" + g14 + "," + g9 +
        " L" + g12 + "," + g15 +
        " L" + g12 + "," + g10 +
        " L" + g11 + "," + g10 +
        ` zM` + g14 + "," + g9 +
        " L" + g14 + "," + g15 +
        " L" + g12 + "," + g15 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：结束
function genActionButtonEnd(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
    const g11: any = hc - dx2;
    const g12: any = hc + dx2;
    const g13: any = ss * 3 / 4;
    const g14: any = g13 * 3 / 4;
    const g15: any = g13 * 7 / 8;
    const g16: any = g11 + g14;
    const g17: any = g11 + g15;
    let d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        ` z M` + g17 + "," + g9 +
        " L" + g12 + "," + g9 +
        " L" + g12 + "," + g10 +
        " L" + g17 + "," + g10 +
        ` z M` + g16 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：前进/下一个
function genActionButtonForwardNext(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
let d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        ` z M` + g12 + "," + vc +
        " L" + g11 + "," + g9 +
        " L" + g11 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：帮助
function genActionButtonHelp(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
    const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g11: any = hc - dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 / 7;
const g15: any = g13 * 3 / 14;
const g16: any = g13 * 2 / 7;
    let g19 = g13 * 3 / 7;
    let g20 = g13 * 4 / 7;
    let g21 = g13 * 17 / 28;
    let g23 = g13 * 21 / 28;
    let g24 = g13 * 11 / 14;
    let g27 = g9 + g16;
    let g29 = g9 + g21;
    let g30 = g9 + g23;
    let g31 = g9 + g24;
    let g33 = g11 + g15;
    let g36 = g11 + g19;
    let g37 = g11 + g20;
    let g41 = g13 / 14;
    let g42 = g13 * 3 / 28;
    const cX1 = g33 + g16;
    const cX2 = g36 + g14;
    const cY3 = g31 + g42;
    const cX4 = (g37 + g36 + g16) / 2;

    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` zM` + g33 + "," + g27 +
        PPTXShapeUtils.shapeArc(cX1, g27, g16, g16, 180, 360, false).replace("M", "L") +
        PPTXShapeUtils.shapeArc(cX4, g27, g14, g15, 0, 90, false).replace("M", "L") +
        PPTXShapeUtils.shapeArc(cX4, g29, g41, g42, 270, 180, false).replace("M", "L") +
        " L" + g37 + "," + g30 +
        " L" + g36 + "," + g30 +
        " L" + g36 + "," + g29 +
        PPTXShapeUtils.shapeArc(cX2, g29, g14, g15, 180, 270, false).replace("M", "L") +
        PPTXShapeUtils.shapeArc(g37, g27, g41, g42, 90, 0, false).replace("M", "L") +
        PPTXShapeUtils.shapeArc(cX1, g27, g14, g14, 0, -180, false).replace("M", "L") +
        ` zM` + hc + "," + g31 +
        PPTXShapeUtils.shapeArc(hc, cY3, g42, g42, 270, 630, false).replace("M", "L") +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：主页
function genActionButtonHome(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 / 16;
const g15: any = g13 / 8;
const g16: any = g13 * 3 / 16;
const g17: any = g13 * 5 / 16;
    let g18 = g13 * 7 / 16;
const g19: any = g13 * 9 / 16;
const g20: any = g13 * 11 / 16;
const g21: any = g13 * 3 / 4;
    let g22 = g13 * 13 / 16;
const g23: any = g13 * 7 / 8;
const g24: any = g9 + g14;
    let g25 = g9 + g16;
    let g26 = g9 + g17;
const g27: any = g9 + g21;
    let g28 = g11 + g15;
const g29: any = g11 + g18;
const g30: any = g11 + g19;
    const g31: any = g11 + g20;
    let g32 = g11 + g22;
    const g33: any = g11 + g23;
    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` z M` + hc + "," + g9 +
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
        ` z M` + g29 + "," + g27 +
        " L" + g30 + "," + g27 +
        " L" + g30 + "," + g10 +
        " L" + g29 + "," + g10 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

    // 操作按钮：信息
function genActionButtonInformation(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g11: any = hc - dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 / 32;
const g17: any = g13 * 5 / 16;
const g18: any = g13 * 3 / 8;
const g19: any = g13 * 13 / 32;
const g20: any = g13 * 19 / 32;
const g22: any = g13 * 11 / 16;
const g23: any = g13 * 13 / 16;
const g24: any = g13 * 7 / 8;
const g25: any = g9 + g14;
const g28: any = g9 + g17;
const g29: any = g9 + g18;
const g30: any = g9 + g23;
const g31: any = g9 + g24;
const g32: any = g11 + g17;
    let g34 = g11 + g19;
    let g35 = g11 + g20;
    const g37: any = g11 + g22;
    let g38 = g13 * 3 / 32;
    const cY1 = g9 + dx2;
    const cY2 = g25 + g38;
    let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` zM` + hc + "," + g9 +
        PPTXShapeUtils.shapeArc(hc, cY1, dx2, dx2, 270, 630, false).replace("M", "L") +
        ` zM` + hc + "," + g25 +
        PPTXShapeUtils.shapeArc(hc, cY2, g38, g38, 270, 630, false).replace("M", "L") +
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
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 * 1455 / 21600;
const g15: any = g13 * 1905 / 21600;
const g16: any = g13 * 2325 / 21600;
const g17: any = g13 * 16155 / 21600;
const g18: any = g13 * 17010 / 21600;
const g19: any = g13 * 19335 / 21600;
const g20: any = g13 * 19725 / 21600;
const g21: any = g13 * 20595 / 21600;
const g22: any = g13 * 5280 / 21600;
const g23: any = g13 * 5730 / 21600;
const g24: any = g13 * 6630 / 21600;
const g25: any = g13 * 7492 / 21600;
const g26: any = g13 * 9067 / 21600;
const g27: any = g13 * 9555 / 21600;
const g28: any = g13 * 13342 / 21600;
const g29: any = g13 * 14580 / 21600;
const g30: any = g13 * 15592 / 21600;
const g31: any = g11 + g14;
const g32: any = g11 + g15;
const g33: any = g11 + g16;
const g34: any = g11 + g17;
const g35: any = g11 + g18;
const g36: any = g11 + g19;
const g37: any = g11 + g20;
const g38: any = g11 + g21;
    let g39 = g9 + g22;
    let g40 = g9 + g23;
const g41: any = g9 + g24;
const g42: any = g9 + g25;
    let g43 = g9 + g26;
    let g44 = g9 + g27;
    let g45 = g9 + g28;
    let g46 = g9 + g29;
    let g47 = g9 + g30;
    let g48 = g9 + g31;

    let d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        ` zM` + g11 + "," + g39 +
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
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 * 7 / 8;
const g15: any = g13 * 3 / 4;
const g16: any = g13 * 5 / 8;
const g17: any = g13 * 3 / 8;
const g18: any = g13 / 4;
const g19: any = g9 + g15;
const g20: any = g9 + g16;
const g21: any = g9 + g18;
const g22: any = g11 + g14;
const g23: any = g11 + g15;
const g24: any = g11 + g16;
const g25: any = g11 + g17;
const g26: any = g11 + g18;
const g27: any = g13 / 8;
const cX1: any = g24 - g27;
const cY2: any = g19 - g27;
    const cX3 = g11 + g17;
    const cY4 = g10 - g17;

    let d = "M" + 0 + "," + h +
        " L" + w + "," + h +
        " L" + w + "," + 0 +
        " L" + 0 + "," + 0 +
        ` z M` + g12 + "," + g21 +
        " L" + g23 + "," + g9 +
        " L" + hc + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g24 + "," + g20 +
        PPTXShapeUtils.shapeArc(cX1, g20, g27, g27, 0, 90, false).replace("M", "L") +
        " L" + g25 + "," + g19 +
        PPTXShapeUtils.shapeArc(g25, cY2, g27, g27, 90, 180, false).replace("M", "L") +
        " L" + g26 + "," + g21 +
        " L" + g11 + "," + g21 +
        " L" + g11 + "," + g20 +
        PPTXShapeUtils.shapeArc(cX3, g20, g17, g17, 180, 90, false).replace("M", "L") +
        " L" + hc + "," + g10 +
        PPTXShapeUtils.shapeArc(hc, cY4, g17, g17, 90, 0, false).replace("M", "L") +
        " L" + g22 + "," + g21 +
        " z";
    return genPath(d, imgFillFlg, grndFillFlg, shpId, fillColor, border);
}

// 操作按钮：声音
function genActionButtonSound(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border) {
const hc: number = w / 2, vc: number = h / 2, ss: number = Math.min(w, h);
const dx2: any = ss * 3 / 8;
const g9: any = vc - dx2;
const g10: any = vc + dx2;
const g11: any = hc - dx2;
const g12: any = hc + dx2;
const g13: any = ss * 3 / 4;
const g14: any = g13 / 8;
const g15: any = g13 * 5 / 16;
const g16: any = g13 * 5 / 8;
const g17: any = g13 * 11 / 16;
const g18: any = g13 * 3 / 4;
const g19: any = g13 * 7 / 8;
const g20: any = g9 + g14;
const g21: any = g9 + g15;
const g22: any = g9 + g17;
const g23: any = g9 + g19;
const g24: any = g11 + g15;
const g25: any = g11 + g16;
const g26: any = g11 + g18;

let d = "M" + 0 + "," + 0 +
        " L" + w + "," + 0 +
        " L" + w + "," + h +
        " L" + 0 + "," + h +
        ` z M` + g11 + "," + g21 +
        " L" + g24 + "," + g21 +
        " L" + g25 + "," + g9 +
        " L" + g25 + "," + g10 +
        " L" + g24 + "," + g22 +
        " L" + g11 + "," + g22 +
        ` z M` + g26 + "," + g21 +
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
