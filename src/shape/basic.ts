interface Border {
    color: string;
    width: number;
    strokeDasharray: string;
}

// 生成矩形
function genRect(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border, oShadowSvgUrlStr: string): string {
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    return "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' " + oShadowSvgUrlStr + " />";
}

// 生成椭圆
function genEllipse(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    return "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成菱形
function genDiamond(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    return "<polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成三角形
function genTriangle(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    return "<polygon points='" + (w / 2) + " 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成圆形
function genCircle(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    const r: number = Math.min(w, h) / 2;
    const cx: number = w / 2;
    const cy: number = h / 2;
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    return "<circle cx='" + cx + "' cy='" + cy + "' r='" + r + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成五边形
function genPentagon(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    const hc: number = w / 2, vc: number = h / 2, r: number = Math.min(w, h) / 2;
    let dx2: number = r * Math.sin(Math.PI * 2 / 5);
    let dy2: number = r * Math.cos(Math.PI * 2 / 5);
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    let d: string = "M" + hc + "," + (vc - r) +
        " L" + (hc + dx2) + "," + (vc - dy2) +
        " L" + (hc + dx2) + "," + (vc + dy2) +
        " L" + (hc - dx2) + "," + (vc + dy2) +
        " L" + (hc - dx2) + "," + (vc - dy2) +
        " z";
    return "<path d='" + d + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成六边形
function genHexagon(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border): string {
    const hc: number = w / 2, vc: number = h / 2, r: number = Math.min(w, h) / 2;
    const dx2: number = r * Math.sin(Math.PI / 3);
    const dy2: number = r * Math.cos(Math.PI / 3);
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }
    const d: string = "M" + hc + "," + (vc - r) +
        " L" + (hc + dx2) + "," + (vc - dy2) +
        " L" + (hc + dx2) + "," + (vc + dy2) +
        " L" + hc + "," + (vc + r) +
        " L" + (hc - dx2) + "," + (vc + dy2) +
        " L" + (hc - dx2) + "," + (vc - dy2) +
        " z";
    return "<path d='" + d + "' fill='" + fillAttr +
        "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

// 生成带特殊标记的矩形（用于流程图）
function genRectWithDecoration(w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, fillColor: string, border: Border, oShadowSvgUrlStr: string, shapType: string): string {
    let result: string = genRect(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr);

    if (shapType == "flowChartPredefinedProcess") {
        result += "<rect x='" + w * (1 / 8) + "' y='0' width='" + w * (6 / 8) + "' height='" + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    } else if (shapType == "flowChartInternalStorage") {
        result += " <polyline points='" + w * (1 / 8) + " 0," + w * (1 / 8) + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
        result += " <polyline points='0 " + h * (1 / 8) + "," + w + " " + h * (1 / 8) + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
    }

    return result;
}

const PPTXBasicShapes = {
    genRect,
    genEllipse,
    genDiamond,
    genTriangle,
    genCircle,
    genPentagon,
    genHexagon,
    genRectWithDecoration
};

export { PPTXBasicShapes };