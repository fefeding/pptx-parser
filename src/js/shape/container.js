import { PPTXUtils } from '../core/utils.js';
import { PPTXShapeFillsUtils } from './fills.js';
// 辅助函数：获取形状变换参数（flip, rotate等）
function getShapeTransformParams(node) {
    var result = {
        isFlipV: false,
        isFlipH: false,
        flip: "",
        rotate: undefined,
        txtRotate: undefined
    };

    var xfrmList = ["p:spPr", "a:xfrm"];
    var slideXfrmNode = PPTXUtils.getTextByPathList(node, xfrmList);

    if (PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1") {
        result.isFlipV = true;
    }
    if (PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
        result.isFlipH = true;
    }

    if (result.isFlipH && !result.isFlipV) {
        result.flip = " scale(-1,1)";
    } else if (!result.isFlipH && result.isFlipV) {
        result.flip = " scale(1,-1)";
    } else if (result.isFlipH && result.isFlipV) {
        result.flip = " scale(-1,-1)";
    }

    result.rotate = PPTXUtils.angleToDegrees(PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));

    var txtXframeNode = PPTXUtils.getTextByPathList(node, ["p:txXfrm"]);
    if (txtXframeNode !== undefined) {
        var txtXframeRot = PPTXUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
        if (txtXframeRot !== undefined) {
            result.txtRotate = PPTXUtils.angleToDegrees(txtXframeRot) + 90;
        }
    } else {
        result.txtRotate = result.rotate;
    }

    return result;
}

    // 辅助函数：生成 SVG 容器开标签
function getSvgContainerStart(shpId, id, idx, type, name, w, h, svgCssName, effectsClassName, rotate, flip, order, slideXfrmNode, pNode, getPosition, getSize, sType) {
    var result = "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "'" +
        " style='" +
        getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
        getSize(slideXfrmNode, undefined, undefined) +
        " z-index: " + order + ";" +
        "transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + ";" +
        "'>";

    result += '<defs>';
    return result;
}

    // 辅助函数：生成填充属性字符串（基于填充类型）
function getFillAttr(fillColor, imgFillFlg, grndFillFlg, shpId, w, h, clrFillType, warpObj, node) {
    if (clrFillType == "PIC_FILL") {
        return "url(#imgPtrn_" + shpId + ")";
    } else if (clrFillType == "GRADIENT_FILL") {
        return "url(#linGrd_" + shpId + ")";
    } else {
        return fillColor;
    }
}

    // 辅助函数：生成填充属性字符串（基于标志变量）
function getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId) {
    if (imgFillFlg) {
        return "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        return "url(#linGrd_" + shpId + ")";
    } else {
        return fillColor;
    }
}

    // 辅助函数：生成形状元素的通用属性
function getShapeAttributes(w, h, shpId, fillColor, imgFillFlg, grndFillFlg, border, oShadowSvgUrlStr, shapType) {
    var fillAttr;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }

    var result = "fill='" + fillAttr + "'";
    result += " stroke='" + border.color + "'";
    result += " stroke-width='" + border.width + "'";
    result += " stroke-dasharray='" + border.strokeDasharray + "'";
    if (oShadowSvgUrlStr) {
        result += " " + oShadowSvgUrlStr;
    }

    return result;
}

    // 辅助函数：生成箭头标记
function getTriangleMarker(shpId, border, headEndNodeAttrs, tailEndNodeAttrs) {
    if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
        (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
        return "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
            "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
    }
    return "";
}

    // 辅助函数：处理阴影效果
function processShadowEffect(outerShdwNode, slideFactor, styleTable, effectsClassName, warpObj) {
    var svg_css_shadow = "";
    if (outerShdwNode !== undefined) {
        var chdwClrNode = PPTXColorUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
        var outerShdwAttrs = outerShdwNode["attrs"];

        var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
        var dist = parseInt(outerShdwAttrs["dist"]) * slideFactor;
        var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor) : "";

        var vx = dist * Math.sin(dir * Math.PI / 180);
        var hx = dist * Math.cos(dir * Math.PI / 180);

        svg_css_shadow = "filter:drop-shadow(" + hx + "px " + vx + "px " + blurRad + "px #" + chdwClrNode + ");";

        if (svg_css_shadow in styleTable) {
            svg_css_shadow += "do-nothing: " + effectsClassName + ";";
        }

        styleTable[svg_css_shadow] = {
            "name": effectsClassName,
            "text": svg_css_shadow
        };
    }
    return "";
}

    // 辅助函数：生成 defs 内容（渐变、图案填充等）
function getDefsContent(node, pNode, warpObj, source, shpId, w, h, fillColor, styleTable, svgCssName, headEndNodeAttrs, tailEndNodeAttrs, border) {
    var result = "";
    var grndFillFlg = false;
    var imgFillFlg = false;

    var clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    if (clrFillType == "GROUP_FILL") {
        clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
    }

    // 处理渐变填充
    if (clrFillType == "GRADIENT_FILL") {
        grndFillFlg = true;
        var color_arry = fillColor.color;
        var angl = fillColor.rot + 90;
        var svgGrdnt = PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, shpId);
        result += svgGrdnt;
    }
    // 处理图片填充
    else if (clrFillType == "PIC_FILL") {
        imgFillFlg = true;
        var imgFill = typeof fillColor === 'object' && fillColor.img ? fillColor.img : fillColor;
        var svgBgImg = PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, shpId, warpObj);
        result += svgBgImg;
    }
    // 处理图案填充
    else if (clrFillType == "PATTERN_FILL") {
        var styleText = fillColor;
        if (styleText in styleTable) {
            styleText += "do-nothing: " + svgCssName + ";";
        }
        styleTable[styleText] = {
            "name": svgCssName,
            "text": styleText
        };
    }

    // 处理箭头标记
    result += getTriangleMarker(shpId, border, headEndNodeAttrs, tailEndNodeAttrs);

    return {
        content: result,
        grndFillFlg: grndFillFlg,
        imgFillFlg: imgFillFlg,
        clrFillType: clrFillType,
        fillColor: fillColor
    };
}

const PPTXShapeContainer = {
    getShapeTransformParams,
    getSvgContainerStart,
    getFillAttr,
    getFillAttrFromFlags,
    getShapeAttributes,
    getTriangleMarker,
    processShadowEffect,
    getDefsContent
};

export { PPTXShapeContainer };

// Also export to global scope for backward compatibility
// window.PPTXShapeContainer = PPTXShapeContainer; // Removed for ES modules
