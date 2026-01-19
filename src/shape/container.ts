import { PPTXUtils } from '../core/utils';
import { PPTXShapeFillsUtils } from './fills';

// 定义接口
interface TransformParams {
    isFlipV: boolean;
    isFlipH: boolean;
    flip: string;
    rotate: number | undefined;
    txtRotate: number | undefined;
}

interface DefsResult {
    content: string;
    grndFillFlg: boolean;
    imgFillFlg: boolean;
    clrFillType: string;
    fillColor: any;
}

interface Border {
    color: string;
    width: number;
    strokeDasharray: string;
}

// 辅助函数：获取形状变换参数（flip, rotate等）
function getShapeTransformParams(node: any): TransformParams {
    let result: TransformParams = {
        isFlipV: false,
        isFlipH: false,
        flip: "",
        rotate: undefined,
        txtRotate: undefined
    };

    const xfrmList: string[] = ["p:spPr", "a:xfrm"];
    const slideXfrmNode: any = PPTXUtils.getTextByPathList(node, xfrmList);

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

    const txtXframeNode: any = PPTXUtils.getTextByPathList(node, ["p:txXfrm"]);
    if (txtXframeNode !== undefined) {
        const txtXframeRot: any = PPTXUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
        if (txtXframeRot !== undefined) {
            result.txtRotate = PPTXUtils.angleToDegrees(txtXframeRot) + 90;
        }
    } else {
        result.txtRotate = result.rotate;
    }

    return result;
}

    // 辅助函数：生成 SVG 容器开标签
function getSvgContainerStart(shpId: string | number, id: string, idx: string, type: string, name: string, w: number, h: number, svgCssName: string, effectsClassName: string, rotate: number | undefined, flip: string, order: number, slideXfrmNode: any, pNode: any, getPosition: Function, getSize: Function, sType: string): string {
    let result: string = "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + `' style='` +
        getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
        getSize(slideXfrmNode, undefined, undefined) +
        " z-index: " + order + `;transform: rotate(` + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + `;'>`;

    result += '<defs>';
    return result;
}

    // 辅助函数：生成填充属性字符串（基于填充类型）
function getFillAttr(fillColor: string, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number, w: number, h: number, clrFillType: string, warpObj: any, node: any): string {
    if (clrFillType == "PIC_FILL") {
        return "url(#imgPtrn_" + shpId + ")";
    } else if (clrFillType == "GRADIENT_FILL") {
        return "url(#linGrd_" + shpId + ")";
    } else {
        return fillColor;
    }
}

    // 辅助函数：生成填充属性字符串（基于标志变量）
function getFillAttrFromFlags(fillColor: string, imgFillFlg: boolean, grndFillFlg: boolean, shpId: string | number): string {
    if (imgFillFlg) {
        return "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        return "url(#linGrd_" + shpId + ")";
    } else {
        return fillColor;
    }
}

    // 辅助函数：生成形状元素的通用属性
function getShapeAttributes(w: number, h: number, shpId: string | number, fillColor: string, imgFillFlg: boolean, grndFillFlg: boolean, border: Border, oShadowSvgUrlStr: string, shapType: string): string {
    let fillAttr: string;
    if (imgFillFlg) {
        fillAttr = "url(#imgPtrn_" + shpId + ")";
    } else if (grndFillFlg) {
        fillAttr = "url(#linGrd_" + shpId + ")";
    } else {
        fillAttr = fillColor;
    }

    let result: string = "fill='" + fillAttr + "'";
    result += " stroke='" + border.color + "'";
    result += " stroke-width='" + border.width + "'";
    result += " stroke-dasharray='" + border.strokeDasharray + "'";
    if (oShadowSvgUrlStr) {
        result += " " + oShadowSvgUrlStr;
    }

    return result;
}

    // 辅助函数：生成箭头标记
function getTriangleMarker(shpId: string | number, border: Border, headEndNodeAttrs: any, tailEndNodeAttrs: any): string {
    if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
        (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
        return "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
            "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
    }
    return "";
}

    // 辅助函数：处理阴影效果
function processShadowEffect(outerShdwNode: any, slideFactor: number, styleTable: any, effectsClassName: string, warpObj: any): string {
    let svg_css_shadow: string = "";
    if (outerShdwNode !== undefined) {
        // Assuming PPTXColorUtils is available, perhaps import or define
        const chdwClrNode: any = (globalThis as any).PPTXColorUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
        const outerShdwAttrs: any = outerShdwNode["attrs"];

        let dir: number = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
        const dist: number = parseInt(outerShdwAttrs["dist"]) * slideFactor;
        const blurRad: string = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor).toString() : "";

        const vx: number = dist * Math.sin(dir * Math.PI / 180);
        const hx: number = dist * Math.cos(dir * Math.PI / 180);

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
function getDefsContent(node: any, pNode: any, warpObj: any, source: any, shpId: string | number, w: number, h: number, fillColor: any, styleTable: any, svgCssName: string, headEndNodeAttrs: any, tailEndNodeAttrs: any, border: Border): DefsResult {
    let result: string = "";
    let grndFillFlg: boolean = false;
    let imgFillFlg: boolean = false;

    // Assuming PPTXColorUtils is available
    let clrFillType: string = (globalThis as any).PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    if (clrFillType == "GROUP_FILL") {
        clrFillType = (globalThis as any).PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
    }

    // 处理渐变填充
    if (clrFillType == "GRADIENT_FILL") {
        grndFillFlg = true;
        const color_arry: any = fillColor.color;
        const angl: number = fillColor.rot + 90;
        const svgGrdnt: string = PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, String(shpId));
        result += svgGrdnt;
    }
    // 处理图片填充
    else if (clrFillType == "PIC_FILL") {
        imgFillFlg = true;
        const imgFill: any = typeof fillColor === 'object' && fillColor.img ? fillColor.img : fillColor;
        const svgBgImg: string = PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, String(shpId), warpObj);
        result += svgBgImg;
    }
    // 处理图案填充
    else if (clrFillType == "PATTERN_FILL") {
        const styleText: string = fillColor;
        if (styleText in styleTable) {
            // Note: this seems like a bug in original, probably should be styleTable[styleText] += ...
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