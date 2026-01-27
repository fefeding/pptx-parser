import { PPTXUtils } from '../core/utils.js';
import { PPTXColorUtils } from '../core/color.js';
import tinycolor from 'tinycolor2';

interface PPTXShapeFillsUtilsType {
    getShapeFill: (node: any, pNode: any, isSvgMode: boolean, warpObj: any, source: string) => string;
    getSvgGradient: (w: number, h: number, angl: number, color_arry: any[], shpId: string) => string;
    getSvgImagePattern: (node: any, fill: any, shpId: string, warpObj: any) => string;
}

const PPTXShapeFillsUtils = {} as PPTXShapeFillsUtilsType;

    

    /**
 * Get shape fill (solid, gradient, pattern, or picture)
 * @param {Object} node - The node containing fill info
 * @param {Object} pNode - The parent node
 * @param {Boolean} isSvgMode - Whether to return SVG format
 * @param {Object} warpObj - The warp object containing theme content
 * @param {String} source - The source ("slideLayoutBg", "slideMasterBg", etc.)
 * @returns {String} Fill CSS or SVG format
 */
PPTXShapeFillsUtils.getShapeFill = function(node: any, pNode: any, isSvgMode: boolean, warpObj: any, source: string): string {
    const fillType: string = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    let fillColor: any;

    if (fillType == "NO_FILL") {
        return isSvgMode ? "none" : "";
    } else if (fillType == "SOLID_FILL") {
        const shpFill: any = node["p:spPr"]["a:solidFill"];
        fillColor = PPTXColorUtils.getSolidFill(shpFill, undefined, undefined, warpObj);
    } else if (fillType == "GRADIENT_FILL") {
        const shpFill: any = node["p:spPr"]["a:gradFill"];
        fillColor = PPTXColorUtils.getGradientFill(shpFill, warpObj);
    } else if (fillType == "PATTERN_FILL") {
        const shpFill: any = node["p:spPr"]["a:pattFill"];
        fillColor = PPTXColorUtils.getPatternFill(shpFill, warpObj);
    } else if (fillType == "PIC_FILL") {
        const shpFill: any = node["p:spPr"]["a:blipFill"];
        fillColor = PPTXColorUtils.getPicFill(source, shpFill, warpObj);
    }

    // drawingML namespace
    if (fillColor === undefined) {
        const clrName: any = PPTXUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
        let idx: number = parseInt(PPTXUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
        if (idx == 0 || idx == 1000) {
            return isSvgMode ? "none" : "";
        } else if (idx > 0 && idx < 1000) {
            // <a:fillStyleLst> fill
        } else if (idx > 1000) {
            //<a:bgFillStyleLst>
        }
        fillColor = PPTXColorUtils.getSolidFill(clrName, undefined, undefined, warpObj);
    }

    // is group fill
    if (fillColor === undefined) {
        const grpFill: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
        if (grpFill !== undefined) {
            const grpShpFill: any = pNode["p:grpSpPr"];
            const spShpNode: any = { "p:spPr": grpShpFill };
            return PPTXShapeFillsUtils.getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
        } else if (fillType == "NO_FILL") {
            return isSvgMode ? "none" : "";
        }
    }

    if (fillColor !== undefined) {
        if (fillType == "GRADIENT_FILL") {
            if (isSvgMode) {
                return fillColor;
            } else {
                const colorAry: any[] = fillColor.color;
                const rot: number = fillColor.rot;
                let bgcolor: string = "background: linear-gradient(" + rot + "deg,";
                for (let i: number = 0; i < colorAry.length; i++) {
                    if (i == colorAry.length - 1) {
                        bgcolor += "#" + colorAry[i] + ");";
                    } else {
                        bgcolor += "#" + colorAry[i] + ", ";
                    }
                }
                return bgcolor;
            }
        } else if (fillType == "PIC_FILL") {
            if (isSvgMode) {
                return fillColor;
            } else {
                return "background-image:url(" + fillColor + ");";
            }
        } else if (fillType == "PATTERN_FILL") {
            let bgPtrn: string = "", bgSize: string = "", bgPos: string = "";
            bgPtrn = fillColor[0];
            if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                bgSize = " background-size:" + fillColor[1] + ";";
            }
            if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                bgPos = " background-position:" + fillColor[2] + ";";
            }
            return "background: " + bgPtrn + ";" + bgSize + bgPos;
        } else {
            if (isSvgMode) {
                const color: any = tinycolor(fillColor);
                fillColor = color.toRgbString();
                return fillColor;
            } else {
                return "background-color: #" + fillColor + ";";
            }
        }
    } else {
        if (isSvgMode) {
            return "none";
        } else {
            return "background-color: transparent;";
        }
    }
};

    /**
 * Get SVG gradient definition
 * @param {Number} w - Width
 * @param {Number} h - Height
 * @param {Number} angl - Angle
 * @param {Array} color_arry - Array of color hex values
 * @param {String} shpId - Shape ID
 * @returns {String} SVG gradient XML
 */
PPTXShapeFillsUtils.getSvgGradient = function(w: number, h: number, angl: number, color_arry: any[], shpId: string): string {
    const stopsArray: any[] = PPTXColorUtils.getMiddleStops(color_arry.length - 2);

    const svgHeight: number = h,
        svgWidth: number = w;
    let svgAngle: string = '',
        svg: string = '';
    const xy_ary: number[] = PPTXColorUtils.SVGangle(angl, svgHeight, svgWidth),
        x1: number = xy_ary[0],
        y1: number = xy_ary[1],
        x2: number = xy_ary[2],
        y2: number = xy_ary[3];

    const sal: number = stopsArray.length,
        sr: number = sal < 20 ? 100 : 1000;
    svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
    svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
    svg += svgAngle;

    for (let i: number = 0; i < sal; i++) {
        const tinClr: any = tinycolor("#" + color_arry[i]);
        let alpha: number = tinClr.getAlpha();
        svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
        svg += '/>\n';
    }

    svg += '</linearGradient>\n';

    return svg;
};

    /**
 * Get SVG image pattern definition
 * @param {Object} node - The node containing image info
 * @param {String} fill - Fill value
 * @param {String} shpId - Shape ID
 * @param {Object} warpObj - The warp object
 * @returns {String} SVG pattern XML
 */
PPTXShapeFillsUtils.getSvgImagePattern = function(node: any, fill: any, shpId: string, warpObj: any): string {
    // 处理 fill 可能是对象的情况（当 getPicFill 返回包含属性的对象时）
    const fillValue: any = typeof fill === 'object' && fill.img ? fill.img : fill;
    // 优先使用 imgData 获取图片尺寸（base64格式）
    let dimSrc: any = null;
    if (typeof fill === 'object' && fill.imgData) {
        dimSrc = fill.imgData;
    } else if (fillValue && typeof fillValue === 'string' && fillValue.indexOf("data:image/") === 0) {
        // fillValue 是 data: URI，可以直接用于尺寸获取
        dimSrc = fillValue;
    }
    let pic_dim: number[] = [0, 0];
    if (dimSrc) {
        pic_dim = PPTXColorUtils.getBase64ImageDimensions(dimSrc);
    }
    let width: number = pic_dim[0];
    let height: number = pic_dim[1];

    const blipFillNode: any = node["p:spPr"]["a:blipFill"];
    const tileNode: any = PPTXUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"]);
    let sx: number | undefined, sy: number | undefined;
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        sx = (parseInt(tileNode["sx"]) / 100000) * width;
        sy = (parseInt(tileNode["sy"]) / 100000) * height;
    }

    const blipNode: any = node["p:spPr"]["a:blipFill"]["a:blip"];
    const tialphaModFixNode: any = PPTXUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"]);
    let imgOpacity: string = "";
    if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
        const amt: number = parseInt(tialphaModFixNode["amt"]) / 100000;
        const opacity: number = amt;
        imgOpacity = "opacity='" + opacity + "'";
    }

    let ptrn: string;
    if (sx !== undefined && sx != 0) {
        ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
    } else {
        ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
    }

    const duotoneNode: any = PPTXUtils.getTextByPathList(blipNode, ["a:duotone"]);
    let fillterNode: string = "";
    let filterUrl: string = "";

    if (duotoneNode !== undefined) {
        const clr_ary: any[] = [];
        Object.keys(duotoneNode).forEach(function (clr_type: string) {
            if (clr_type != "attrs") {
                let obj: any = {};
                obj[clr_type] = duotoneNode[clr_type];
                const hexClr: string = PPTXColorUtils.getSolidFill(obj, undefined, undefined, warpObj);
                const color: any = tinycolor("#" + hexClr);
                clr_ary.push(color.toRgb());
            }
        });

        if (clr_ary.length == 2) {
            fillterNode = '<filter id="svg_image_duotone"> ' +
                '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '.33 .33 .33 0 0' +
                '0 0 0 1 0">' +
                '</feColorMatrix>' +
                '<feComponentTransfer color-interpolation-filters="sRGB">' +
                '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                '</feComponentTransfer>' +
                ' </filter>';
        }

        filterUrl = 'filter="url(#svg_image_duotone)"';
        ptrn += fillterNode;
    }

    // Check if fill already contains blob: or data: URI prefix
    const imgSrc: string = (fillValue && typeof fillValue === 'string' && (fillValue.indexOf("blob:") === 0 || fillValue.indexOf("data:") === 0)) ? fillValue : "data:image/png;base64," + (typeof fillValue === 'string' ? fillValue : '');
    ptrn += '<image x="0" y="0" width="' + width + '" height="' + height + '" xlink:href="' + imgSrc + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
    ptrn += '</pattern>';

    return ptrn;
};

export { PPTXShapeFillsUtils };