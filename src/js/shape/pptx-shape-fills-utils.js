
import { PPTXUtils } from '../utils/utils.js';
import { PPTXColorUtils } from '../core/pptx-color-utils.js';

var PPTXShapeFillsUtils = {};

    

    /**
 * Get shape fill (solid, gradient, pattern, or picture)
 * @param {Object} node - The node containing fill info
 * @param {Object} pNode - The parent node
 * @param {Boolean} isSvgMode - Whether to return SVG format
 * @param {Object} warpObj - The warp object containing theme content
 * @param {String} source - The source ("slideLayoutBg", "slideMasterBg", etc.)
 * @returns {String} Fill CSS or SVG format
 */
PPTXShapeFillsUtils.getShapeFill = function(node, pNode, isSvgMode, warpObj, source) {
    var fillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    var fillColor;

    if (fillType == "NO_FILL") {
        return isSvgMode ? "none" : "";
    } else if (fillType == "SOLID_FILL") {
        var shpFill = node["p:spPr"]["a:solidFill"];
        fillColor = PPTXColorUtils.getSolidFill(shpFill, undefined, undefined, warpObj);
    } else if (fillType == "GRADIENT_FILL") {
        var shpFill = node["p:spPr"]["a:gradFill"];
        fillColor = PPTXColorUtils.getGradientFill(shpFill, warpObj);
    } else if (fillType == "PATTERN_FILL") {
        var shpFill = node["p:spPr"]["a:pattFill"];
        fillColor = PPTXColorUtils.getPatternFill(shpFill, warpObj);
    } else if (fillType == "PIC_FILL") {
        var shpFill = node["p:spPr"]["a:blipFill"];
        fillColor = PPTXColorUtils.getPicFill(source, shpFill, warpObj);
    }

    // drawingML namespace
    if (fillColor === undefined) {
        var clrName = PPTXUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
        var idx = parseInt(PPTXUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
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
        var grpFill = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
        if (grpFill !== undefined) {
            var grpShpFill = pNode["p:grpSpPr"];
            var spShpNode = { "p:spPr": grpShpFill };
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
                var colorAry = fillColor.color;
                var rot = fillColor.rot;
                var bgcolor = "background: linear-gradient(" + rot + "deg,";
                for (var i = 0; i < colorAry.length; i++) {
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
            var bgPtrn = "", bgSize = "", bgPos = "";
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
                var color = tinycolor(fillColor);
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
PPTXShapeFillsUtils.getSvgGradient = function(w, h, angl, color_arry, shpId) {
    var stopsArray = PPTXColorUtils.getMiddleStops(color_arry - 2);

    var svgAngle = '',
        svgHeight = h,
        svgWidth = w,
        svg = '',
        xy_ary = PPTXColorUtils.SVGangle(angl, svgHeight, svgWidth),
        x1 = xy_ary[0],
        y1 = xy_ary[1],
        x2 = xy_ary[2],
        y2 = xy_ary[3];

    var sal = stopsArray.length,
        sr = sal < 20 ? 100 : 1000;
    svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
    svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
    svg += svgAngle;

    for (var i = 0; i < sal; i++) {
        var tinClr = tinycolor("#" + color_arry[i]);
        var alpha = tinClr.getAlpha();
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
PPTXShapeFillsUtils.getSvgImagePattern = function(node, fill, shpId, warpObj) {
    // 处理 fill 可能是对象的情况（当 getPicFill 返回包含属性的对象时）
    var fillValue = typeof fill === 'object' && fill.img ? fill.img : fill;
    var pic_dim = PPTXColorUtils.getBase64ImageDimensions(fillValue);
    var width = pic_dim[0];
    var height = pic_dim[1];

    var blipFillNode = node["p:spPr"]["a:blipFill"];
    var tileNode = PPTXUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"]);
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        var sx = (parseInt(tileNode["sx"]) / 100000) * width;
        var sy = (parseInt(tileNode["sy"]) / 100000) * height;
    }

    var blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
    var tialphaModFixNode = PPTXUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"]);
    var imgOpacity = "";
    if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
        var amt = parseInt(tialphaModFixNode["amt"]) / 100000;
        var opacity = amt;
        var imgOpacity = "opacity='" + opacity + "'";
    }

    if (sx !== undefined && sx != 0) {
        var ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
    } else {
        var ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
    }

    var duotoneNode = PPTXUtils.getTextByPathList(blipNode, ["a:duotone"]);
    var fillterNode = "";
    var filterUrl = "";

    if (duotoneNode !== undefined) {
        var clr_ary = [];
        Object.keys(duotoneNode).forEach(function (clr_type) {
            if (clr_type != "attrs") {
                var obj = {};
                obj[clr_type] = duotoneNode[clr_type];
                var hexClr = PPTXColorUtils.getSolidFill(obj, undefined, undefined, warpObj);
                var color = tinycolor("#" + hexClr);
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
    var imgSrc = (fillValue && (fillValue.indexOf("blob:") === 0 || fillValue.indexOf("data:") === 0)) ? fillValue : "data:image/png;base64," + fillValue;
    ptrn += '<image x="0" y="0" width="' + width + '" height="' + height + '" xlink:href="' + imgSrc + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
    ptrn += '</pattern>';

    return ptrn;
};

    // Export to window

export { PPTXShapeFillsUtils };

// Also export to global scope for backward compatibility
// window.PPTXShapeFillsUtils = PPTXShapeFillsUtils; // Removed for ES modules
