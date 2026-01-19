// @ts-nocheck
/**
 * pptx-shape-generator.js
 * 形状生成器模块
 *
 * 这个模块包含了 genShape 函数，用于生成形状的SVG HTML表示
 * 由于 genShape 函数非常庞大（超过5000行），它处理所有形状类型的生成逻辑
 */

import { PPTXUtils } from '../core/utils'
import { PPTXConstants } from '../core/constants'
import { PPTXShapePropertyExtractor } from './property-extractor';
import { PPTXShapeFillsUtils } from './fills';
import { PPTXStyleManager } from '../core/style-manager';
import { PPTXColorUtils } from '../core/color';
import { PPTXTextStyleUtils } from '../text/style';
import { PPTXTextElementUtils } from '../text/element.js';
import { PPTXBasicShapes } from './basic.js';
import { PPTXStarShapes } from './star.js';
import { PPTXFlowchartShapes } from './flowchart.js';
import { PPTXActionButtonShapes } from './actionbutton.js';
import { PPTXArrowShapes } from './arrow.js';
import { PPTXCalloutShapes } from './callout.js';
import { PPTXShapeContainer } from './container.js';
import { PPTXShapeUtils } from './shape.js';
import { PPTXMathShapes } from './math.js';

    const slideFactor = PPTXConstants.SLIDE_FACTOR;

/**
         * 生成形状的SVG HTML表示
         * 
         * @param {Object} node - 形状节点对象
         * @param {Object} pNode - 父节点对象
         * @param {Object} slideLayoutSpNode - 幻灯片布局中的形状节点
         * @param {Object} slideMasterSpNode - 幻灯片母版中的形状节点  
         * @param {string} id - 形状ID
         * @param {string} name - 形状名称
         * @param {string} idx - 形状索引
         * @param {string} type - 形状类型
         * @param {number} order - 显示顺序
         * @param {Object} warpObj - 包装对象，包含解析上下文
         * @param {boolean} isUserDrawnBg - 是否用户绘制的背景
         * @param {string} sType - 形状类型标识
         * @param {string} source - 来源标识
         * @returns {string} 形状的SVG HTML字符串
         */
        function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source, styleTable={}) {
            //const dltX = 0;
            //const dltY = 0;
            let result = "";
            let dVal;
            // 使用属性提取器获取形状属性
            const props = PPTXShapePropertyExtractor.extractShapeProperties(node, slideFactor, pNode, slideLayoutSpNode, slideMasterSpNode);
            const slideXfrmNode = props.slideXfrmNode;
            const shapType = props.shapType;
            const custShapType = props.custShapType;
            const rotate = props.rotate;
            const flip = props.flip;
            const txtRotate = props.txtRotate;
            const shpId = props.shpId;
            const w = props.w;
            const h = props.h;
            let x = props.x;
            let y = props.y;
            const slideLayoutXfrmNode = props.slideLayoutXfrmNode;
            const slideMasterXfrmNode = props.slideMasterXfrmNode;

            let grndFillFlg = false;
            let imgFillFlg = false;
            let fillColor;
            let border;
            let oShadowSvgUrlStr;
            let headEndNodeAttrs;
            let tailEndNodeAttrs;

            if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {

                const svgCssName = "_svg_css_" + (Object.keys(styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
                //console.log("name:", name, "svgCssName: ", svgCssName)
                const effectsClassName = svgCssName + "_effects";
                     result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + "' style='" +
                         PPTXUtils.getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
                          PPTXUtils.getSize(slideXfrmNode, undefined, undefined) +
                          " z-index: " + order + ";transform: rotate(" + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + "'>";
                result += '<defs>'
                // Fill Color
                fillColor = PPTXShapeFillsUtils.getShapeFill(node, pNode, true, warpObj, source);
                //console.log("genShape: fillColor: ", fillColor)
                grndFillFlg = false;
                imgFillFlg = false;
                let clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
                if (clrFillType == "GROUP_FILL") {
                    clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
                }
                // if (clrFillType == "") {
                //     const clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:style","a:fillRef"]));
                // }
                //console.log("genShape: fillColor: ", fillColor, ", clrFillType: ", clrFillType, ", node: ", node)
                /////////////////////////////////////////                    
                if (clrFillType == "GRADIENT_FILL") {
                    grndFillFlg = true;
                    const color_arry = fillColor.color;
                    const angl = fillColor.rot + 90;
                    const svgGrdnt = PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, shpId);
                    //fill="url(#linGrd)"
                    //console.log("genShape: svgGrdnt: ", svgGrdnt)
                    result += svgGrdnt;

                } else if (clrFillType == "PIC_FILL") {
                    imgFillFlg = true;
                    // 提取图片 URL（fillColor 可能是对象或字符串）
                    const imgFill = typeof fillColor === 'object' && fillColor.img ? fillColor.img : fillColor;
                    const svgBgImg = PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, shpId, warpObj);
                    //fill="url(#imgPtrn)"
                    //console.log(svgBgImg)
                    result += svgBgImg;
                } else if (clrFillType == "PATTERN_FILL") {
                    let styleText = fillColor;
                    if (styleText in styleTable) {
                        styleText += "do-nothing: " + svgCssName +";";
                    }
                    styleTable[styleText] = {
                        "name": svgCssName,
                        "text": styleText
                    };
                    //}
                    fillColor = "none";
                } else {
                    if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                        (shapType == "arc" ||
                            shapType == "bracketPair" ||
                            shapType == "bracePair" ||
                            shapType == "leftBracket" ||
                            shapType == "leftBrace" ||
                            shapType == "rightBrace" ||
                            shapType == "rightBracket")) { 
                        // 临时解决方案：对于弧形、括号等形状，当填充类型不是实心或图案时，设置为无填充
                        // 这是因为这些形状在某些情况下应该显示为轮廓，但需要更精确的处理
                        fillColor = "none";
                    }
                }
                // Border Color
                border = PPTXStyleManager.getBorder(node, pNode, true, "shape", warpObj);

                headEndNodeAttrs = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
                tailEndNodeAttrs = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
                // type: none, triangle, stealth, diamond, oval, arrow

                ////////////////////effects/////////////////////////////////////////////////////
                //p:spPr => a:effectLst =>
                //"a:blur"
                //"a:fillOverlay"
                //"a:glow"
                //"a:innerShdw"
                //"a:outerShdw"
                //"a:prstShdw"
                //"a:reflection"
                //"a:softEdge"
                //p:spPr => a:scene3d
                //"a:camera"
                //"a:lightRig"
                //"a:backdrop"
                //"a:extLst"?
                //p:spPr => a:sp3d
                //"a:bevelT"
                //"a:bevelB"
                //"a:extrusionClr"
                //"a:contourClr"
                //"a:extLst"?
                //////////////////////////////outerShdw///////////////////////////////////////////
                //not support sizing the shadow
                const outerShdwNode = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
                oShadowSvgUrlStr = ""
                if (outerShdwNode !== undefined) {
                    const chdwClrNode = PPTXColorUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj);
                    const outerShdwAttrs = outerShdwNode["attrs"];

                    //const algn = outerShdwAttrs["algn"];
                    let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                    const dist = parseInt(outerShdwAttrs["dist"]) * slideFactor;//(px) //* (3 / 4); //(pt)
                    //const rotWithShape = outerShdwAttrs["rotWithShape"];
                    const blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor) : ""; //+ "px"
                    //let sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                    //let sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                    const vx = dist * Math.sin(dir * Math.PI / 180);
                    const hx = dist * Math.cos(dir * Math.PI / 180);
                    //SVG
                    //const oShadowId = "outerhadow_" + shpId;
                    //oShadowSvgUrlStr = "filter='url(#" + oShadowId+")'";
                    //const shadowFilterStr = '<filter id="' + oShadowId + '" x="0" y="0" width="' + w * (6 / 8) + '" height="' + h + '">';
                    //1:
                    //shadowFilterStr += '<feDropShadow dx="' + vx + '" dy="' + hx + '" stdDeviation="' + blurRad * (3 / 4) + '" flood-color="#' + chdwClrNode +'" flood-opacity="1" />'
                    //2:
                    //shadowFilterStr += '<feFlood result="floodColor" flood-color="red" flood-opacity="0.5"   width="' + w * (6 / 8) + '" height="' + h + '"  />'; //#' + chdwClrNode +'
                    //shadowFilterStr += '<feOffset result="offOut" in="SourceGraph ccfsdf-+ic"  dx="' + vx + '" dy="' + hx + '"/>'; //how much to offset
                    //shadowFilterStr += '<feGaussianBlur result="blurOut" in="offOut" stdDeviation="' + blurRad*(3/4) +'"/>'; //tdDeviation is how much to blur
                    //shadowFilterStr += '<feComponentTransfer><feFuncA type="linear" slope="0.5"/></feComponentTransfer>'; //slope is the opacity of the shadow
                    //shadowFilterStr += '<feBlend in="SourceGraphic" in2="blurOut"  mode="normal" />'; //this contains the element that the filter is applied to
                    //shadowFilterStr += '</filter>'; 
                    //result += shadowFilterStr;

                    //css:
                    let svg_css_shadow = "filter:drop-shadow(" + hx + "px " + vx + "px " + blurRad + "px #" + chdwClrNode + ");";

                    if (svg_css_shadow in styleTable) {
                        svg_css_shadow += "do-nothing: " + svgCssName + ";";
                    }

                    styleTable[svg_css_shadow] = {
                        "name": effectsClassName,
                        "text": svg_css_shadow
                    };

                } 
                ////////////////////////////////////////////////////////////////////////////////////////
                if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                    (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                    const triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border.color + "' fill='" + border.color +
                        "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
                    result += triangleMarker;
                }
                result += '</defs>'
            }
            if (shapType !== undefined && custShapType === undefined) {
                //console.log("shapType: ", shapType)
                let d = "", d_val, points;
                let x1, x2, y1, y2, c3d4, cd4, cd2, wd2, hd2;
                let fillAttr, shapAdjst_ary, sAdj1, sAdj2, sAdj1_val, sAdj2_val, sAdj_name;
                let tranglRott, adjst_val, max_adj_const;
                let adj, adj1, adj2, adj3, adj4, adj5, adj6, adj7, adj8;
                let cnstVal, cnstVal1, cnstVal2, cnstVal3, cnstVal4, cnstVal5;
                let angVal, angVal1, angVal2, angVal3, angVal4;
                let a, a1, a2, a3;
                let dx, dy, dx1, dx2, dx3, dx4, dx5, dy1, dy2, dy3, dy4, dy5, dz;
                let vc, hc, idy, ib, iDx, il, ir, it;
                let ss, maxAdj, maxAdj1, maxAdj2, maxAdj3, minWH;
                let refr, H, isClose, shapAdjst1, shapAdjst2;
                let x3, x4, x5, x6, x7, y3, y4, y5, y6;
                let t, l, b, r, wd8, wd32;
                let g0, g1, g2, q1, q2, q3, q4, q5, q6, q7, q8, q9, q10, q11;
                let shd2, vf, dr, iwd2, ihd2;
                let ct, st, m, n, drd2, dang, dang2, swAng, t3, stAng, istAng, sw11, sw12, iswAng;
                let stAng1, stAng2, stAng1deg, stAng2deg, swAng2deg;
                let ct1, st1, m1, n1;
                let pieVals, hR, wR;
                let ang, ang2rad;
                let cX1, cY1, cX2, cY2, cy1, cy3;
                let bl, br, dt;
                let prcnt, dfltBultSizeNoPt, font_val;
                let offAttrs;
                let ang1, ang1Dg;
                switch (shapType) {
                    case "rect":
                    case "flowChartProcess":
                    case "flowChartPredefinedProcess":
                    case "flowChartInternalStorage":
                    case "actionButtonBlank":
                        result += PPTXBasicShapes.genRectWithDecoration(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border, oShadowSvgUrlStr, shapType);
                        break;
                    case "flowChartCollate":
                        result += PPTXFlowchartShapes.genFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "flowChartDocument":
                        result += PPTXFlowchartShapes.genFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "flowChartMultidocument":
                        result += PPTXFlowchartShapes.genFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonBackPrevious":
                        result += PPTXActionButtonShapes.genActionButtonBackPrevious(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonBeginning":
                        result += PPTXActionButtonShapes.genActionButtonBeginning(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonDocument":
                        result += PPTXActionButtonShapes.genActionButtonDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonEnd":
                        result += PPTXActionButtonShapes.genActionButtonEnd(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonForwardNext":
                        result += PPTXActionButtonShapes.genActionButtonForwardNext(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonHelp":
                        result += PPTXActionButtonShapes.genActionButtonHelp(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonHome":
                        result += PPTXActionButtonShapes.genActionButtonHome(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonInformation":
                        result += PPTXActionButtonShapes.genActionButtonInformation(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonMovie":
                        result += PPTXActionButtonShapes.genActionButtonMovie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonReturn":
                        result += PPTXActionButtonShapes.genActionButtonReturn(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "actionButtonSound":
                        result += PPTXActionButtonShapes.genActionButtonSound(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        break;
                    case "irregularSeal1":
                    case "irregularSeal2":
                        if (shapType == "irregularSeal1") {
                            d = "M" + w * 10800 / 21600 + "," + h * 5800 / 21600 +
                                " L" + w * 14522 / 21600 + "," + 0 +
                                " L" + w * 14155 / 21600 + "," + h * 5325 / 21600 +
                                " L" + w * 18380 / 21600 + "," + h * 4457 / 21600 +
                                " L" + w * 16702 / 21600 + "," + h * 7315 / 21600 +
                                " L" + w * 21097 / 21600 + "," + h * 8137 / 21600 +
                                " L" + w * 17607 / 21600 + "," + h * 10475 / 21600 +
                                " L" + w + "," + h * 13290 / 21600 +
                                " L" + w * 16837 / 21600 + "," + h * 12942 / 21600 +
                                " L" + w * 18145 / 21600 + "," + h * 18095 / 21600 +
                                " L" + w * 14020 / 21600 + "," + h * 14457 / 21600 +
                                " L" + w * 13247 / 21600 + "," + h * 19737 / 21600 +
                                " L" + w * 10532 / 21600 + "," + h * 14935 / 21600 +
                                " L" + w * 8485 / 21600 + "," + h +
                                " L" + w * 7715 / 21600 + "," + h * 15627 / 21600 +
                                " L" + w * 4762 / 21600 + "," + h * 17617 / 21600 +
                                " L" + w * 5667 / 21600 + "," + h * 13937 / 21600 +
                                " L" + w * 135 / 21600 + "," + h * 14587 / 21600 +
                                " L" + w * 3722 / 21600 + "," + h * 11775 / 21600 +
                                " L" + 0 + "," + h * 8615 / 21600 +
                                " L" + w * 4627 / 21600 + "," + h * 7617 / 21600 +
                                " L" + w * 370 / 21600 + "," + h * 2295 / 21600 +
                                " L" + w * 7312 / 21600 + "," + h * 6320 / 21600 +
                                " L" + w * 8352 / 21600 + "," + h * 2295 / 21600 +
                                " z";
                        } else if (shapType == "irregularSeal2") {
                            d = "M" + w * 11462 / 21600 + "," + h * 4342 / 21600 +
                                " L" + w * 14790 / 21600 + "," + 0 +
                                " L" + w * 14525 / 21600 + "," + h * 5777 / 21600 +
                                " L" + w * 18007 / 21600 + "," + h * 3172 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 6532 / 21600 +
                                " L" + w + "," + h * 6645 / 21600 +
                                " L" + w * 16985 / 21600 + "," + h * 9402 / 21600 +
                                " L" + w * 18270 / 21600 + "," + h * 11290 / 21600 +
                                " L" + w * 16380 / 21600 + "," + h * 12310 / 21600 +
                                " L" + w * 18877 / 21600 + "," + h * 15632 / 21600 +
                                " L" + w * 14640 / 21600 + "," + h * 14350 / 21600 +
                                " L" + w * 14942 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 12180 / 21600 + "," + h * 15935 / 21600 +
                                " L" + w * 11612 / 21600 + "," + h * 18842 / 21600 +
                                " L" + w * 9872 / 21600 + "," + h * 17370 / 21600 +
                                " L" + w * 8700 / 21600 + "," + h * 19712 / 21600 +
                                " L" + w * 7527 / 21600 + "," + h * 18125 / 21600 +
                                " L" + w * 4917 / 21600 + "," + h +
                                " L" + w * 4805 / 21600 + "," + h * 18240 / 21600 +
                                " L" + w * 1285 / 21600 + "," + h * 17825 / 21600 +
                                " L" + w * 3330 / 21600 + "," + h * 15370 / 21600 +
                                " L" + 0 + "," + h * 12877 / 21600 +
                                " L" + w * 3935 / 21600 + "," + h * 11592 / 21600 +
                                " L" + w * 1172 / 21600 + "," + h * 8270 / 21600 +
                                " L" + w * 5372 / 21600 + "," + h * 7817 / 21600 +
                                " L" + w * 4502 / 21600 + "," + h * 3625 / 21600 +
                                " L" + w * 8550 / 21600 + "," + h * 6382 / 21600 +
                                " L" + w * 9722 / 21600 + "," + h * 1887 / 21600 +
                                " z";
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartTerminator":
                        cd2 = 180, cd4 = 90, c3d4 = 270;
                        x1 = w * 3475 / 21600;
                        x2 = w * 18125 / 21600;
                        y1 = h * 10800 / 21600;
                        //path attrs: w = 21600; h = 21600; 
                        d = "M" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartPunchedTape":
                        cd2 = 180;
                        x1 = w * 5 / 20;
                        y1 = h * 2 / 20;
                        y2 = h * 18 / 20;
                        d = "M" + 0 + "," + y1 +
                            PPTXShapeUtils.shapeArc(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
                            " L" + w + "," + y2 +
                            PPTXShapeUtils.shapeArc(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartOnlineStorage":
                        c3d4 = 270, cd4 = 90;
                        x1 = w * 1 / 6;
                        y1 = h * 3 / 6;
                        d = "M" + x1 + "," + 0 +
                            " L" + w + "," + 0 +
                            PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartDisplay":
                        c3d4 = 270, cd2 = 180;
                        x1 = w * 1 / 6;
                        x2 = w * 5 / 6;
                        y1 = h * 3 / 6;
                        //path attrs: w = 6; h = 6; 
                        d = "M" + 0 + "," + y1 +
                            " L" + x1 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartDelay":
                        wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        d = "M" + 0 + "," + 0 +
                            " L" + wd2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L") +
                            " L" + 0 + "," + h +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "flowChartMagneticTape":
                        wd2 = w / 2, hd2 = h / 2, cd2 = 180, c3d4 = 270, cd4 = 90;
                        idy = hd2 * Math.sin(Math.PI / 4);
                        ib = hd2 + idy;
                        ang1 = Math.atan(h / w);
                        ang1Dg = ang1 * 180 / Math.PI;
                        d = "M" + wd2 + "," + h +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L") +
                            " L" + w + "," + ib +
                            " L" + w + "," + h +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "ellipse":
                    case "flowChartConnector":
                    case "flowChartSummingJunction":
                    case "flowChartOr":
                        result += PPTXBasicShapes.genEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor, border);
                        if (shapType == "flowChartOr") {
                            result += " <polyline points='" + w / 2 + " " + 0 + "," + w / 2 + " " + h + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + 0 + " " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        } else if (shapType == "flowChartSummingJunction") {
                            hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
                            const angVal = Math.PI / 4;
                            iDx = wd2 * Math.cos(angVal);
                            idy = hd2 * Math.sin(angVal);
                            il = hc - iDx;
                            ir = hc + iDx;
                            it = vc - idy;
                            ib = vc + idy;
                            result += " <polyline points='" + il + " " + it + "," + ir + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                            result += " <polyline points='" + ir + " " + it + "," + il + " " + ib + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    case "roundRect":
                    case "round1Rect":
                    case "round2DiagRect":
                    case "round2SameRect":
                    case "snip1Rect":
                    case "snip2DiagRect":
                    case "snip2SameRect":
                    case "flowChartAlternateProcess":
                    case "flowChartPunchedCard":
                        let shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, sAdj1_val;// = 0.33334;
                        let sAdj2, sAdj2_val;// = 0.33334;
                        let shpTyp, adjTyp;
                        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
                                const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
                            const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                            sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
                            sAdj2_val = 0;
                        }
                        //console.log("shapType: ",shapType,",node: ",node )
                        tranglRott = "";
                        switch (shapType) {
                            case "roundRect":
                            case "flowChartAlternateProcess":
                                shpTyp = "round";
                                adjTyp = "cornrAll";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            case "round1Rect":
                                shpTyp = "round";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            case "round2DiagRect":
                                shpTyp = "round";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            case "round2SameRect":
                                shpTyp = "round";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            case "snip1Rect":
                            case "flowChartPunchedCard":
                                shpTyp = "snip";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                if (shapType == "flowChartPunchedCard") {
                                    tranglRott = "transform='translate(" + w + ",0) scale(-1,1)'";
                                }
                                break;
                            case "snip2DiagRect":
                                shpTyp = "snip";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0;
                                if (sAdj2_val === undefined) sAdj2_val = 0.33334;
                                break;
                            case "snip2SameRect":
                                shpTyp = "snip";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                        }
                        let d_val: any = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path " + tranglRott + "  d='" + d_val + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "snipRoundRect":
                        shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        sAdj1_val = 0.33334;
                        sAdj2 = undefined;
                        sAdj2_val = 0.33334;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        }
                        d_val = "M0," + h + " L" + w + "," + h + " L" + w + "," + (h / 2) * sAdj2_val +
                            " L" + (w / 2 + (w / 2) * (1 - sAdj2_val)) + ",0 L" + (w / 2) * sAdj1_val + ",0 Q0,0 0," + (h / 2) * sAdj1_val + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d_val + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bentConnector2":
                        // if (isFlipV) {
                        //     d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
                        // } else {
                        d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
                        //}
                        result += "<path d='" + d + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    case "rtTriangle":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='0 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "triangle":
                    case "flowChartExtract":
                    case "flowChartMerge":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let shapAdjst_val: any = 0.5;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) * slideFactor;
                            //console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                        }
                        tranglRott = "";
                        if (shapType == "flowChartMerge") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon " + tranglRott + " points='" + (w * shapAdjst_val) + " 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "diamond":
                    case "flowChartDecision":
                    case "flowChartSort":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        if (shapType == "flowChartSort") {
                            result += " <polyline points='0 " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        }
                        break;
                    case "trapezoid":
                    case "flowProc":
                    case "flowChartManualOperation":
                    case "flowChartManualInput":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adjst_val = 0.2;
                        max_adj_const = 0.7407;
                        if (shapAdjst !== undefined) {
                            const adjst = parseInt(shapAdjst.substr(4)) * slideFactor;
                            adjst_val = (adjst * 0.5) / max_adj_const;
                            // console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nadjst_val: "+adjst_val);
                        }
                        let cnstVal = 0;
                        tranglRott = "";
                        if (shapType == "flowChartManualOperation") {
                            tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
                        }
                        if (shapType == "flowChartManualInput") {
                            adjst_val = 0;
                            cnstVal = h / 5;
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon " + tranglRott + " points='" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "parallelogram":
                    case "flowChartInputOutput":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adjst_val = 0.25;
                        max_adj_const = undefined;
                        if (w > h) {
                            max_adj_const = w / h;
                        } else {
                            max_adj_const = h / w;
                        }
                        if (shapAdjst !== undefined) {
                            const adjst = parseInt(shapAdjst.substr(4)) / 100000;
                            adjst_val = adjst / max_adj_const;
                            //console.log("w: "+w+"\nh: "+h+"\nadjst: "+adjst_val+"\nmax_adj_const: "+max_adj_const);
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                        break;
                    case "pentagon":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + (0.5 * w) + " 0,0 " + (0.375 * h) + "," + (0.15 * w) + " " + h + "," + 0.85 * w + " " + h + "," + w + " " + 0.375 * h + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "hexagon":
                    case "flowChartPreparation":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj = 25000 * slideFactor;
                        const vf: any = 115470 * slideFactor;
                        cnstVal1 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const angVal1: any = 60 * Math.PI / 180;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        vc = h / 2;
                        hd2 = h / 2;
                        const ss: any = Math.min(w, h);
                        const maxAdj: any = cnstVal1 * w / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        const shd2: any = hd2 * vf / cnstVal2;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        dy1 = shd2 * Math.sin(angVal1);
                        y1 = vc - dy1;
                        y2 = vc + dy1;

                        d = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "heptagon":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + (0.5 * w) + " 0," + w / 8 + " " + h / 4 + ",0 " + (5 / 8) * h + "," + w / 4 + " " + h + "," + (3 / 4) * w + " " + h + "," +
                            w + " " + (5 / 8) * h + "," + (7 / 8) * w + " " + h / 4 + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "octagon":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        adj2 = (1 - adj1);
                        //console.log("adj1: "+adj1+"\nadj2: "+adj2);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + adj1 * w + " 0,0 " + adj1 * h + ",0 " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," +
                            w + " " + adj2 * h + "," + w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "decagon":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + h / 2 + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + h / 2 + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "dodecagon":
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + (3 / 8) * h + ",0 " + (5 / 8) * h + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
                            (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + (5 / 8) * h + "," + w + " " + (3 / 8) * h + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star4":
                        d = PPTXStarShapes.genStar4(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star5":
                        d = PPTXStarShapes.genStar5(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star6":
                        d = PPTXStarShapes.genStar6(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star7":
                        d = PPTXStarShapes.genStar7(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star8":
                        d = PPTXStarShapes.genStar8(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                    case "star10":
                        d = PPTXStarShapes.genStar10(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star12":
                        d = PPTXStarShapes.genStar12(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star16":
                        d = PPTXStarShapes.genStar16(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star24":
                        d = PPTXStarShapes.genStar24(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "star32":
                        d = PPTXStarShapes.genStar32(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;

                    case "pie":
                    case "pieWedge":
                    case "arc":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        adj1 = undefined;
                        adj2 = undefined;
                        let H, shapAdjst1, shapAdjst2, isClose;
                        if (shapType == "pie") {
                            adj1 = 0;
                            adj2 = 270;
                            H = h;
                            isClose = true;
                        } else if (shapType == "pieWedge") {
                            adj1 = 180;
                            adj2 = 270;
                            H = 2 * h;
                            isClose = true;
                        } else if (shapType == "arc") {
                            adj1 = 270;
                            adj2 = 0;
                            H = h;
                            isClose = false;
                        }
                        if (shapAdjst !== undefined) {
                            shapAdjst1 = PPTXUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
                            shapAdjst2 = shapAdjst1;
                            if (shapAdjst1 === undefined) {
                                shapAdjst1 = shapAdjst[0]["attrs"]["fmla"];
                                shapAdjst2 = shapAdjst[1]["attrs"]["fmla"];
                            }
                            if (shapAdjst1 !== undefined) {
                                adj1 = parseInt(shapAdjst1.substr(4)) / 60000;
                            }
                            if (shapAdjst2 !== undefined) {
                                adj2 = parseInt(shapAdjst2.substr(4)) / 60000;
                            }
                        }
                        const pieVals: any = PPTXShapeUtils.shapePie(H, w, adj1, adj2, isClose);
                        //console.log("shapType: ",shapType,"\nimgFillFlg: ",imgFillFlg,"\ngrndFillFlg: ",grndFillFlg,"\nshpId: ",shpId,"\nfillColor: ",fillColor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + pieVals[0] + "' transform='" + pieVals[1] + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "chord":
                        shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        sAdj1_val = 45;
                        sAdj2 = undefined;
                        sAdj2_val = 270;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 60000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 60000;
                                }
                            }
                        }
                        const hR: any = h / 2;
                        const wR: any = w / 2;
                        d_val = PPTXShapeUtils.shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
                        //console.log("shapType: ",shapType,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val)
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "frame":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj1 = 12500 * slideFactor;
                        cnstVal1 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        a1 = undefined, x1;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        x1 = Math.min(w, h) * a1 / cnstVal2;
                        x4 = w - x1;
                        y4 = h - x1;
                        d = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            ` zM` + x1 + "," + x1 +
                            " L" + x1 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + x1 +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "donut":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj = 25000 * slideFactor;
                        cnstVal1 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        a = undefined;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        d = "M" + 0 + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            ` zM` + dr + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "noSmoking":
                        shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj = 18750 * slideFactor;
                        cnstVal1 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        a = undefined;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        ang = Math.atan(h / w);
                        //ang2rad = ang*Math.PI/180;
                        ct = ihd2 * Math.cos(ang);
                        st = iwd2 * Math.sin(ang);
                        m = Math.sqrt(ct * ct + st * st); //"mod ct st 0"
                        n = iwd2 * ihd2 / m;
                        drd2 = dr / 2;
                        dang = Math.atan(drd2 / n);
                        dang2 = dang * 2;
                        swAng = -Math.PI + dang2;
                        //t3 = Math.atan(h/w);
                        stAng1 = ang - dang;
                        stAng2 = stAng1 - Math.PI;
                        ct1 = ihd2 * Math.cos(stAng1);
                        st1 = iwd2 * Math.sin(stAng1);
                        m1 = Math.sqrt(ct1 * ct1 + st1 * st1); //"mod ct1 st1 0"
                        n1 = iwd2 * ihd2 / m1;
                        dx1 = n1 * Math.cos(stAng1);
                        dy1 = n1 * Math.sin(stAng1);
                        x1 = w / 2 + dx1;
                        y1 = h / 2 + dy1;
                        x2 = w / 2 - dx1;
                        y2 = h / 2 - dy1;
                        ct1 = ihd2 * Math.cos(stAng1);
                        st1 = iwd2 * Math.sin(stAng1);
                        m1 = Math.sqrt(ct1 * ct1 + st1 * st1); //"mod ct1 st1 0"
                        n1 = iwd2 * ihd2 / m1;
                        dx1 = n1 * Math.cos(stAng1);
                        dy1 = n1 * Math.sin(stAng1);
                        x1 = w / 2 + dx1;
                        y1 = h / 2 + dy1;
                        x2 = w / 2 - dx1;
                        y2 = h / 2 - dy1;
                        stAng1deg = stAng1 * 180 / Math.PI;
                        stAng2deg = stAng2 * 180 / Math.PI;
                        swAng2deg = swAng * 180 / Math.PI;
                        d = "M" + 0 + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
                            ` zM` + x1 + "," + y1 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
                            ` zM` + x2 + "," + y2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
                            " z";
                        //console.log("adj: ",adj,"x1:",x1,",y1:",y1," x2:",x2,",y2:",y2,",stAng1:",stAng1,",stAng1deg:",stAng1deg,",stAng2:",stAng2,",stAng2deg:",stAng2deg,",swAng:",swAng,",swAng2deg:",swAng2deg)

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "halfFrame":
                        shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        sAdj1_val = 3.5;
                        sAdj2 = undefined;
                        sAdj2_val = 3.5;
                        const cnsVal: any = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj1_val: any = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj2_val: any = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        minWH = Math.min(w, h);
                        const maxAdj2: any = (cnsVal * w) / minWH;
                        a1 = undefined, a2;
                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        x1 = (minWH * a2) / cnsVal;
                        const g1: any = h * x1 / w;
                        const g2: any = h - g1;
                        maxAdj1 = (cnsVal * g2) / minWH;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val
                        y1 = minWH * a1 / cnsVal;
                        const dx2: any = y1 * w / h;
                        x2 = w - dx2;
                        const dy2: any = x1 * h / w;
                        y2 = h - dy2;
                        d = `M0,0 L` + w + "," + 0 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L0," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        //console.log("w: ",w,", h: ",h,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val,",maxAdj1: ",maxAdj1,",maxAdj2: ",maxAdj2)
                        break;
                    case "blockArc":
                        shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        adj1 = 180;
                        sAdj2 = undefined;
                        adj2 = 0;
                        adj3 = 25000 * slideFactor;
                        cnstVal1 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) / 60000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) / 60000;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }

                        const cd1: any = 360;
                        if (adj1 < 0) stAng = 0
                        else if (adj1 > cd1) stAng = cd1
                        else stAng = adj1 //180

                        if (adj2 < 0) istAng = 0
                        else if (adj2 > cd1) istAng = cd1
                        else istAng = adj2 //0

                        if (adj3 < 0) a3 = 0
                        else if (adj3 > cnstVal1) a3 = cnstVal1
                        else a3 = adj3

                        sw11 = istAng - stAng; // -180
                        sw12 = sw11 + cd1; //180
                        swAng = (sw11 > 0) ? sw11 : sw12; //180
                        iswAng = -swAng; //-180

                        const endAng: any = stAng + swAng;
                        const iendAng: any = istAng + iswAng;

                        const stRd: any = stAng * (Math.PI) / 180;
                        const istRd: any = istAng * (Math.PI) / 180;
                        wd2 = w / 2;
                        hd2 = h / 2;
                        hc = w / 2;
                        vc = h / 2;
                        if (stAng > 90 && stAng < 270) {
                            const wt1: any = wd2 * (Math.sin((Math.PI) / 2 - stRd));
                            const ht1: any = hd2 * (Math.cos((Math.PI) / 2 - stRd));

                            dx1 = wd2 * (Math.cos(Math.atan(ht1 / wt1)));
                            dy1 = hd2 * (Math.sin(Math.atan(ht1 / wt1)));

                            x1 = hc - dx1;
                            y1 = vc - dy1;
                        } else {
                            const wt1: any = wd2 * (Math.sin(stRd));
                            const ht1: any = hd2 * (Math.cos(stRd));

                            dx1 = wd2 * (Math.cos(Math.atan(wt1 / ht1)));
                            dy1 = hd2 * (Math.sin(Math.atan(wt1 / ht1)));

                            x1 = hc + dx1;
                            y1 = vc + dy1;
                        }
                        dr = Math.min(w, h) * a3 / cnstVal2;
                        iwd2 = wd2 - dr;
                        ihd2 = hd2 - dr;
                        //console.log("stAng: ",stAng," swAng: ",swAng ," endAng:",endAng)
                        if ((endAng <= 450 && endAng > 270) || ((endAng >= 630 && endAng < 720))) {
                            const wt2: any = iwd2 * (Math.sin(istRd));
                            const ht2: any = ihd2 * (Math.cos(istRd));
                            const dx2: any = iwd2 * (Math.cos(Math.atan(wt2 / ht2)));
                            const dy2: any = ihd2 * (Math.sin(Math.atan(wt2 / ht2)));
                            x2 = hc + dx2;
                            y2 = vc + dy2;
                        } else {
                            const wt2: any = iwd2 * (Math.sin((Math.PI) / 2 - istRd));
                            const ht2: any = ihd2 * (Math.cos((Math.PI) / 2 - istRd));

                            const dx2: any = iwd2 * (Math.cos(Math.atan(ht2 / wt2)));
                            const dy2: any = ihd2 * (Math.sin(Math.atan(ht2 / wt2)));
                            x2 = hc - dx2;
                            y2 = vc - dy2;
                        }
                        d = "M" + x1 + "," + y1 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, stAng, endAng, false).replace("M", "L") +
                            " L" + x2 + "," + y2 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, iwd2, ihd2, istAng, iendAng, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "bracePair":
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 8333 * slideFactor;
                        const cnstVal1: any = 25000 * slideFactor;
                        cnstVal2 = 50000 * slideFactor;
                        const cnstVal3: any = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        vc = h / 2;
                        let cd: number = 360, cd2: number = 180, cd4 = 90, c3d4 = 270, a, x1, x2, x3, y2, y3;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        minWH = Math.min(w, h);
                        x1 = minWH * a / cnstVal3;
                        x2 = minWH * a / cnstVal2;
                        x3 = w - x2;
                        x4 = w - x1;
                        y2 = vc - x1;
                        y3 = vc + x1;
                        y4 = h - x1;
                        //console.log("w:",w," h:",h," x1:",x1," x2:",x2," x3:",x3," x4:",x4," y2:",y2," y3:",y3," y4:",y4)
                        d = "M" + x2 + "," + h +
                            PPTXShapeUtils.shapeArc(x2, y4, x1, x1, cd4, cd2, false).replace("M", "L") +
                            " L" + x1 + "," + y3 +
                            PPTXShapeUtils.shapeArc(0, y3, x1, x1, 0, (-cd4), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(0, y2, x1, x1, cd4, 0, false).replace("M", "L") +
                            " L" + x1 + "," + x1 +
                            PPTXShapeUtils.shapeArc(x2, x1, x1, x1, cd2, c3d4, false).replace("M", "L") +
                            " M" + x3 + "," + 0 +
                            PPTXShapeUtils.shapeArc(x3, x1, x1, x1, c3d4, cd, false).replace("M", "L") +
                            " L" + x4 + "," + y2 +
                            PPTXShapeUtils.shapeArc(w, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w, y3, x1, x1, c3d4, cd2, false).replace("M", "L") +
                            " L" + x4 + "," + y4 +
                            PPTXShapeUtils.shapeArc(x3, y4, x1, x1, 0, cd4, false).replace("M", "L");

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftBrace": {
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        adj1 = 8333 * slideFactor;
                        sAdj2 = undefined;
                        adj2 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
                                const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        cd2 = 180;
                        cd4 = 90;
                        c3d4 = 270;
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal2) a2 = cnstVal2
                        else a2 = adj2
                        minWH = Math.min(w, h);
                        const q1: any = cnstVal2 - a2;
                        if (q1 < a2) q2 = q1
                        else q2 = a2
                        const q3: any = q2 / 2;
                        maxAdj1 = q3 * h / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        y1 = minWH * a1 / cnstVal2;
                        y3 = h * a2 / cnstVal2;
                        y2 = y3 - y1;
                        y4 = y3 + y1;
                        //console.log("w:",w," h:",h," q1:",q1," q2:",q2," q3:",q3," y1:",y1," y3:",y3," y4:",y4," maxAdj1:",maxAdj1)
                        d = "M" + w + "," + h +
                            PPTXShapeUtils.shapeArc(w, h - y1, w / 2, y1, cd4, cd2, false).replace("M", "L") +
                            " L" + w / 2 + "," + y4 +
                            PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, (-cd4), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(0, y2, w / 2, y1, cd4, 0, false).replace("M", "L") +
                            " L" + w / 2 + "," + y1 +
                            PPTXShapeUtils.shapeArc(w, y1, w / 2, y1, cd2, c3d4, false).replace("M", "L");

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "rightBrace": {
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        sAdj1 = undefined;
                        adj1 = 8333 * slideFactor;
                        sAdj2 = undefined;
                        adj2 = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
                                const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        cd = 360;
                        cd2 = 180; cd4 = 90; c3d4 = 270;
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal2) a2 = cnstVal2
                        else a2 = adj2
                        minWH = Math.min(w, h);
                        const q1: any = cnstVal2 - a2;
                        if (q1 < a2) q2 = q1
                        else q2 = a2
                        const q3: any = q2 / 2;
                        maxAdj1 = q3 * h / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        y1 = minWH * a1 / cnstVal2;
                        y3 = h * a2 / cnstVal2;
                        y2 = y3 - y1;
                        y4 = h - y1;
                        //console.log("w:",w," h:",h," q1:",q1," q2:",q2," q3:",q3," y1:",y1," y2:",y2," y3:",y3," y4:",y4," maxAdj1:",maxAdj1)
                        d = "M" + 0 + "," + 0 +
                            PPTXShapeUtils.shapeArc(0, y1, w / 2, y1, c3d4, cd, false).replace("M", "L") +
                            " L" + w / 2 + "," + y2 +
                            PPTXShapeUtils.shapeArc(w, y2, w / 2, y1, cd2, cd4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w, y3 + y1, w / 2, y1, c3d4, cd2, false).replace("M", "L") +
                            " L" + w / 2 + "," + y4 +
                            PPTXShapeUtils.shapeArc(0, y4, w / 2, y1, 0, cd4, false).replace("M", "L");

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "bracketPair": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 16667 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let r: number = w, b: number = h, cd2: number = 180, cd4 = 90, c3d4 = 270, a, x1, x2, y2;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        x1 = Math.min(w, h) * a / cnstVal2;
                        x2 = r - x1;
                        y2 = b - x1;
                        //console.log("w:",w," h:",h," x1:",x1," x2:",x2," y2:",y2)
                        d = PPTXShapeUtils.shapeArc(x1, x1, x1, x1, c3d4, cd2, false) +
                            PPTXShapeUtils.shapeArc(x1, y2, x1, x1, cd2, cd4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(x2, x1, x1, x1, c3d4, (c3d4 + cd4), false) +
                            PPTXShapeUtils.shapeArc(x2, y2, x1, x1, 0, cd4, false).replace("M", "L");
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "leftBracket": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 8333 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const maxAdj: any = cnstVal1 * h / Math.min(w, h);
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let r: number = w, b: number = h, cd2: number = 180, cd4 = 90, c3d4 = 270, a, y1, y2;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        y1 = Math.min(w, h) * a / cnstVal2;
                        if (y1 > w) y1 = w;
                        y2 = b - y1;
                        d = "M" + r + "," + b +
                            PPTXShapeUtils.shapeArc(y1, y2, y1, y1, cd4, cd2, false).replace("M", "L") +
                            " L" + 0 + "," + y1 +
                            PPTXShapeUtils.shapeArc(y1, y1, y1, y1, cd2, c3d4, false).replace("M", "L") +
                            " L" + r + "," + 0
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "rightBracket": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 8333 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const maxAdj: any = cnstVal1 * h / Math.min(w, h);
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let cd: number = 360, cd2: number = 180, cd4 = 90, c3d4 = 270, a, y1, y2, y3;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        y1 = Math.min(w, h) * a / cnstVal2;
                        y2 = h - y1;
                        y3 = w - y1;
                        //console.log("w:",w," h:",h," y1:",y1," y2:",y2," y3:",y3)
                        d = "M" + 0 + "," + h +
                            PPTXShapeUtils.shapeArc(y3, y2, y1, y1, cd4, 0, false).replace("M", "L") +
                            //" L"+ r + "," + y2 +
                            " L" + w + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(y3, y1, y1, y1, cd, c3d4, false).replace("M", "L") +
                            " L" + 0 + "," + 0
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "moon": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 0.5;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) / 100000;//*96/914400;;
                        }
                        hd2 = h / 2;
                        cd2 = 180;
                        cd4 = 90;

                        adj2 = (1 - adj) * w;
                        d = "M" + w + "," + h +
                            PPTXShapeUtils.shapeArc(w, hd2, w, hd2, cd4, (cd4 + cd2), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (cd4 + cd2), cd4, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "corner":
                        const shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, sAdj1_val = 50000 * slideFactor;
                        let sAdj2: any = undefined, sAdj2_val = 50000 * slideFactor;
                        const cnsVal: any = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
                                const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj1_val: any = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj2_val: any = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        minWH = Math.min(w, h);
                        maxAdj1 = cnsVal * h / minWH;
                        const maxAdj2: any = cnsVal * w / minWH;
                        a1 = undefined, a2, x1, dy1, y1;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val

                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        x1 = minWH * a2 / cnsVal;
                        dy1 = minWH * a1 / cnsVal;
                        y1 = h - dy1;

                        d = `M0,0 L` + x1 + "," + 0 +
                            " L" + x1 + "," + y1 +
                            " L" + w + "," + y1 +
                            " L" + w + "," + h +
                            " L0," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "diagStripe": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let sAdj1_val: any = 50000 * slideFactor;
                        const cnsVal: any = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            sAdj1_val = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let a1 = undefined, x2, y2;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > cnsVal) a1 = cnsVal
                        else a1 = sAdj1_val
                        x2 = w * a1 / cnsVal;
                        y2 = h * a1 / cnsVal;
                        d = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + 0 + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "'  fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "gear6":
                    case "gear9": {
                        const txtRotate: any = 0;
                        const gearNum: any = shapType.substr(4);
                        if (gearNum == "6") {
                            d = PPTXShapeUtils.shapeGear(w, h / 3.5, parseInt(gearNum));
                        } else { //gearNum=="9"
                            d = PPTXShapeUtils.shapeGear(w, h / 3.5, parseInt(gearNum));
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d + "' transform='rotate(20," + (3 / 7) * h + "," + (3 / 7) * h + ")' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "bentConnector3": {
                        let shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const shapAdjst_val: any = 0.5;
                        if (shapAdjst !== undefined) {
                            const shapAdjst_val: any = parseInt(shapAdjst.substr(4)) / 100000;
                            // if (isFlipV) {
                            //     result += " <polyline points='" + w + " 0," + ((1 - shapAdjst_val) * w) + " 0," + ((1 - shapAdjst_val) * w) + " " + h + ",0 " + h + "' fill='transparent'" +
                            //         "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            // } else {
                            result += " <polyline points='0 0," + (shapAdjst_val * w) + " 0," + (shapAdjst_val * w) + " " + h + "," + w + " " + h + "' fill='transparent' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                            //}
                            if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                            }
                            if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                                result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                            }
                            result += "/>";
                        }
                        break;
                    }
                    case "plus":{
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        adj2 = (1 - adj1);
                        result += " <polygon points='" + adj1 * w + " 0," + adj1 * w + " " + adj1 * h + ",0 " + adj1 * h + ",0 " + adj2 * h + "," +
                            adj1 * w + " " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," + adj2 * w + " " + adj2 * h + "," + w + " " + adj2 * h + "," +
                            +w + " " + adj1 * h + "," + adj2 * w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    }
                    case "teardrop":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj1 = 100000 * slideFactor;
                        const cnsVal1: any = adj1;
                        const cnsVal2: any = 200000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        a1 = undefined, r2, tw, th, sw, sh, dx1, dy1, x1, y1, x2, y2, rd45;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnsVal2) a1 = cnsVal2
                        else a1 = adj1
                        const r2: any = Math.sqrt(2);
                        const tw: any = r2 * (w / 2);
                        const th: any = r2 * (h / 2);
                        const sw: any = (tw * a1) / cnsVal1;
                        const sh: any = (th * a1) / cnsVal1;
                        const rd45: any = (45 * (Math.PI) / 180);
                        const dx1: any = sw * (Math.cos(rd45));
                        dy1 = sh * (Math.cos(rd45));
                        x1 = (w / 2) + dx1;
                        y1 = (h / 2) - dy1;
                        x2 = ((w / 2) + x1) / 2;
                        y2 = ((h / 2) + y1) / 2;

                        d_val = PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false) +
                            "Q " + x2 + ",0 " + x1 + "," + y1 +
                            "Q " + w + "," + y2 + " " + w + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") + " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        // console.log("shapAdjst: ",shapAdjst,", adj1: ",adj1);
                        break;
                    case "plaque":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        adj1 = 16667 * slideFactor;
                        const cnsVal1: any = 50000 * slideFactor;
                        const cnsVal2: any = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        a1 = undefined, x1, x2, y2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnsVal1) a1 = cnsVal1
                        else a1 = adj1
                        x1 = a1 * (Math.min(w, h)) / cnsVal2;
                        x2 = w - x1;
                        y2 = h - x1;

                        d_val = "M0," + x1 +
                            PPTXShapeUtils.shapeArc(0, 0, x1, x1, 90, 0, false).replace("M", "L") +
                            " L" + x2 + "," + 0 +
                            PPTXShapeUtils.shapeArc(w, 0, x1, x1, 180, 90, false).replace("M", "L") +
                            " L" + w + "," + y2 +
                            PPTXShapeUtils.shapeArc(w, h, x1, x1, 270, 180, false).replace("M", "L") +
                            " L" + x1 + "," + h +
                            PPTXShapeUtils.shapeArc(0, h, x1, x1, 0, -90, false).replace("M", "L") + " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "sun":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        adj1 = 25000 * refr;
                        const cnstVal1: any = 12500 * refr;
                        cnstVal2 = 46875 * refr;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        let a1;
                        if (adj1 < cnstVal1) a1 = cnstVal1
                        else if (adj1 > cnstVal2) a1 = cnstVal2
                        else a1 = adj1

                        const cnstVa3: any = 50000 * refr;
                        const cnstVa4: any = 100000 * refr;
                        g0 = cnstVa3 - a1,
                            g1 = g0 * (30274 * refr) / (32768 * refr),
                            g2 = g0 * (12540 * refr) / (32768 * refr),
                            g3 = g1 + cnstVa3,
                            g4 = g2 + cnstVa3,
                            g5 = cnstVa3 - g1,
                            g6 = cnstVa3 - g2,
                            g7 = g0 * (23170 * refr) / (32768 * refr),
                            g8 = cnstVa3 + g7,
                            g9 = cnstVa3 - g7,
                            g10 = g5 * 3 / 4,
                            g11 = g6 * 3 / 4,
                            g12 = g10 + 3662 * refr,
                            g13 = g11 + 36620 * refr,
                            g14 = g11 + 12500 * refr,
                            g15 = cnstVa4 - g10,
                            g16 = cnstVa4 - g12,
                            g17 = cnstVa4 - g13,
                            g18 = cnstVa4 - g14,
                            ox1 = w * (18436 * refr) / (21600 * refr),
                            oy1 = h * (3163 * refr) / (21600 * refr),
                            ox2 = w * (3163 * refr) / (21600 * refr),
                            oy2 = h * (18436 * refr) / (21600 * refr),
                            x8 = w * g8 / cnstVa4,
                            x9 = w * g9 / cnstVa4,
                            x10 = w * g10 / cnstVa4,
                            x12 = w * g12 / cnstVa4,
                            x13 = w * g13 / cnstVa4,
                            x14 = w * g14 / cnstVa4,
                            x15 = w * g15 / cnstVa4,
                            x16 = w * g16 / cnstVa4,
                            x17 = w * g17 / cnstVa4,
                            x18 = w * g18 / cnstVa4,
                            x19 = w * a1 / cnstVa4,
                            wR = w * g0 / cnstVa4,
                            hR = h * g0 / cnstVa4,
                            y8 = h * g8 / cnstVa4,
                            y9 = h * g9 / cnstVa4,
                            y10 = h * g10 / cnstVa4,
                            y12 = h * g12 / cnstVa4,
                            y13 = h * g13 / cnstVa4,
                            y14 = h * g14 / cnstVa4,
                            y15 = h * g15 / cnstVa4;
                        const y16: number = h * g16 / cnstVa4, y17: number = h * g17 / cnstVa4, y18: number = h * g18 / cnstVa4;

                        d_val = "M" + w + "," + h / 2 +
                            " L" + x15 + "," + y18 +
                            " L" + x15 + "," + y14 +
                            `z M` + ox1 + "," + oy1 +
                            " L" + x16 + "," + y17 +
                            " L" + x13 + "," + y12 +
                            `z M` + w / 2 + "," + 0 +
                            " L" + x18 + "," + y10 +
                            " L" + x14 + "," + y10 +
                            `z M` + ox2 + "," + oy1 +
                            " L" + x17 + "," + y12 +
                            " L" + x12 + "," + y17 +
                            `z M` + 0 + "," + h / 2 +
                            " L" + x10 + "," + y14 +
                            " L" + x10 + "," + y18 +
                            `z M` + ox2 + "," + oy2 +
                            " L" + x12 + "," + y13 +
                            " L" + x17 + "," + y16 +
                            `z M` + w / 2 + "," + h +
                            " L" + x14 + "," + y15 +
                            " L" + x18 + "," + y15 +
                            `z M` + ox1 + "," + oy2 +
                            " L" + x13 + "," + y16 +
                            " L" + x16 + "," + y13 +
                            ` z M` + x19 + "," + h / 2 +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L") +
                            " z";
                        //console.log("adj1: ",adj1,d_val);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    case "heart":
                        const dx1: any = w * 49 / 48;
                        const dx2: any = w * 10 / 48;
                        x1 = w / 2 - dx1;
                        x2 = w / 2 - dx2;
                        x3 = w / 2 + dx2;
                        x4 = w / 2 + dx1;
                        y1 = -h / 3;
                        d_val = "M" + w / 2 + "," + h / 4 +
                            "C" + x3 + "," + y1 + " " + x4 + "," + h / 4 + " " + w / 2 + "," + h +
                            "C" + x1 + "," + h / 4 + " " + x2 + "," + y1 + " " + w / 2 + "," + h / 4 + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path   d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "lightningBolt":
                        x1 = w * 5022 / 21600,
                            x2 = w * 11050 / 21600,
                            x3 = w * 8472 / 21600,
                            x4 = w * 8757 / 21600,
                            x5 = w * 10012 / 21600,
                            x6 = w * 14767 / 21600,
                            x7 = w * 12222 / 21600,
                            x8 = w * 12860 / 21600,
                            x9 = w * 13917 / 21600,
                            x10 = w * 7602 / 21600,
                            x11 = w * 16577 / 21600,
                            y1 = h * 3890 / 21600,
                            y2 = h * 6080 / 21600,
                            y3 = h * 6797 / 21600,
                            y4 = h * 7437 / 21600,
                            y5 = h * 12877 / 21600,
                            y6 = h * 9705 / 21600,
                            y7 = h * 12007 / 21600,
                            y8 = h * 13987 / 21600;
                        const y9: number = h * 8382 / 21600, y10: number = h * 14277 / 21600, y11: number = h * 14915 / 21600;

                        d_val = "M" + x3 + "," + 0 +
                            " L" + x8 + "," + y2 +
                            " L" + x2 + "," + y3 +
                            " L" + x11 + "," + y7 +
                            " L" + x6 + "," + y5 +
                            " L" + w + "," + h +
                            " L" + x5 + "," + y11 +
                            " L" + x7 + "," + y8 +
                            " L" + x1 + "," + y6 +
                            " L" + x10 + "," + y9 +
                            " L" + 0 + "," + y1 + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "cube":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        const adj: any = 25000 * refr;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        const d_val: any = undefined;
                        cnstVal2 = 100000 * refr;
                        const ss: any = Math.min(w, h);
                        y4 = undefined;
                        x4 = undefined;
                        const a: any = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        y1 = ss * a / cnstVal2;
                        y4 = h - y1;
                        x4 = w - y1;
                        d_val = "M" + 0 + "," + y1 +
                            " L" + y1 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y4 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h +
                            ` zM` + 0 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " M" + x4 + "," + y1 +
                            " L" + w + "," + 0 +
                            "M" + x4 + "," + y1 +
                            " L" + x4 + "," + h;

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bevel":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        const adj: any = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 50000 * refr;
                        cnstVal2 = 100000 * refr;
                        const ss: any = Math.min(w, h);
                        let a: any = undefined, x1, x2, y2;
                        const a: any = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        y2 = h - x1;
                        d_val = "M" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + h +
                            " L" + 0 + "," + h +
                            ` z M` + x1 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x2 + "," + y2 +
                            " L" + x1 + "," + y2 +
                            ` z M` + 0 + "," + 0 +
                            " L" + x1 + "," + x1 +
                            " M" + 0 + "," + h +
                            " L" + x1 + "," + y2 +
                            " M" + w + "," + 0 +
                            " L" + x2 + "," + x1 +
                            " M" + w + "," + h +
                            " L" + x2 + "," + y2;

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "foldedCorner":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        const adj: any = 16667 * refr;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 50000 * refr;
                        cnstVal2 = 100000 * refr;
                        const ss: any = Math.min(w, h);
                        let a: any = undefined, dy2, dy1, x1, x2, y2, y1;
                        const a: any = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        const dy2: any = ss * a / cnstVal2;
                        dy1 = dy2 / 5;
                        x1 = w - dy2;
                        x2 = x1 + dy1;
                        y2 = h - dy2;
                        y1 = y2 + dy1;
                        d_val = "M" + x1 + "," + h +
                            " L" + x2 + "," + y1 +
                            " L" + w + "," + y2 +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + 0 + "," + 0 +
                            " L" + w + "," + 0 +
                            " L" + w + "," + y2;

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "cloud":
                    case "cloudCallout":
                        const d1: any = PPTXCalloutShapes.genCloudCallout(w, h, node, slideFactor, shapType);
                        result += "<path d='" + d1 + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "smileyFace":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        const adj: any = 4653 * refr;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 50000 * refr;
                        cnstVal2 = 100000 * refr;
                        const cnstVal3: any = 4653 * refr;
                        const ss: any = Math.min(w, h);
                        let a: any = undefined, x1, x2, x3, x4, y1, y3, dy2, y2, y4, dy3, y5, wR, hR, wd2, hd2;
                        wd2 = w / 2;
                        hd2 = h / 2;
                        const a: any = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
                        x1 = w * 4969 / 21699;
                        x2 = w * 6215 / 21600;
                        x3 = w * 13135 / 21600;
                        x4 = w * 16640 / 21600;
                        y1 = h * 7570 / 21600;
                        y3 = h * 16515 / 21600;
                        const dy2: any = h * a / cnstVal2;
                        y2 = y3 - dy2;
                        y4 = y3 + dy2;
                        const dy3: any = h * a / cnstVal1;
                        const y5: any = y4 + dy3;
                        const wR: any = w * 1125 / 21600;
                        const hR: any = h * 1125 / 21600;
                        const cX1: any = x2 - wR * Math.cos(Math.PI);
                        const cY1: any = y1 - hR * Math.sin(Math.PI);
                        const cX2: any = x3 - wR * Math.cos(Math.PI);
                        d_val = //eyes
                            PPTXShapeUtils.shapeArc(cX1, cY1, wR, hR, 180, 540, false) +
                            PPTXShapeUtils.shapeArc(cX2, cY1, wR, hR, 180, 540, false) +
                            //mouth
                            " M" + x1 + "," + y2 +
                            " Q" + wd2 + "," + y5 + " " + x4 + "," + y2 +
                            " Q" + wd2 + "," + y5 + " " + x1 + "," + y2 +
                            //head
                            " M" + 0 + "," + hd2 +
                            PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "verticalScroll":
                    case "horizontalScroll":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr: any = slideFactor;
                        const adj: any = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 25000 * refr;
                        cnstVal2 = 100000 * refr;
                        const ss: any = Math.min(w, h);
                        const t: number = 0, l: number = 0, b: number = h, r = w;
                        let a: any = undefined, ch, ch2, ch4;
                        const a: any = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        const ch: any = ss * a / cnstVal2;
                        const ch2: any = ch / 2;
                        const ch4: any = ch / 4;
                        if (shapType == "verticalScroll") {
                            let x3, x4, x6, x7, x5, y3, y4;
                            x3 = ch + ch2;
                            x4 = ch + ch;
                            const x6: any = r - ch;
                            const x7: any = r - ch2;
                            const x5: any = x6 - ch2;
                            y3 = b - ch;
                            y4 = b - ch2;

                            d_val = "M" + ch + "," + y3 +
                                " L" + ch + "," + ch2 +
                                PPTXShapeUtils.shapeArc(x3, ch2, ch2, ch2, 180, 270, false).replace("M", "L") +
                                " L" + x7 + "," + t +
                                PPTXShapeUtils.shapeArc(x7, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                                " L" + x6 + "," + ch +
                                " L" + x6 + "," + y4 +
                                PPTXShapeUtils.shapeArc(x5, y4, ch2, ch2, 0, 90, false).replace("M", "L") +
                                " L" + ch2 + "," + b +
                                PPTXShapeUtils.shapeArc(ch2, y4, ch2, ch2, 90, 270, false).replace("M", "L") +
                                ` z M` + x3 + "," + t +
                                PPTXShapeUtils.shapeArc(x3, ch2, ch2, ch2, 270, 450, false).replace("M", "L") +
                                PPTXShapeUtils.shapeArc(x3, x3 / 2, ch4, ch4, 90, 270, false).replace("M", "L") +
                                " L" + x4 + "," + ch2 +
                                " M" + x6 + "," + ch +
                                " L" + x3 + "," + ch +
                                " M" + ch + "," + y4 +
                                PPTXShapeUtils.shapeArc(ch2, y4, ch2, ch2, 0, 270, false).replace("M", "L") +
                                PPTXShapeUtils.shapeArc(ch2, (y4 + y3) / 2, ch4, ch4, 270, 450, false).replace("M", "L") +
                                ` z M` + ch + "," + y4 +
                                " L" + ch + "," + y3;
                        } else if (shapType == "horizontalScroll") {
                            y3, y4, y6, y7, y5, x3, x4;
                            y3 = ch + ch2;
                            y4 = ch + ch;
                            const y6: any = b - ch;
                            const y7: any = b - ch2;
                            const y5: any = y6 - ch2;
                            x3 = r - ch;
                            x4 = r - ch2;

                            d_val = "M" + l + "," + y3 +
                                PPTXShapeUtils.shapeArc(ch2, y3, ch2, ch2, 180, 270, false).replace("M", "L") +
                                " L" + x3 + "," + ch +
                                " L" + x3 + "," + ch2 +
                                PPTXShapeUtils.shapeArc(x4, ch2, ch2, ch2, 180, 360, false).replace("M", "L") +
                                " L" + r + "," + y5 +
                                PPTXShapeUtils.shapeArc(x4, y5, ch2, ch2, 0, 90, false).replace("M", "L") +
                                " L" + ch + "," + y6 +
                                " L" + ch + "," + y7 +
                                PPTXShapeUtils.shapeArc(ch2, y7, ch2, ch2, 0, 180, false).replace("M", "L") +
                                ` zM` + x4 + "," + ch +
                                PPTXShapeUtils.shapeArc(x4, ch2, ch2, ch2, 90, -180, false).replace("M", "L") +
                                PPTXShapeUtils.shapeArc((x3 + x4) / 2, ch2, ch4, ch4, 180, 0, false).replace("M", "L") +
                                ` z M` + x4 + "," + ch +
                                " L" + x3 + "," + ch +
                                " M" + ch2 + "," + y4 +
                                " L" + ch2 + "," + y3 +
                                PPTXShapeUtils.shapeArc(y3 / 2, y3, ch4, ch4, 180, 360, false).replace("M", "L") +
                                PPTXShapeUtils.shapeArc(ch2, y3, ch2, ch2, 0, 180, false).replace("M", "L") +
                                " M" + ch + "," + y3 +
                                " L" + ch + "," + y6;
                        }

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "wedgeEllipseCallout":
                        const d_val: any = PPTXCalloutShapes.genWedgeEllipseCallout(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "wedgeRectCallout":
                        const d_val: any = PPTXCalloutShapes.genWedgeRectCallout(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "wedgeRoundRectCallout":
                        const d_val: any = PPTXCalloutShapes.genWedgeRoundRectCallout(w, h, node, slideFactor);
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "accentBorderCallout1":
                    case "accentBorderCallout2":
                    case "accentBorderCallout3":
                    case "borderCallout1":
                    case "borderCallout2":
                    case "borderCallout3":
                    case "accentCallout1":
                    case "accentCallout2":
                    case "accentCallout3":
                    case "callout1":
                    case "callout2":
                    case "callout3":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr: any = slideFactor;
                        let sAdj1: any = undefined, adj1 = 18750 * refr;
                        let sAdj2: any = undefined, adj2 = -8333 * refr;
                        let sAdj3: any = undefined, adj3 = 18750 * refr;
                        let sAdj4: any = undefined, adj4 = -16667 * refr;
                        let sAdj5: any = undefined, adj5 = 100000 * refr;
                        let sAdj6, adj6 = -16667 * refr;
                        let sAdj7, adj7 = 112963 * refr;
                        let sAdj8, adj8 = -8333 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * refr;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * refr;
                                } else if (sAdj_name == "adj5") {
                                    const sAdj5: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj5: any = parseInt(sAdj5.substr(4)) * refr;
                                } else if (sAdj_name == "adj6") {
                                    const sAdj6: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj6: any = parseInt(sAdj6.substr(4)) * refr;
                                } else if (sAdj_name == "adj7") {
                                    const sAdj7: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj7: any = parseInt(sAdj7.substr(4)) * refr;
                                } else if (sAdj_name == "adj8") {
                                    const sAdj8: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj8: any = parseInt(sAdj8.substr(4)) * refr;
                                }
                            }
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 100000 * refr;
                        const isBorder: any = true;
                        switch (shapType) {
                            case "borderCallout1":
                            case "callout1":
                                if (shapType == "borderCallout1") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 112500 * refr;
                                    const adj4: any = -38333 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "borderCallout2":
                            case "callout2":
                                if (shapType == "borderCallout2") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 18750 * refr;
                                    const adj4: any = -16667 * refr;

                                    const adj5: any = 112500 * refr;
                                    const adj6: any = -46667 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined, y3 = undefined, x3 = undefined;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;

                                break;
                            case "borderCallout3":
                            case "callout3":
                                if (shapType == "borderCallout3") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 18750 * refr;
                                    const adj4: any = -16667 * refr;

                                    const adj5: any = 100000 * refr;
                                    const adj6: any = -16667 * refr;

                                    const adj7: any = 112963 * refr;
                                    const adj8: any = -8333 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined, y3 = undefined, x3 = undefined, y4 = undefined, x4 = undefined;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;

                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " L" + x3 + "," + y3 +

                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2;
                                break;
                            case "accentBorderCallout1":
                            case "accentCallout1":
                                if (shapType == "accentBorderCallout1") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }

                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 112500 * refr;
                                    const adj4: any = -38333 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                            case "accentBorderCallout2":
                            case "accentCallout2":
                                if (shapType == "accentBorderCallout2") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 18750 * refr;
                                    const adj4: any = -16667 * refr;
                                    const adj5: any = 112500 * refr;
                                    const adj6: any = -46667 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined, y3 = undefined, x3 = undefined;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;

                                break;
                            case "accentBorderCallout3":
                            case "accentCallout3":
                                if (shapType == "accentBorderCallout3") {
                                    const isBorder: any = true;
                                } else {
                                    const isBorder: any = false;
                                }
                                const isBorder: any = true;
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    const adj3: any = 18750 * refr;
                                    const adj4: any = -16667 * refr;
                                    const adj5: any = 100000 * refr;
                                    const adj6: any = -16667 * refr;
                                    const adj7: any = 112963 * refr;
                                    const adj8: any = -8333 * refr;
                                }
                                let y1: number = undefined, x1: number = undefined, y2: number = undefined, x2 = undefined, y3 = undefined, x3 = undefined, y4 = undefined, x4 = undefined;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = "M" + 0 + "," + 0 +
                                    " L" + w + "," + 0 +
                                    " L" + w + "," + h +
                                    " L" + 0 + "," + h +
                                    ` z M` + x1 + "," + y1 +
                                    " L" + x2 + "," + y2 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x4 + "," + y4 +
                                    " L" + x3 + "," + y3 +
                                    " L" + x2 + "," + y2 +

                                    " M" + x1 + "," + 0 +
                                    " L" + x1 + "," + h;
                                break;
                        }

                        //console.log("shapType: ", shapType, ",isBorder:", isBorder)
                        //if(isBorder){
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        //}else{
                        //    result += "<path d='"+d_val+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        //        "' stroke='none' />";

                        //}
                        break;
                    case "leftRightRibbon":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr: any = slideFactor;
                        let sAdj1: any = undefined, adj1 = 50000 * refr;
                        let sAdj2: any = undefined, adj2 = 50000 * refr;
                        let sAdj3: any = undefined, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 33333 * refr;
                        cnstVal2 = 100000 * refr;
                        const cnstVal3: any = 200000 * refr;
                        const cnstVal4: any = 400000 * refr;
                        const ss: any = Math.min(w, h);
                        a3, maxAdj1, a1, w1, maxAdj2, a2, x1, x4, dy1, dy2, ly1, ry4, ly2, ry3, ly4, ry1,
                            ly3, ry2, hR, x2, x3, y1, y2, wd32 = w / 32, vc = h / 2;
                        hc = w / 2;

                        const a3: any = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
                        maxAdj1 = cnstVal2 - a3;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const w1: any = hc - wd32;
                        const maxAdj2: any = cnstVal2 * w1 / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        x1 = ss * a2 / cnstVal2;
                        x4 = w - x1;
                        dy1 = h * a1 / cnstVal3;
                        const dy2: any = h * a3 / -cnstVal3;
                        const ly1: any = vc + dy2 - dy1;
                        const ry4: any = vc + dy1 - dy2;
                        const ly2: any = ly1 + dy1;
                        const ry3: any = h - ly2;
                        const ly4: any = ly2 * 2;
                        const ry1: any = h - ly4;
                        const ly3: any = ly4 - ly1;
                        const ry2: any = h - ly3;
                        const hR: any = a3 * ss / cnstVal4;
                        x2 = hc - wd32;
                        x3 = hc + wd32;
                        y1 = ly1 + hR;
                        y2 = ry2 - hR;

                        d_val = "M" + 0 + "," + ly2 +
                            "L" + x1 + "," + 0 +
                            "L" + x1 + "," + ly1 +
                            "L" + hc + "," + ly1 +
                            PPTXShapeUtils.shapeArc(hc, y1, wd32, hR, 270, 450, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(hc, y2, wd32, hR, 270, 90, false).replace("M", "L") +
                            "L" + x4 + "," + ry2 +
                            "L" + x4 + "," + ry1 +
                            "L" + w + "," + ry3 +
                            "L" + x4 + "," + h +
                            "L" + x4 + "," + ry4 +
                            "L" + hc + "," + ry4 +
                            PPTXShapeUtils.shapeArc(hc, ry4 - hR, wd32, hR, 90, 180, false).replace("M", "L") +
                            "L" + x2 + "," + ly3 +
                            "L" + x1 + "," + ly3 +
                            "L" + x1 + "," + ly4 +
                            ` zM` + x3 + "," + y1 +
                            "L" + x3 + "," + ry2 +
                            "M" + x2 + "," + y2 +
                            "L" + x2 + "," + ly3;

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "ribbon":
                    case "ribbon2":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 16667 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 25000 * slideFactor;
                        cnstVal2 = 33333 * slideFactor;
                        const cnstVal3: any = 75000 * slideFactor;
                        const cnstVal4: any = 100000 * slideFactor;
                        const cnstVal5: any = 200000 * slideFactor;
                        const cnstVal6: any = 400000 * slideFactor;
                        const hc: number = w / 2, t: number = 0, l: number = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        a1 = undefined, a2, x10, dx2, x2, x9, x3, x8, x5, x6, x4, x7, y1, y2, y4, y3, hR, y6;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        const a2: any = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        const x10: any = r - wd8;
                        const dx2: any = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        const x9: any = hc + dx2;
                        x3 = x2 + wd32;
                        const x8: any = x9 - wd32;
                        const x5: any = x2 + wd8;
                        const x6: any = x9 - wd8;
                        x4 = x5 - wd32;
                        const x7: any = x6 + wd32;
                        const hR: any = h * a1 / cnstVal6;
                        if (shapType == "ribbon2") {
                            dy1, dy2, y7;
                            dy1 = h * a1 / cnstVal5;
                            y1 = b - dy1;
                            const dy2: any = h * a1 / cnstVal4;
                            y2 = b - dy2;
                            y4 = t + dy2;
                            y3 = (y4 + b) / 2;
                            const y6: any = b - hR;///////////////////
                            const y7: any = y1 - hR;

                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y3 +
                                " L" + l + "," + y4 +
                                " L" + x2 + "," + y4 +
                                " L" + x2 + "," + hR +
                                PPTXShapeUtils.shapeArc(x3, hR, wd32, hR, 180, 270, false).replace("M", "L") +
                                " L" + x8 + "," + t +
                                PPTXShapeUtils.shapeArc(x8, hR, wd32, hR, 270, 360, false).replace("M", "L") +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + r + "," + y4 +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + b +
                                " L" + x7 + "," + b +
                                PPTXShapeUtils.shapeArc(x7, y6, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + x8 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x8, y7, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x3 + "," + y2 +
                                PPTXShapeUtils.shapeArc(x3, y7, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x4 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x4, y6, wd32, hR, 270, 450, false).replace("M", "L") +
                                ` z M` + x5 + "," + y2 +
                                " L" + x5 + "," + y6 +
                                "M" + x6 + "," + y6 +
                                " L" + x6 + "," + y2 +
                                "M" + x2 + "," + y7 +
                                " L" + x2 + "," + y4 +
                                "M" + x9 + "," + y4 +
                                " L" + x9 + "," + y7;
                        } else if (shapType == "ribbon") {
                            y5;
                            y1 = h * a1 / cnstVal5;
                            y2 = h * a1 / cnstVal4;
                            y4 = b - y2;
                            y3 = y4 / 2;
                            const y5: any = b - hR; ///////////////////////
                            const y6: any = y2 - hR;
                            d_val = "M" + l + "," + t +
                                " L" + x4 + "," + t +
                                PPTXShapeUtils.shapeArc(x4, hR, wd32, hR, 270, 450, false).replace("M", "L") +
                                " L" + x3 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x3, y6, wd32, hR, 270, 90, false).replace("M", "L") +
                                " L" + x8 + "," + y2 +
                                PPTXShapeUtils.shapeArc(x8, y6, wd32, hR, 90, -90, false).replace("M", "L") +
                                " L" + x7 + "," + y1 +
                                PPTXShapeUtils.shapeArc(x7, hR, wd32, hR, 90, 270, false).replace("M", "L") +
                                " L" + r + "," + t +
                                " L" + x10 + "," + y3 +
                                " L" + r + "," + y4 +
                                " L" + x9 + "," + y4 +
                                " L" + x9 + "," + y5 +
                                PPTXShapeUtils.shapeArc(x8, y5, wd32, hR, 0, 90, false).replace("M", "L") +
                                " L" + x3 + "," + b +
                                PPTXShapeUtils.shapeArc(x3, y5, wd32, hR, 90, 180, false).replace("M", "L") +
                                " L" + x2 + "," + y4 +
                                " L" + l + "," + y4 +
                                " L" + wd8 + "," + y3 +
                                ` z M` + x5 + "," + hR +
                                " L" + x5 + "," + y2 +
                                "M" + x6 + "," + y2 +
                                " L" + x6 + "," + hR +
                                "M" + x2 + "," + y4 +
                                " L" + x2 + "," + y6 +
                                "M" + x9 + "," + y6 +
                                " L" + x9 + "," + y4;
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "doubleWave":
                    case "wave":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = (shapType == "doubleWave") ? 6250 * slideFactor : 12500 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 0;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        const d_val: any = undefined;
                        cnstVal2 = -10000 * slideFactor;
                        const cnstVal3: any = 50000 * slideFactor;
                        const cnstVal4: any = 100000 * slideFactor;
                        const hc: number = w / 2, t: number = 0, l: number = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
                        if (shapType == "doubleWave") {
                            const cnstVal1 = 12500 * slideFactor;
                            a1 = undefined, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx8, x8, dx3, x3, dx4, x4, x5, x6, x7, x9, x15, x10, x11, x12, x13, x14;
                            const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            const a2: any = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            const dy2: any = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            const y5: any = y4 - dy2;
                            const y6: any = y4 + dy2;
                            const of2: any = w * a2 / cnstVal3;
                            const dx2: any = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            const dx8: any = (of2 > 0) ? of2 : 0;
                            const x8: any = r - dx8;
                            const dx3: any = (dx2 + x8) / 6;
                            x3 = x2 + dx3;
                            const dx4: any = (dx2 + x8) / 3;
                            x4 = x2 + dx4;
                            const x5: any = (x2 + x8) / 2;
                            const x6: any = x5 + dx3;
                            const x7: any = (x6 + x8) / 2;
                            const x9: any = l + dx8;
                            const x15: any = r + dx2;
                            const x10: any = x9 + dx3;
                            const x11: any = x9 + dx4;
                            const x12: any = (x9 + x15) / 2;
                            const x13: any = x12 + dx3;
                            const x14: any = (x13 + x15) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " C" + x6 + "," + y2 + " " + x7 + "," + y3 + " " + x8 + "," + y1 +
                                " L" + x15 + "," + y4 +
                                " C" + x14 + "," + y6 + " " + x13 + "," + y5 + " " + x12 + "," + y4 +
                                " C" + x11 + "," + y6 + " " + x10 + "," + y5 + " " + x9 + "," + y4 +
                                " z";
                        } else if (shapType == "wave") {
                            const cnstVal5 = 20000 * slideFactor;
                            a1 = undefined, a2, y1, dy2, y2, y3, y4, y5, y6, of2, dx2, x2, dx5, x5, dx3, x3, x4, x6, x10, x7, x8;
                            const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            const a2: any = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            const dy2: any = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            const y5: any = y4 - dy2;
                            const y6: any = y4 + dy2;
                            const of2: any = w * a2 / cnstVal3;
                            const dx2: any = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            const dx5: any = (of2 > 0) ? of2 : 0;
                            const x5: any = r - dx5;
                            const dx3: any = (dx2 + x5) / 3;
                            x3 = x2 + dx3;
                            x4 = (x3 + x5) / 2;
                            const x6: any = l + dx5;
                            const x10: any = r + dx2;
                            const x7: any = x6 + dx3;
                            const x8: any = (x7 + x10) / 2;

                            d_val = "M" + x2 + "," + y1 +
                                " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
                                " L" + x10 + "," + y4 +
                                " C" + x8 + "," + y6 + " " + x7 + "," + y5 + " " + x6 + "," + y4 +
                                " z";
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "ellipseRibbon":
                    case "ellipseRibbon2":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        const d_val: any = undefined;
                        const cnstVal1: any = 25000 * slideFactor;
                        const cnstVal3: any = 75000 * slideFactor;
                        const cnstVal4: any = 100000 * slideFactor;
                        const cnstVal5: any = 200000 * slideFactor;
                        const hc: number = w / 2, t: number = 0, l: number = 0, b = h, r = w, wd8 = w / 8;
                        a1 = undefined, a2, q10, q11, q12, minAdj3, a3, dx2, x2, x3, x4, x5, x6, dy1, f1, q1, q2,
                            cx1, cx2, q1, dy3, q3, q4, q5, rh, q8, cx4, q9, cx5;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal4) ? cnstVal4 : adj1;
                        const a2: any = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        const q10: any = cnstVal4 - a1;
                        const q11: any = q10 / 2;
                        const q12: any = a1 - q11;
                        const minAdj3: any = (0 > q12) ? 0 : q12;
                        const a3: any = (adj3 < minAdj3) ? minAdj3 : (adj3 > a1) ? a1 : adj3;
                        const dx2: any = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x3 = x2 + wd8;
                        x4 = r - x3;
                        const x5: any = r - x2;
                        const x6: any = r - wd8;
                        dy1 = h * a3 / cnstVal4;
                        const f1: any = 4 * dy1 / w;
                        const q1: any = x3 * x3 / w;
                        const q2: any = x3 - q1;
                        const cx1: any = x3 / 2;
                        const cx2: any = r - cx1;
                        const q1: any = h * a1 / cnstVal4;
                        const dy3: any = q1 - dy1;
                        const q3: any = x2 * x2 / w;
                        const q4: any = x2 - q3;
                        const q5: any = f1 * q4;
                        const rh: any = b - q1;
                        const q8: any = dy1 * 14 / 16;
                        const cx4: any = x2 / 2;
                        const q9: any = f1 * cx4;
                        const cx5: any = r - cx4;
                        if (shapType == "ellipseRibbon") {
                            y1 = undefined, cy1 = undefined, y3 = undefined, q6 = undefined, q7 = undefined, cy3 = undefined, y2 = undefined, y5 = undefined, y6 = undefined,
                                cy4, cy6, y7, cy7, y8;
                            y1 = f1 * q2;
                            const cy1: any = f1 * cx1;
                            y3 = q5 + dy3;
                            const q6: any = dy1 + dy3 - y3;
                            const q7: any = q6 + dy1;
                            const cy3: any = q7 + dy3;
                            y2 = (q8 + rh) / 2;
                            const y5: any = q5 + rh;
                            const y6: any = y3 + rh;
                            const cy4: any = q9 + rh;
                            const cy6: any = cy3 + rh;
                            const y7: any = y1 + dy3;
                            const cy7: any = q1 + q1 - y7;
                            const y8: any = b - dy1;
                            //
                            d_val = "M" + l + "," + t +
                                " Q" + cx1 + "," + cy1 + " " + x3 + "," + y1 +
                                " L" + x2 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x5 + "," + y3 +
                                " L" + x4 + "," + y1 +
                                " Q" + cx2 + "," + cy1 + " " + r + "," + t +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + rh +
                                " Q" + cx5 + "," + cy4 + " " + x5 + "," + y5 +
                                " L" + x5 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x2 + "," + y6 +
                                " L" + x2 + "," + y5 +
                                " Q" + cx4 + "," + cy4 + " " + l + "," + rh +
                                " L" + wd8 + "," + y2 +
                                ` zM` + x2 + "," + y5 +
                                " L" + x2 + "," + y3 +
                                "M" + x5 + "," + y3 +
                                " L" + x5 + "," + y5 +
                                "M" + x3 + "," + y1 +
                                " L" + x3 + "," + y7 +
                                "M" + x4 + "," + y7 +
                                " L" + x4 + "," + y1;
                        } else if (shapType == "ellipseRibbon2") {
                            u1, y1, cu1, cy1, q3, q5, u3, y3, q6, q7, cu3, cy3, rh, q8, u2, y2,
                                u5, y5, u6, y6, cu4, cy4, cu6, cy6, u7, y7, cu7, cy7;
                            const u1: any = f1 * q2;
                            y1 = b - u1;
                            const cu1: any = f1 * cx1;
                            const cy1: any = b - cu1;
                            const u3: any = q5 + dy3;
                            y3 = b - u3;
                            const q6: any = dy1 + dy3 - u3;
                            const q7: any = q6 + dy1;
                            const cu3: any = q7 + dy3;
                            const cy3: any = b - cu3;
                            const u2: any = (q8 + rh) / 2;
                            y2 = b - u2;
                            const u5: any = q5 + rh;
                            const y5: any = b - u5;
                            const u6: any = u3 + rh;
                            const y6: any = b - u6;
                            const cu4: any = q9 + rh;
                            const cy4: any = b - cu4;
                            const cu6: any = cu3 + rh;
                            const cy6: any = b - cu6;
                            const u7: any = u1 + dy3;
                            const y7: any = b - u7;
                            const cu7: any = q1 + q1 - u7;
                            const cy7: any = b - cu7;
                            //
                            d_val = "M" + l + "," + b +
                                " L" + wd8 + "," + y2 +
                                " L" + l + "," + q1 +
                                " Q" + cx4 + "," + cy4 + " " + x2 + "," + y5 +
                                " L" + x2 + "," + y6 +
                                " Q" + hc + "," + cy6 + " " + x5 + "," + y6 +
                                " L" + x5 + "," + y5 +
                                " Q" + cx5 + "," + cy4 + " " + r + "," + q1 +
                                " L" + x6 + "," + y2 +
                                " L" + r + "," + b +
                                " Q" + cx2 + "," + cy1 + " " + x4 + "," + y1 +
                                " L" + x5 + "," + y3 +
                                " Q" + hc + "," + cy3 + " " + x2 + "," + y3 +
                                " L" + x3 + "," + y1 +
                                " Q" + cx1 + "," + cy1 + " " + l + "," + b +
                                ` zM` + x2 + "," + y3 +
                                " L" + x2 + "," + y5 +
                                "M" + x5 + "," + y5 +
                                " L" + x5 + "," + y3 +
                                "M" + x3 + "," + y7 +
                                " L" + x3 + "," + y1 +
                                "M" + x4 + "," + y1 +
                                " L" + x4 + "," + y7;
                        }
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "line":
                    case "straightConnector1":
                    case "bentConnector4":
                    case "bentConnector5":
                    case "curvedConnector2":
                    case "curvedConnector3":
                    case "curvedConnector4":
                    case "curvedConnector5":
                        // if (isFlipV) {
                        //     result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color +
                        //         "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        // } else {
                        result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color +
                            "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                        //}
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                        }
                        result += "/>";
                        break;
                    case "rightArrow":
                        const points: any = PPTXArrowShapes.genRightArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftArrow":
                        const points: any = PPTXArrowShapes.genLeftArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "downArrow":
                    case "flowChartOffpageConnector":
                        const points: any = PPTXArrowShapes.genDownArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "upArrow":
                        const points: any = PPTXArrowShapes.genUpArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "leftRightArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, sAdj1_val = 0.25;
                        let sAdj2: any = undefined, sAdj2_val = 0.25;
                        const max_sAdj2_const: any = w / h;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj1_val: any = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                                    const sAdj2_val: any = (sAdj2_val2) / max_sAdj2_const;
                                }
                            }
                        }
                        //console.log("w: "+w+"\nh: "+h+"\nsAdj1: "+sAdj1_val+"\nsAdj2: "+sAdj2_val);

                        result += " <polygon points='0 " + h / 2 + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + "," + (1 - sAdj2_val) * w + " " + (1 - sAdj1_val) * h +
                            "," + (1 - sAdj2_val) * w + " " + h + "," + w + " " + h / 2 + ", " + (1 - sAdj2_val) * w + " 0," + (1 - sAdj2_val) * w + " " + sAdj1_val * h + "," +
                            sAdj2_val * w + " " + sAdj1_val * h + "," + sAdj2_val * w + " 0' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "upDownArrow":
                        const points: any = PPTXArrowShapes.genUpDownArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
                        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                        break;
                    case "quadArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 22500 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 22500 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 22500 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        hc = w / 2;
                        let a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, y3, y4, y5, y6, maxAdj1, maxAdj3;
                        minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        const q1: any = cnstVal2 - maxAdj1;
                        const maxAdj3: any = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        const dx2: any = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        const x5: any = hc + dx2;
                        const dx3: any = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        const x6: any = w - x1;
                        y2 = vc - dx2;
                        const y5: any = vc + dx2;
                        y3 = vc - dx3;
                        y4 = vc + dx3;
                        const y6: any = h - x1;
                        d_val = "M" + 0 + "," + vc +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + vc +
                            " L" + x6 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y6 +
                            " L" + x5 + "," + y6 +
                            " L" + hc + "," + h +
                            " L" + x2 + "," + y6 +
                            " L" + x3 + "," + y6 +
                            " L" + x3 + "," + y4 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightUpArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        hc = w / 2;
                        let a1, a2, a3, q1, x1, x2, dx2, x3, dx3, x4, x5, x6, y2, dy2, y3, y4, y5, maxAdj1, maxAdj3;
                        minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        const q1: any = cnstVal2 - maxAdj1;
                        const maxAdj3: any = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        const dx2: any = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        const x5: any = hc + dx2;
                        const dx3: any = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        const x6: any = w - x1;
                        const dy2: any = minWH * a2 / cnstVal1;
                        y2 = h - dy2;
                        y4 = h - dx2;
                        y3 = y4 - dx3;
                        const y5: any = y4 + dx3;
                        d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + hc + "," + 0 +
                            " L" + x5 + "," + x1 +
                            " L" + x4 + "," + x1 +
                            " L" + x4 + "," + y3 +
                            " L" + x6 + "," + y3 +
                            " L" + x6 + "," + y2 +
                            " L" + w + "," + y4 +
                            " L" + x6 + "," + h +
                            " L" + x6 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftUpArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        hc = w / 2;
                        let a1, a2, a3, x1, x2, dx4, dx3, x3, x4, x5, y2, y3, y4, y5, maxAdj1, maxAdj3;
                        minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        const maxAdj3: any = cnstVal2 - maxAdj1;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        const dx2: any = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        y2 = h - dx2;
                        const dx4: any = minWH * a2 / cnstVal2;
                        x4 = w - dx4;
                        y4 = h - dx4;
                        const dx3: any = minWH * a1 / cnstVal3;
                        x3 = x4 - dx3;
                        const x5: any = x4 + dx3;
                        y3 = y4 - dx3;
                        const y5: any = y4 + dx3;
                        d_val = "M" + 0 + "," + y4 +
                            " L" + x1 + "," + y2 +
                            " L" + x1 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + x1 +
                            " L" + x2 + "," + x1 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + x1 +
                            " L" + x5 + "," + x1 +
                            " L" + x5 + "," + y5 +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bentUpArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        hc = w / 2;
                        let a1, a2, a3, dx1, x1, dx2, x2, dx3, x3, x4, y1, y2, dy2;
                        minWH = Math.min(w, h);
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        y1 = minWH * a3 / cnstVal2;
                        const dx1: any = minWH * a2 / cnstVal1;
                        x1 = w - dx1;
                        const dx3: any = minWH * a2 / cnstVal2;
                        x3 = w - dx3;
                        const dx2: any = minWH * a1 / cnstVal3;
                        x2 = x3 - dx2;
                        x4 = x3 + dx2;
                        const dy2: any = minWH * a1 / cnstVal2;
                        y2 = h - dy2;
                        d_val = "M" + 0 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + x3 + "," + 0 +
                            " L" + w + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x4 + "," + h +
                            " L" + 0 + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "bentArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 43750 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        a1 = undefined, a2, a3, a4, x3, x4, y3, y4, y5, y6, maxAdj1, maxAdj4;
                        minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > cnstVal1) a3 = cnstVal1
                        else a3 = adj3
                        th = undefined, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2;
                        const th: any = minWH * a1 / cnstVal2;
                        const aw2: any = minWH * a2 / cnstVal2;
                        const th2: any = th / 2;
                        const dh2: any = aw2 - th2;
                        const ah: any = minWH * a3 / cnstVal2;
                        const bw: any = w - ah;
                        const bh: any = h - dh2;
                        const bs: any = (bw < bh) ? bw : bh;
                        const maxAdj4: any = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        const bd: any = minWH * a4 / cnstVal2;
                        const bd3: any = bd - th;
                        const bd2: any = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x4 = w - ah;
                        y3 = dh2 + th;
                        y4 = y3 + dh2;
                        const y5: any = dh2 + bd;
                        const y6: any = y3 + bd2;

                        d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + y5 +
                            PPTXShapeUtils.shapeArc(bd, y5, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + dh2 +
                            " L" + x4 + "," + 0 +
                            " L" + w + "," + aw2 +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            PPTXShapeUtils.shapeArc(x3, y6, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "uturnArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 43750 * slideFactor;
                        let sAdj5: any = undefined, adj5 = 75000 * slideFactor;
                        const cnstVal1: any = 25000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj5") {
                                    const sAdj5: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj5: any = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        a1 = undefined, a2, a3, a4, a5, q1, q2, q3, x3, x4, x5, x6, x7, x8, x9, y4, y5, minAdj5, maxAdj1, maxAdj3, maxAdj4;
                        minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        const q2: any = a1 * minWH / h;
                        const q3: any = cnstVal2 - q2;
                        const maxAdj3: any = q3 * h / minWH;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        const q1: any = a3 + a1;
                        const minAdj5: any = q1 * minWH / h;
                        if (adj5 < minAdj5) a5 = minAdj5
                        else if (adj5 > cnstVal2) a5 = cnstVal2
                        else a5 = adj5

                        th = undefined, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2;
                        const th: any = minWH * a1 / cnstVal2;
                        const aw2: any = minWH * a2 / cnstVal2;
                        const th2: any = th / 2;
                        const dh2: any = aw2 - th2;
                        const y5: any = h * a5 / cnstVal2;
                        const ah: any = minWH * a3 / cnstVal2;
                        y4 = y5 - ah;
                        const x9: any = w - dh2;
                        const bw: any = x9 / 2;
                        const bs: any = (bw < y4) ? bw : y4;
                        const maxAdj4: any = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        const bd: any = minWH * a4 / cnstVal2;
                        const bd3: any = bd - th;
                        const bd2: any = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        const x8: any = w - aw2;
                        const x6: any = x8 - aw2;
                        const x7: any = x6 + dh2;
                        x4 = x9 - bd;
                        const x5: any = x7 - bd2;
                        cx = (th + x7) / 2
                        cy = (y4 + th) / 2
                        d_val = "M" + 0 + "," + h +
                            " L" + 0 + "," + bd +
                            PPTXShapeUtils.shapeArc(bd, bd, bd, bd, 180, 270, false).replace("M", "L") +
                            " L" + x4 + "," + 0 +
                            PPTXShapeUtils.shapeArc(x4, bd, bd, bd, 270, 360, false).replace("M", "L") +
                            " L" + x9 + "," + y4 +
                            " L" + w + "," + y4 +
                            " L" + x8 + "," + y5 +
                            " L" + x6 + "," + y4 +
                            " L" + x7 + "," + y4 +
                            " L" + x7 + "," + x3 +
                            PPTXShapeUtils.shapeArc(x5, x3, bd2, bd2, 0, -90, false).replace("M", "L") +
                            " L" + x3 + "," + th +
                            PPTXShapeUtils.shapeArc(x3, x3, bd2, bd2, 270, 180, false).replace("M", "L") +
                            " L" + th + "," + h + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "stripedRightArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 50000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        const cnstVal1: any = 100000 * slideFactor;
                        cnstVal2 = 200000 * slideFactor;
                        const cnstVal3: any = 84375 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        a1 = undefined, a2, x4, x5, dx5, x6, dx6, y1, dy1, y2, maxAdj2, vc = h / 2;
                        minWH = Math.min(w, h);
                        const maxAdj2: any = cnstVal3 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        x4 = minWH * 5 / 32;
                        const dx5: any = minWH * a2 / cnstVal1;
                        const x5: any = w - dx5;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        //dx6 = dy1*dx5/hd2;
                        //x6 = w-dx6;
                        const ssd8: number = minWH / 8, ssd16: number = minWH / 16, ssd32: number = minWH / 32;
                        d_val = "M" + 0 + "," + y1 +
                            " L" + ssd32 + "," + y1 +
                            " L" + ssd32 + "," + y2 +
                            " L" + 0 + "," + y2 + ` z M` + ssd16 + "," + y1 +
                            " L" + ssd8 + "," + y1 +
                            " L" + ssd8 + "," + y2 +
                            " L" + ssd16 + "," + y2 + ` z M` + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            " L" + x5 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x5 + "," + h +
                            " L" + x5 + "," + y2 +
                            " L" + x4 + "," + y2 + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "notchedRightArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 50000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        const cnstVal1: any = 100000 * slideFactor;
                        cnstVal2 = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                }
                            }
                        }
                        a1 = undefined, a2, x1, x2, dx2, y1, dy1, y2, maxAdj2, vc = h / 2, hd2 = vc;
                        minWH = Math.min(w, h);
                        const maxAdj2: any = cnstVal1 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        const dx2: any = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        x1 = dy1 * dx2 / hd2;
                        d_val = "M" + 0 + "," + y1 +
                            " L" + x2 + "," + y1 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + x2 + "," + y2 +
                            " L" + 0 + "," + y2 +
                            " L" + x1 + "," + vc + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "homePlate":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 50000 * slideFactor;
                        const cnstVal1: any = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let a: any = undefined, x1, dx1, maxAdj, vc = h / 2;
                        minWH = Math.min(w, h);
                        const maxAdj: any = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        const dx1: any = minWH * a / cnstVal1;
                        x1 = w - dx1;
                        d_val = "M" + 0 + "," + 0 +
                            " L" + x1 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x1 + "," + h +
                            " L" + 0 + "," + h + " z";

                        result += "<path  d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "chevron":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 50000 * slideFactor;
                        const cnstVal1: any = 100000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        let a: any = undefined, x1, dx1, x2, maxAdj, vc = h / 2;
                        minWH = Math.min(w, h);
                        const maxAdj: any = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        x1 = minWH * a / cnstVal1;
                        x2 = w - x1;
                        d_val = "M" + 0 + "," + 0 +
                            " L" + x2 + "," + 0 +
                            " L" + w + "," + vc +
                            " L" + x2 + "," + h +
                            " L" + 0 + "," + h +
                            " L" + x1 + "," + vc + " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";


                        break;
                    case "rightArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 64977 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        let maxAdj2: any = undefined, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, dx3, x3, x2, x1;
                        vc = h / 2;
                        let r: number = w, b: number = h;
                        l = 0, t = 0;
                        const ss: any = Math.min(w, h);
                        const maxAdj2: any = cnstVal1 * h / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal2 * w / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * ss / w;
                        const maxAdj4: any = cnstVal - q2;
                        const a4: any = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        const dy2: any = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        const dx3: any = ss * a3 / cnstVal2;
                        x3 = r - dx3;
                        x2 = w * a4 / cnstVal2;
                        x1 = x2 / 2;
                        d_val = "M" + l + "," + t +
                            " L" + x2 + "," + t +
                            " L" + x2 + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x3 + "," + y4 +
                            " L" + x3 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + b +
                            " L" + l + "," + b +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "downArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 64977 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        let maxAdj2: any = undefined, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, dy3, y3, y2, y1;
                        const hc: number = w / 2, r: number = w, b: number = h, l = 0, t = 0;
                        const ss: any = Math.min(w, h);

                        const maxAdj2: any = cnstVal1 * w / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal2 * h / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * ss / h;
                        const maxAdj4: any = cnstVal2 - q2;
                        const a4: any = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        const dx1: any = ss * a2 / cnstVal2;
                        const dx2: any = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        const dy3: any = ss * a3 / cnstVal2;
                        y3 = b - dy3;
                        y2 = h * a4 / cnstVal2;
                        y1 = y2 / 2;
                        d_val = "M" + l + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + y2 +
                            " L" + x3 + "," + y2 +
                            " L" + x3 + "," + y3 +
                            " L" + x4 + "," + y3 +
                            " L" + hc + "," + b +
                            " L" + x1 + "," + y3 +
                            " L" + x2 + "," + y3 +
                            " L" + x2 + "," + y2 +
                            " L" + l + "," + y2 +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 64977 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        let maxAdj2: any = undefined, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, dx2, x2, x3;
                        vc = h / 2;
                        let r: number = w, b: number = h;
                        l = 0, t = 0;
                        const ss: any = Math.min(w, h);

                        const maxAdj2: any = cnstVal1 * h / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal2 * w / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * ss / w;
                        const maxAdj4: any = cnstVal2 - q2;
                        const a4: any = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        const dy2: any = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        const dx2: any = w * a4 / cnstVal2;
                        x2 = r - dx2;
                        x3 = (x2 + r) / 2;
                        d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + r + "," + t +
                            " L" + r + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "upArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 64977 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        let maxAdj2: any = undefined, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2, x1, x2, x3, x4, y1, dy2, y2, y3;
                        const hc: number = w / 2, r: number = w, b: number = h, l = 0, t = 0;
                        const ss: any = Math.min(w, h);
                        const maxAdj2: any = cnstVal1 * w / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal2 * h / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * ss / h;
                        const maxAdj4: any = cnstVal2 - q2;
                        const a4: any = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        const dx1: any = ss * a2 / cnstVal2;
                        const dx2: any = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        y1 = ss * a3 / cnstVal2;
                        const dy2: any = h * a4 / cnstVal2;
                        y2 = b - dy2;
                        y3 = (y2 + b) / 2;

                        d_val = "M" + l + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + y1 +
                            " L" + x1 + "," + y1 +
                            " L" + hc + "," + t +
                            " L" + x4 + "," + y1 +
                            " L" + x3 + "," + y1 +
                            " L" + x3 + "," + y2 +
                            " L" + r + "," + y2 +
                            " L" + r + "," + b +
                            " L" + l + "," + b +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 25000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 48123 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        let maxAdj2: any = undefined, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2, y3, y4, x1, x4, dx2, x2, x3;
                        vc = h / 2;
                        let hc: number = w / 2, r: number = w, b = h;
                        l = 0, t = 0;
                        const ss: any = Math.min(w, h);
                        const maxAdj2: any = cnstVal1 * h / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal1 * w / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * ss / wd2;
                        const maxAdj4: any = cnstVal2 - q2;
                        const a4: any = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        const dy2: any = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        x4 = r - x1;
                        const dx2: any = w * a4 / cnstVal3;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        d_val = "M" + l + "," + vc +
                            " L" + x1 + "," + y1 +
                            " L" + x1 + "," + y2 +
                            " L" + x2 + "," + y2 +
                            " L" + x2 + "," + t +
                            " L" + x3 + "," + t +
                            " L" + x3 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + y1 +
                            " L" + r + "," + vc +
                            " L" + x4 + "," + y4 +
                            " L" + x4 + "," + y3 +
                            " L" + x3 + "," + y3 +
                            " L" + x3 + "," + b +
                            " L" + x2 + "," + b +
                            " L" + x2 + "," + y3 +
                            " L" + x1 + "," + y3 +
                            " L" + x1 + "," + y4 +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "quadArrowCallout":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 18515 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 18515 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 18515 * slideFactor;
                        let sAdj4: any = undefined, adj4 = 48123 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = parseInt(sAdj4.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, r: number = w, b = h;
                        l = 0, t = 0;
                        const ss: any = Math.min(w, h);
                        a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx2, dx3, ah, dx1, dy1, x8, x2, x7, x3, x6, x4, x5, y8, y2, y7, y3, y6, y4, y5;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > cnstVal1) ? cnstVal1 : adj2;
                        maxAdj1 = a2 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const maxAdj3: any = cnstVal1 - a2;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const q2: any = a3 * 2;
                        const maxAdj4: any = cnstVal2 - q2;
                        const a4: any = (adj4 < a1) ? a1 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        const dx2: any = ss * a2 / cnstVal2;
                        const dx3: any = ss * a1 / cnstVal3;
                        const ah: any = ss * a3 / cnstVal2;
                        const dx1: any = w * a4 / cnstVal3;
                        dy1 = h * a4 / cnstVal3;
                        const x8: any = r - ah;
                        x2 = hc - dx1;
                        const x7: any = hc + dx1;
                        x3 = hc - dx2;
                        const x6: any = hc + dx2;
                        x4 = hc - dx3;
                        const x5: any = hc + dx3;
                        const y8: any = b - ah;
                        y2 = vc - dy1;
                        const y7: any = vc + dy1;
                        y3 = vc - dx2;
                        const y6: any = vc + dx2;
                        y4 = vc - dx3;
                        const y5: any = vc + dx3;
                        d_val = "M" + l + "," + vc +
                            " L" + ah + "," + y3 +
                            " L" + ah + "," + y4 +
                            " L" + x2 + "," + y4 +
                            " L" + x2 + "," + y2 +
                            " L" + x4 + "," + y2 +
                            " L" + x4 + "," + ah +
                            " L" + x3 + "," + ah +
                            " L" + hc + "," + t +
                            " L" + x6 + "," + ah +
                            " L" + x5 + "," + ah +
                            " L" + x5 + "," + y2 +
                            " L" + x7 + "," + y2 +
                            " L" + x7 + "," + y4 +
                            " L" + x8 + "," + y4 +
                            " L" + x8 + "," + y3 +
                            " L" + r + "," + vc +
                            " L" + x8 + "," + y6 +
                            " L" + x8 + "," + y5 +
                            " L" + x7 + "," + y5 +
                            " L" + x7 + "," + y7 +
                            " L" + x5 + "," + y7 +
                            " L" + x5 + "," + y8 +
                            " L" + x6 + "," + y8 +
                            " L" + hc + "," + b +
                            " L" + x3 + "," + y8 +
                            " L" + x4 + "," + y8 +
                            " L" + x4 + "," + y7 +
                            " L" + x2 + "," + y7 +
                            " L" + x2 + "," + y5 +
                            " L" + ah + "," + y5 +
                            " L" + ah + "," + y6 +
                            " z";

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedDownArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, wd2: number = w / 2, r = w, b = h;
                        l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        const ss: any = Math.min(w, h);
                        let maxAdj2: any = undefined, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, stAng, stAng2, swAng2, swAng3;

                        const maxAdj2: any = cnstVal1 * w / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        const th: any = ss * a1 / cnstVal2;
                        const aw: any = ss * a2 / cnstVal2;
                        const q1: any = (th + aw) / 4;
                        const wR: any = wd2 - q1;
                        const q7: any = wR * 2;
                        const q8: any = q7 * q7;
                        const q9: any = th * th;
                        const q10: any = q8 - q9;
                        const q11: any = Math.sqrt(q10);
                        const idy: any = q11 * h / q7;
                        const maxAdj3: any = cnstVal2 * idy / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const ah: any = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        const q2: any = h * h;
                        const q3: any = ah * ah;
                        const q4: any = q2 - q3;
                        const q5: any = Math.sqrt(q4);
                        const dx: any = q5 * wR / h;
                        const x5: any = wR + dx;
                        const x7: any = x3 + dx;
                        const q6: any = aw - th;
                        const dh: any = q6 / 2;
                        x4 = x5 - dh;
                        const x8: any = x7 + dh;
                        const aw2: any = aw / 2;
                        const x6: any = r - aw2;
                        y1 = b - ah;
                        const swAng: any = Math.atan(dx / ah);
                        const swAngDeg: any = swAng * 180 / Math.PI;
                        const mswAng: any = -swAngDeg;
                        const iy: any = b - idy;
                        const ix: any = (wR + x3) / 2;
                        const q12: any = th / 2;
                        const dang2: any = Math.atan(q12 / idy);
                        const dang2Deg: any = dang2 * 180 / Math.PI;
                        const stAng: any = c3d4 + swAngDeg;
                        const stAng2: any = c3d4 - dang2Deg;
                        const swAng2: any = dang2Deg - cd4;
                        const swAng3: any = cd4 + dang2Deg;
                        //const cX = x5 - Math.cos(stAng*Math.PI/180) * wR;
                        //const cY = y1 - Math.sin(stAng*Math.PI/180) * h;

                        d_val = "M" + x6 + "," + b +
                            " L" + x4 + "," + y1 +
                            " L" + x5 + "," + y1 +
                            PPTXShapeUtils.shapeArc(wR, h, wR, h, stAng, (stAng + mswAng), false).replace("M", "L") +
                            " L" + x3 + "," + t +
                            PPTXShapeUtils.shapeArc(x3, h, wR, h, c3d4, (c3d4 + swAngDeg), false).replace("M", "L") +
                            " L" + (x5 + th) + "," + y1 +
                            " L" + x8 + "," + y1 +
                            ` zM` + x3 + "," + t +
                            PPTXShapeUtils.shapeArc(x3, h, wR, h, stAng2, (stAng2 + swAng2), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(wR, h, wR, h, cd2, (cd2 + swAng3), false).replace("M", "L");

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedLeftArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, hd2: number = h / 2, r = w, b = h;
                        l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        const ss: any = Math.min(w, h);
                        let maxAdj2: any = undefined, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        const maxAdj2: any = cnstVal1 * h / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        const th: any = ss * a1 / cnstVal2;
                        const aw: any = ss * a2 / cnstVal2;
                        const q1: any = (th + aw) / 4;
                        const hR: any = hd2 - q1;
                        const q7: any = hR * 2;
                        const q8: any = q7 * q7;
                        const q9: any = th * th;
                        const q10: any = q8 - q9;
                        const q11: any = Math.sqrt(q10);
                        const iDx: any = q11 * w / q7;
                        const maxAdj3: any = cnstVal2 * iDx / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const ah: any = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        const q2: any = w * w;
                        const q3: any = ah * ah;
                        const q4: any = q2 - q3;
                        const q5: any = Math.sqrt(q4);
                        const dy: any = q5 * hR / w;
                        const y5: any = hR + dy;
                        const y7: any = y3 + dy;
                        const q6: any = aw - th;
                        const dh: any = q6 / 2;
                        y4 = y5 - dh;
                        const y8: any = y7 + dh;
                        const aw2: any = aw / 2;
                        const y6: any = b - aw2;
                        x1 = l + ah;
                        const swAng: any = Math.atan(dy / ah);
                        const mswAng: any = -swAng;
                        const ix: any = l + iDx;
                        const iy: any = (hR + y3) / 2;
                        const q12: any = th / 2;
                        const dang2: any = Math.atan(q12 / iDx);
                        const swAng2: any = dang2 - swAng;
                        const swAng3: any = swAng + dang2;
                        const stAng3: any = -dang2;
                        let swAngDg, swAng2Dg, swAng3Dg, stAng3dg;
                        const swAngDg: any = swAng * 180 / Math.PI;
                        const swAng2Dg: any = swAng2 * 180 / Math.PI;
                        const swAng3Dg: any = swAng3 * 180 / Math.PI;
                        const stAng3dg: any = stAng3 * 180 / Math.PI;

                        d_val = "M" + r + "," + y3 +
                            PPTXShapeUtils.shapeArc(l, hR, w, hR, 0, -cd4, false).replace("M", "L") +
                            " L" + l + "," + t +
                            PPTXShapeUtils.shapeArc(l, y3, w, hR, c3d4, (c3d4 + cd4), false).replace("M", "L") +
                            " L" + r + "," + y3 +
                            PPTXShapeUtils.shapeArc(l, y3, w, hR, 0, swAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y7 +
                            " L" + x1 + "," + y8 +
                            " L" + l + "," + y6 +
                            " L" + x1 + "," + y4 +
                            " L" + x1 + "," + y5 +
                            PPTXShapeUtils.shapeArc(l, hR, w, hR, swAngDg, (swAngDg + swAng2Dg), false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(l, hR, w, hR, 0, -cd4, false).replace("M", "L") +
                            PPTXShapeUtils.shapeArc(l, y3, w, hR, c3d4, (c3d4 + cd4), false).replace("M", "L");

                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedRightArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, hd2: number = h / 2, r = w, b = h;
                        l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        const ss: any = Math.min(w, h);
                        maxAdj2 = undefined, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy,
                            y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, stAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        const maxAdj2: any = cnstVal1 * h / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        const th: any = ss * a1 / cnstVal2;
                        const aw: any = ss * a2 / cnstVal2;
                        const q1: any = (th + aw) / 4;
                        const hR: any = hd2 - q1;
                        const q7: any = hR * 2;
                        const q8: any = q7 * q7;
                        const q9: any = th * th;
                        const q10: any = q8 - q9;
                        const q11: any = Math.sqrt(q10);
                        const iDx: any = q11 * w / q7;
                        const maxAdj3: any = cnstVal2 * iDx / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const ah: any = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        const q2: any = w * w;
                        const q3: any = ah * ah;
                        const q4: any = q2 - q3;
                        const q5: any = Math.sqrt(q4);
                        const dy: any = q5 * hR / w;
                        const y5: any = hR + dy;
                        const y7: any = y3 + dy;
                        const q6: any = aw - th;
                        const dh: any = q6 / 2;
                        y4 = y5 - dh;
                        const y8: any = y7 + dh;
                        const aw2: any = aw / 2;
                        const y6: any = b - aw2;
                        x1 = r - ah;
                        const swAng: any = Math.atan(dy / ah);
                        const stAng: any = Math.PI + 0 - swAng;
                        const mswAng: any = -swAng;
                        const ix: any = r - iDx;
                        const iy: any = (hR + y3) / 2;
                        const q12: any = th / 2;
                        const dang2: any = Math.atan(q12 / iDx);
                        const swAng2: any = dang2 - Math.PI / 2;
                        const swAng3: any = Math.PI / 2 + dang2;
                        const stAng3: any = Math.PI - dang2;

                        stAngDg, mswAngDg, swAngDg, swAng2dg;
                        const stAngDg: any = stAng * 180 / Math.PI;
                        const mswAngDg: any = mswAng * 180 / Math.PI;
                        const swAngDg: any = swAng * 180 / Math.PI;
                        const swAng2dg: any = swAng2 * 180 / Math.PI;

                        d_val = "M" + l + "," + hR +
                            PPTXShapeUtils.shapeArc(w, hR, w, hR, cd2, cd2 + mswAngDg, false).replace("M", "L") +
                            " L" + x1 + "," + y5 +
                            " L" + x1 + "," + y4 +
                            " L" + r + "," + y6 +
                            " L" + x1 + "," + y8 +
                            " L" + x1 + "," + y7 +
                            PPTXShapeUtils.shapeArc(w, y3, w, hR, stAngDg, stAngDg + swAngDg, false).replace("M", "L") +
                            " L" + l + "," + hR +
                            PPTXShapeUtils.shapeArc(w, hR, w, hR, cd2, cd2 + cd4, false).replace("M", "L") +
                            " L" + r + "," + th +
                            PPTXShapeUtils.shapeArc(w, y3, w, hR, c3d4, c3d4 + swAng2dg, false).replace("M", "L")
                        "";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "curvedUpArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 25000 * slideFactor;
                        let sAdj2: any = undefined, adj2 = 50000 * slideFactor;
                        let sAdj3: any = undefined, adj3 = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = parseInt(sAdj3.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, wd2: number = w / 2, r = w, b = h;
                        l = 0, t = 0, c3d4 = 270, cd2 = 180, cd4 = 90;
                        const ss: any = Math.min(w, h);
                        let maxAdj2: any = undefined, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, swAng2, mswAng2, stAng3, swAng3, stAng2;

                        const maxAdj2: any = cnstVal1 * w / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        const th: any = ss * a1 / cnstVal2;
                        const aw: any = ss * a2 / cnstVal2;
                        const q1: any = (th + aw) / 4;
                        const wR: any = wd2 - q1;
                        const q7: any = wR * 2;
                        const q8: any = q7 * q7;
                        const q9: any = th * th;
                        const q10: any = q8 - q9;
                        const q11: any = Math.sqrt(q10);
                        const idy: any = q11 * h / q7;
                        const maxAdj3: any = cnstVal2 * idy / ss;
                        const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        const ah: any = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        const q2: any = h * h;
                        const q3: any = ah * ah;
                        const q4: any = q2 - q3;
                        const q5: any = Math.sqrt(q4);
                        const dx: any = q5 * wR / h;
                        const x5: any = wR + dx;
                        const x7: any = x3 + dx;
                        const q6: any = aw - th;
                        const dh: any = q6 / 2;
                        x4 = x5 - dh;
                        const x8: any = x7 + dh;
                        const aw2: any = aw / 2;
                        const x6: any = r - aw2;
                        y1 = t + ah;
                        const swAng: any = Math.atan(dx / ah);
                        const mswAng: any = -swAng;
                        const iy: any = t + idy;
                        const ix: any = (wR + x3) / 2;
                        const q12: any = th / 2;
                        const dang2: any = Math.atan(q12 / idy);
                        const swAng2: any = dang2 - swAng;
                        const mswAng2: any = -swAng2;
                        const stAng3: any = Math.PI / 2 - swAng;
                        const swAng3: any = swAng + dang2;
                        const stAng2: any = Math.PI / 2 - dang2;

                        stAng2dg, swAng2dg, swAngDg, swAng2dg;
                        const stAng2dg: any = stAng2 * 180 / Math.PI;
                        const swAng2dg: any = swAng2 * 180 / Math.PI;
                        const stAng3dg: any = stAng3 * 180 / Math.PI;
                        const swAngDg: any = swAng * 180 / Math.PI;

                        d_val = //"M" + ix + "," +iy + 
                            PPTXShapeUtils.shapeArc(wR, 0, wR, h, stAng2dg, stAng2dg + swAng2dg, false) + //.replace("M","L") +
                            " L" + x5 + "," + y1 +
                            " L" + x4 + "," + y1 +
                            " L" + x6 + "," + t +
                            " L" + x8 + "," + y1 +
                            " L" + x7 + "," + y1 +
                            PPTXShapeUtils.shapeArc(x3, 0, wR, h, stAng3dg, stAng3dg + swAngDg, false).replace("M", "L") +
                            " L" + wR + "," + b +
                            PPTXShapeUtils.shapeArc(wR, 0, wR, h, cd4, cd2, false).replace("M", "L") +
                            " L" + th + "," + t +
                            PPTXShapeUtils.shapeArc(x3, 0, wR, h, cd2, cd4, false).replace("M", "L") +
                            "";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "mathDivide":
                    case "mathEqual":
                    case "mathMinus":
                    case "mathMultiply":
                    case "mathNotEqual":
                    case "mathPlus":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1;
                        let sAdj2: any = undefined, adj2;
                        let sAdj3: any = undefined, adj3;
                        if (shapAdjst_ary !== undefined) {
                            if (shapAdjst_ary.constructor === Array) {
                                for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4));
                                    } else if (sAdj_name == "adj2") {
                                        sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj2 = parseInt(sAdj2.substr(4));
                                    } else if (sAdj_name == "adj3") {
                                        const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        const adj3: any = parseInt(sAdj3.substr(4));
                                    }
                                }
                            } else {
                                sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                                adj1 = parseInt(sAdj1.substr(4));
                            }
                        }
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const cnstVal3: any = 200000 * slideFactor;
                        const hc: number = w / 2, vc: number = h / 2, hd2: number = h / 2;
                        if (shapType == "mathNotEqual") {
                            if (shapAdjst_ary === undefined) {
                                adj1 = 23520 * slideFactor;
                                adj2 = 110 * Math.PI / 180;
                                const adj3: any = 11760 * slideFactor;
                            } else {
                                adj1 = adj1 * slideFactor;
                                adj2 = (adj2 / 60000) * Math.PI / 180;
                                const adj3: any = adj3 * slideFactor;
                            }
                            a1 = undefined, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4,
                                cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37,
                                x3, dx27, x2, rx7, rx6, rx5, rx4, rx3, rx2, dx7, rxt, lxt, rx, lx,
                                dy3, dy4, ry, ly, dlx, drx, dly, dry, xC1, xC2, yC1, yC2, yC3, yC4;
                            const angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
                            const cnstVal4 = 73490 * slideFactor;
                            //const cd4 = 90;
                            const a1: any = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            const crAng: any = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
                            const a2a1: any = a1 * 2;
                            const maxAdj3: any = cnstVal2 - a2a1;
                            const a3: any = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                            dy1 = h * a1 / cnstVal2;
                            const dy2: any = h * a3 / cnstVal3;
                            const dx1: any = w * cnstVal4 / cnstVal3;
                            x1 = hc - dx1;
                            const x8: any = hc + dx1;
                            y2 = vc - dy2;
                            y3 = vc + dy2;
                            y1 = y2 - dy1;
                            y4 = y3 + dy1;
                            const cadj2: any = crAng - Math.PI / 2;
                            const xadj2: any = hd2 * Math.tan(cadj2);
                            const len: any = Math.sqrt(xadj2 * xadj2 + hd2 * hd2);
                            const bhw: any = len * dy1 / hd2;
                            const bhw2: any = bhw / 2;
                            const x7: any = hc + xadj2 - bhw2;
                            const dx67: any = xadj2 * y1 / hd2;
                            const x6: any = x7 - dx67;
                            const dx57: any = xadj2 * y2 / hd2;
                            const x5: any = x7 - dx57;
                            const dx47: any = xadj2 * y3 / hd2;
                            x4 = x7 - dx47;
                            const dx37: any = xadj2 * y4 / hd2;
                            x3 = x7 - dx37;
                            const dx27: any = xadj2 * 2;
                            x2 = x7 - dx27;
                            const rx7: any = x7 + bhw;
                            const rx6: any = x6 + bhw;
                            const rx5: any = x5 + bhw;
                            const rx4: any = x4 + bhw;
                            const rx3: any = x3 + bhw;
                            const rx2: any = x2 + bhw;
                            const dx7: any = dy1 * hd2 / len;
                            const rxt: any = x7 + dx7;
                            const lxt: any = rx7 - dx7;
                            const rx: any = (cadj2 > 0) ? rxt : rx7;
                            const lx: any = (cadj2 > 0) ? x7 : lxt;
                            const dy3: any = dy1 * xadj2 / len;
                            const dy4: any = -dy3;
                            const ry: any = (cadj2 > 0) ? dy3 : 0;
                            const ly: any = (cadj2 > 0) ? 0 : dy4;
                            const dlx: any = w - rx;
                            const drx: any = w - lx;
                            const dly: any = h - ry;
                            const dry: any = h - ly;
                            const xC1: any = (rx + lx) / 2;
                            const xC2: any = (drx + dlx) / 2;
                            const yC1: any = (ry + ly) / 2;
                            const yC2: any = (y1 + y2) / 2;
                            const yC3: any = (y3 + y4) / 2;
                            const yC4: any = (dry + dly) / 2;

                            dVal = "M" + x1 + "," + y1 +
                                " L" + x6 + "," + y1 +
                                " L" + lx + "," + ly +
                                " L" + rx + "," + ry +
                                " L" + rx6 + "," + y1 +
                                " L" + x8 + "," + y1 +
                                " L" + x8 + "," + y2 +
                                " L" + rx5 + "," + y2 +
                                " L" + rx4 + "," + y3 +
                                " L" + x8 + "," + y3 +
                                " L" + x8 + "," + y4 +
                                " L" + rx3 + "," + y4 +
                                " L" + drx + "," + dry +
                                " L" + dlx + "," + dly +
                                " L" + x3 + "," + y4 +
                                " L" + x1 + "," + y4 +
                                " L" + x1 + "," + y3 +
                                " L" + x4 + "," + y3 +
                                " L" + x5 + "," + y2 +
                                " L" + x1 + "," + y2 +
                                " z";
                        } else if (shapType == "mathDivide") {
                            if (shapAdjst_ary === undefined) {
                                adj1 = 23520 * slideFactor;
                                adj2 = 5880 * slideFactor;
                                const adj3: any = 11760 * slideFactor;
                            } else {
                                adj1 = adj1 * slideFactor;
                                adj2 = adj2 * slideFactor;
                                const adj3: any = adj3 * slideFactor;
                            }
                            a1 = undefined, ma1, ma3h, ma3w, maxAdj3, a3, m4a3, maxAdj2, a2, dy1, yg, rad, dx1,
                                y3, y4, a, y2, y1, y5, x1, x3, x2;
const cnstVal4: any = 1000 * slideFactor;
const cnstVal5: any = 36745 * slideFactor;
                            const cnstVal6 = 73490 * slideFactor;
                            const a1: any = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            const ma1: any = -a1;
                            const ma3h: any = (cnstVal6 + ma1) / 4;
                            const ma3w: any = cnstVal5 * w / h;
                            const maxAdj3: any = (ma3h < ma3w) ? ma3h : ma3w;
                            const a3: any = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                            const m4a3: any = -4 * a3;
                            const maxAdj2: any = cnstVal6 + m4a3 - a1;
                            const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                            dy1 = h * a1 / cnstVal3;
                            const yg: any = h * a2 / cnstVal2;
                            const rad: any = h * a3 / cnstVal2;
                            const dx1: any = w * cnstVal6 / cnstVal3;
                            y3 = vc - dy1;
                            y4 = vc + dy1;
                            const a: any = yg + rad;
                            y2 = y3 - a;
                            y1 = y2 - rad;
                            const y5: any = h - y1;
                            x1 = hc - dx1;
                            x3 = hc + dx1;
                            x2 = hc - rad;
                            const cd4 = 90, c3d4 = 270;
                            const cX1 = hc - Math.cos(c3d4 * Math.PI / 180) * rad;
                            const cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
                            const cX2 = hc - Math.cos(Math.PI / 2) * rad;
                            const cY2 = y5 - Math.sin(Math.PI / 2) * rad;
                            dVal = "M" + hc + "," + y1 +
                                PPTXShapeUtils.shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
                                ` z M` + hc + "," + y5 +
                                PPTXShapeUtils.shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
                                ` z M` + x1 + "," + y3 +
                                " L" + x3 + "," + y3 +
                                " L" + x3 + "," + y4 +
                                " L" + x1 + "," + y4 +
                                " z";
                        } else if (shapType == "mathEqual") {
                            const dVal: any = PPTXMathShapes.genMathEqual(w, h, node, slideFactor);
                        } else if (shapType == "mathMinus") {
                            const dVal: any = PPTXMathShapes.genMathMinus(w, h, node, slideFactor);
                        } else if (shapType == "mathMultiply") {
                            const dVal: any = PPTXMathShapes.genMathMultiply(w, h, node, slideFactor);
                        } else if (shapType == "mathPlus") {
                            const dVal: any = PPTXMathShapes.genMathPlus(w, h, node, slideFactor);
                        }
                        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        //console.log(shapType);
                        break;
                    case "can":
                    case "flowChartMagneticDisk":
                    case "flowChartMagneticDrum":
                        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const adj: any = 25000 * slideFactor;
                        const cnstVal1: any = 50000 * slideFactor;
                        cnstVal2 = 200000 * slideFactor;
                        if (shapAdjst !== undefined) {
                            const adj: any = parseInt(shapAdjst.substr(4)) * slideFactor;
                        }
                        const ss: any = Math.min(w, h);
                        let maxAdj: any = undefined, a, y1, y2, y3;
                        if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
                            const adj: any = 50000 * slideFactor;
                        }
                        const maxAdj: any = cnstVal1 * h / ss;
                        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        y1 = ss * a / cnstVal2;
                        y2 = y1 + y1;
                        y3 = h - y1;
                        cd2 = 180, wd2 = w / 2;

                        const tranglRott: any = "";
                        if (shapType == "flowChartMagneticDrum") {
                            const tranglRott: any = "transform='rotate(90 " + w / 2 + "," + h / 2 + ")'";
                        }
                        dVal = PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, 0, cd2, false) +
                            PPTXShapeUtils.shapeArc(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace("M", "L") +
                            " L" + w + "," + y3 +
                            PPTXShapeUtils.shapeArc(wd2, y3, wd2, y1, 0, cd2, false).replace("M", "L") +
                            " L" + 0 + "," + y1;

                        result += "<path " + tranglRott + " d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "swooshArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr: any = slideFactor;
                        let sAdj1: any = undefined, adj1 = 25000 * refr;
                        let sAdj2: any = undefined, adj2 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        const cnstVal1: any = 1 * refr;
                        cnstVal2 = 70000 * refr;
                        const cnstVal3: any = 75000 * refr;
                        const cnstVal4: any = 100000 * refr;
                        const ss: any = Math.min(w, h);
                        const ssd8: any = ss / 8;
                        const hd6: any = h / 6;

                        const a1: any = (adj1 < cnstVal1) ? cnstVal1 : (adj1 > cnstVal3) ? cnstVal3 : adj1;
                        const maxAdj2: any = cnstVal2 * w / ss;
                        const a2: any = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        const ad1: any = h * a1 / cnstVal4;
                        const ad2: any = ss * a2 / cnstVal4;
                        const xB: any = w - ad2;
                        const yB: any = ssd8;
                        const alfa: any = (Math.PI / 2) / 14;
                        const dx0: any = ssd8 * Math.tan(alfa);
                        const xC: any = xB - dx0;
                        const dx1: any = ad1 * Math.tan(alfa);
                        const yF: any = yB + ad1;
                        const xF: any = xB + dx1;
                        const xE: any = xF + dx0;
                        const yE: any = yF + ssd8;
                        const dy2: any = yE - 0;
                        const dy22: any = dy2 / 2;
                        const dy3: any = h / 20;
                        const yD: any = dy22 - dy3;
                        const dy4: any = hd6;
                        const yP1: any = hd6 + dy4;
                        const xP1: any = w / 6;
                        const dy5: any = hd6 / 2;
                        const yP2: any = yF + dy5;
                        const xP2: any = w / 4;

                        dVal = "M" + 0 + "," + h +
                            " Q" + xP1 + "," + yP1 + " " + xB + "," + yB +
                            " L" + xC + "," + 0 +
                            " L" + w + "," + yD +
                            " L" + xE + "," + yE +
                            " L" + xF + "," + yF +
                            " Q" + xP2 + "," + yP2 + " " + 0 + "," + h +
                            " z";

                        result += "<path d='" + dVal + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "circularArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 12500 * slideFactor;
                        let sAdj2: any = undefined, adj2 = (1142319 / 60000) * Math.PI / 180;
                        let sAdj3: any = undefined, adj3 = (20457681 / 60000) * Math.PI / 180;
                        let sAdj4: any = undefined, adj4 = (10800000 / 60000) * Math.PI / 180;
                        let sAdj5: any = undefined, adj5 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    const sAdj5: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj5: any = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, r: number = w, b = h;
                        l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        const ss: any = Math.min(w, h);
                        const cnstVal1: any = 25000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const rdAngVal1: any = (1 / 60000) * Math.PI / 180;
                        const rdAngVal2: any = (21599999 / 60000) * Math.PI / 180;
                        const rdAngVal3: any = 2 * Math.PI;

                        const a5: any = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const enAng: any = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        const stAng: any = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4; //////////////////////////////////////////
                        const th: any = ss * a1 / cnstVal2;
                        const thh: any = ss * a5 / cnstVal2;
                        const th2: any = th / 2;
                        const rw1: any = wd2 + th2 - thh;
                        const rh1: any = hd2 + th2 - thh;
                        const rw2: any = rw1 - th;
                        const rh2: any = rh1 - th;
                        const rw3: any = rw2 + th2;
                        const rh3: any = rh2 + th2;
                        const wtH: any = rw3 * Math.sin(enAng);
                        const htH: any = rh3 * Math.cos(enAng);

                        //dxH = rw3*Math.cos(Math.atan(wtH/htH));
                        //dyH = rh3*Math.sin(Math.atan(wtH/htH));
                        const dxH: any = rw3 * Math.cos(Math.atan2(wtH, htH));
                        const dyH: any = rh3 * Math.sin(Math.atan2(wtH, htH));

                        const xH: any = hc + dxH;
                        const yH: any = vc + dyH;
                        const rI: any = (rw2 < rh2) ? rw2 : rh2;
                        const u1: any = dxH * dxH;
                        const u2: any = dyH * dyH;
                        const u3: any = rI * rI;
                        const u4: any = u1 - u3;
                        const u5: any = u2 - u3;
                        const u6: any = u4 * u5 / u1;
                        const u7: any = u6 / u2;
                        const u8: any = 1 - u7;
                        const u9: any = Math.sqrt(u8);
                        const u10: any = u4 / dxH;
                        const u11: any = u10 / dyH;
                        const u12: any = (1 + u9) / u11;

                        //u13 = Math.atan(u12/1);
                        const u13: any = Math.atan2(u12, 1);

                        const u14: any = u13 + rdAngVal3;
                        const u15: any = (u13 > 0) ? u13 : u14;
                        const u16: any = u15 - enAng;
                        const u17: any = u16 + rdAngVal3;
                        const u18: any = (u16 > 0) ? u16 : u17;
                        const u19: any = u18 - cd2;
                        const u20: any = u18 - rdAngVal3;
                        const u21: any = (u19 > 0) ? u20 : u18;
                        const maxAng: any = Math.abs(u21);
                        const aAng: any = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
                        const ptAng: any = enAng + aAng;
                        const wtA: any = rw3 * Math.sin(ptAng);
                        const htA: any = rh3 * Math.cos(ptAng);
                        //dxA = rw3*Math.cos(Math.atan(wtA/htA));
                        //dyA = rh3*Math.sin(Math.atan(wtA/htA));
                        const dxA: any = rw3 * Math.cos(Math.atan2(wtA, htA));
                        const dyA: any = rh3 * Math.sin(Math.atan2(wtA, htA));

                        const xA: any = hc + dxA;
                        const yA: any = vc + dyA;
                        const wtE: any = rw1 * Math.sin(stAng);
                        const htE: any = rh1 * Math.cos(stAng);

                        //dxE = rw1*Math.cos(Math.atan(wtE/htE));
                        //dyE = rh1*Math.sin(Math.atan(wtE/htE));
                        const dxE: any = rw1 * Math.cos(Math.atan2(wtE, htE));
                        const dyE: any = rh1 * Math.sin(Math.atan2(wtE, htE));

                        const xE: any = hc + dxE;
                        const yE: any = vc + dyE;
                        const dxG: any = thh * Math.cos(ptAng);
                        const dyG: any = thh * Math.sin(ptAng);
                        const xG: any = xH + dxG;
                        const yG: any = yH + dyG;
                        const dxB: any = thh * Math.cos(ptAng);
                        const dyB: any = thh * Math.sin(ptAng);
                        const xB: any = xH - dxB;
                        const yB: any = yH - dyB;
                        const sx1: any = xB - hc;
                        const sy1: any = yB - vc;
                        const sx2: any = xG - hc;
                        const sy2: any = yG - vc;
                        const rO: any = (rw1 < rh1) ? rw1 : rh1;
                        const x1O: any = sx1 * rO / rw1;
                        const y1O: any = sy1 * rO / rh1;
                        const x2O: any = sx2 * rO / rw1;
                        const y2O: any = sy2 * rO / rh1;
                        const dxO: any = x2O - x1O;
                        const dyO: any = y2O - y1O;
                        const dO: any = Math.sqrt(dxO * dxO + dyO * dyO);
                        const q1: any = x1O * y2O;
                        const q2: any = x2O * y1O;
                        const DO: any = q1 - q2;
                        const q3: any = rO * rO;
                        const q4: any = dO * dO;
                        const q5: any = q3 * q4;
                        const q6: any = DO * DO;
                        const q7: any = q5 - q6;
                        const q8: any = (q7 > 0) ? q7 : 0;
                        const sdelO: any = Math.sqrt(q8);
                        const ndyO: any = dyO * -1;
                        const sdyO: any = (ndyO > 0) ? -1 : 1;
                        const q9: any = sdyO * dxO;
                        const q10: any = q9 * sdelO;
                        const q11: any = DO * dyO;
                        const dxF1: any = (q11 + q10) / q4;
                        const q12: any = q11 - q10;
                        const dxF2: any = q12 / q4;
                        const adyO: any = Math.abs(dyO);
                        const q13: any = adyO * sdelO;
                        const q14: any = DO * dxO / -1;
                        const dyF1: any = (q14 + q13) / q4;
                        const q15: any = q14 - q13;
                        const dyF2: any = q15 / q4;
                        const q16: any = x2O - dxF1;
                        const q17: any = x2O - dxF2;
                        const q18: any = y2O - dyF1;
                        const q19: any = y2O - dyF2;
                        const q20: any = Math.sqrt(q16 * q16 + q18 * q18);
                        const q21: any = Math.sqrt(q17 * q17 + q19 * q19);
                        const q22: any = q21 - q20;
                        const dxF: any = (q22 > 0) ? dxF1 : dxF2;
                        const dyF: any = (q22 > 0) ? dyF1 : dyF2;
                        const sdxF: any = dxF * rw1 / rO;
                        const sdyF: any = dyF * rh1 / rO;
                        const xF: any = hc + sdxF;
                        const yF: any = vc + sdyF;
                        const x1I: any = sx1 * rI / rw2;
                        const y1I: any = sy1 * rI / rh2;
                        const x2I: any = sx2 * rI / rw2;
                        const y2I: any = sy2 * rI / rh2;
                        const dxI: any = x2I - x1I;
                        const dyI: any = y2I - y1I;
                        const dI: any = Math.sqrt(dxI * dxI + dyI * dyI);
                        const v1: any = x1I * y2I;
                        const v2: any = x2I * y1I;
                        const DI: any = v1 - v2;
                        const v3: any = rI * rI;
                        const v4: any = dI * dI;
                        const v5: any = v3 * v4;
                        const v6: any = DI * DI;
                        const v7: any = v5 - v6;
                        const v8: any = (v7 > 0) ? v7 : 0;
                        const sdelI: any = Math.sqrt(v8);
                        const v9: any = sdyO * dxI;
                        const v10: any = v9 * sdelI;
                        const v11: any = DI * dyI;
                        const dxC1: any = (v11 + v10) / v4;
                        const v12: any = v11 - v10;
                        const dxC2: any = v12 / v4;
                        const adyI: any = Math.abs(dyI);
                        const v13: any = adyI * sdelI;
                        const v14: any = DI * dxI / -1;
                        const dyC1: any = (v14 + v13) / v4;
                        const v15: any = v14 - v13;
                        const dyC2: any = v15 / v4;
                        const v16: any = x1I - dxC1;
                        const v17: any = x1I - dxC2;
                        const v18: any = y1I - dyC1;
                        const v19: any = y1I - dyC2;
                        const v20: any = Math.sqrt(v16 * v16 + v18 * v18);
                        const v21: any = Math.sqrt(v17 * v17 + v19 * v19);
                        const v22: any = v21 - v20;
                        const dxC: any = (v22 > 0) ? dxC1 : dxC2;
                        const dyC: any = (v22 > 0) ? dyC1 : dyC2;
                        const sdxC: any = dxC * rw2 / rI;
                        const sdyC: any = dyC * rh2 / rI;
                        const xC: any = hc + sdxC;
                        const yC: any = vc + sdyC;

                        //ist0 = Math.atan(sdyC/sdxC);
                        const ist0: any = Math.atan2(sdyC, sdxC);

                        const ist1: any = ist0 + rdAngVal3;
                        const istAng: any = (ist0 > 0) ? ist0 : ist1;
                        const isw1: any = stAng - istAng;
                        const isw2: any = isw1 - rdAngVal3;
                        const iswAng: any = (isw1 > 0) ? isw2 : isw1;
                        const p1: any = xF - xC;
                        const p2: any = yF - yC;
                        const p3: any = Math.sqrt(p1 * p1 + p2 * p2);
                        const p4: any = p3 / 2;
                        const p5: any = p4 - thh;
                        const xGp: any = (p5 > 0) ? xF : xG;
                        const yGp: any = (p5 > 0) ? yF : yG;
                        const xBp: any = (p5 > 0) ? xC : xB;
                        const yBp: any = (p5 > 0) ? yC : yB;

                        //en0 = Math.atan(sdyF/sdxF);
                        const en0: any = Math.atan2(sdyF, sdxF);

                        const en1: any = en0 + rdAngVal3;
                        const en2: any = (en0 > 0) ? en0 : en1;
                        const sw0: any = en2 - stAng;
                        const sw1: any = sw0 + rdAngVal3;
                        const swAng: any = (sw0 > 0) ? sw0 : sw1;

                        strtAng = stAng * 180 / Math.PI
                        const endAng: any = strtAng + (swAng * 180 / Math.PI);
                        const stiAng: any = istAng * 180 / Math.PI;
                        const swiAng: any = iswAng * 180 / Math.PI;
                        const ediAng: any = stiAng + swiAng;

                        d_val = PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false) +
                            " L" + xGp + "," + yGp +
                            " L" + xA + "," + yA +
                            " L" + xBp + "," + yBp +
                            " L" + xC + "," + yC +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftCircularArrow":
                        shapAdjst_ary: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1: any = undefined, adj1 = 12500 * slideFactor;
                        let sAdj2: any = undefined, adj2 = (-1142319 / 60000) * Math.PI / 180;
                        let sAdj3: any = undefined, adj3 = (1142319 / 60000) * Math.PI / 180;
                        let sAdj4: any = undefined, adj4 = (10800000 / 60000) * Math.PI / 180;
                        let sAdj5: any = undefined, adj5 = 12500 * slideFactor;
                        if (shapAdjst_ary !== undefined) {
                            for (let i = 0; i < shapAdjst_ary.length; i++) {
const sAdj_name: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * slideFactor;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    const sAdj3: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj3: any = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    const sAdj4: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj4: any = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    const sAdj5: any = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    const adj5: any = parseInt(sAdj5.substr(4)) * slideFactor;
                                }
                            }
                        }
                        vc = h / 2;
                        let hc: number = w / 2, r: number = w, b = h;
                        l = 0, t = 0, wd2 = w / 2, hd2 = h / 2;
                        const ss: any = Math.min(w, h);
                        const cnstVal1: any = 25000 * slideFactor;
                        cnstVal2 = 100000 * slideFactor;
                        const rdAngVal1: any = (1 / 60000) * Math.PI / 180;
                        const rdAngVal2: any = (21599999 / 60000) * Math.PI / 180;
                        const rdAngVal3: any = 2 * Math.PI;

                        const a5: any = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        const a1: any = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        const enAng: any = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        const stAng: any = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                        const th: any = ss * a1 / cnstVal2;
                        const thh: any = ss * a5 / cnstVal2;
                        const th2: any = th / 2;
                        const rw1: any = wd2 + th2 - thh;
                        const rh1: any = hd2 + th2 - thh;
                        const rw2: any = rw1 - th;
                        const rh2: any = rh1 - th;
                        const rw3: any = rw2 + th2;
                        const rh3: any = rh2 + th2;
                        const wtH: any = rw3 * Math.sin(enAng);
                        const htH: any = rh3 * Math.cos(enAng);
                        const dxH: any = rw3 * Math.cos(Math.atan2(wtH, htH));
                        const dyH: any = rh3 * Math.sin(Math.atan2(wtH, htH));
                        const xH: any = hc + dxH;
                        const yH: any = vc + dyH;
                        const rI: any = (rw2 < rh2) ? rw2 : rh2;
                        const u1: any = dxH * dxH;
                        const u2: any = dyH * dyH;
                        const u3: any = rI * rI;
                        const u4: any = u1 - u3;
                        const u5: any = u2 - u3;
                        const u6: any = u4 * u5 / u1;
                        const u7: any = u6 / u2;
                        const u8: any = 1 - u7;
                        const u9: any = Math.sqrt(u8);
                        const u10: any = u4 / dxH;
                        const u11: any = u10 / dyH;
                        const u12: any = (1 + u9) / u11;
                        const u13: any = Math.atan2(u12, 1);
                        const u14: any = u13 + rdAngVal3;
                        const u15: any = (u13 > 0) ? u13 : u14;
                        const u16: any = u15 - enAng;
                        const u17: any = u16 + rdAngVal3;
                        const u18: any = (u16 > 0) ? u16 : u17;
                        const u19: any = u18 - cd2;
                        const u20: any = u18 - rdAngVal3;
                        const u21: any = (u19 > 0) ? u20 : u18;
                        const u22: any = Math.abs(u21);
                        const minAng: any = u22 * -1;
                        const u23: any = Math.abs(adj2);
                        const a2: any = u23 * -1;
                        const aAng: any = (a2 < minAng) ? minAng : (a2 > 0) ? 0 : a2;
                        const ptAng: any = enAng + aAng;
                        const wtA: any = rw3 * Math.sin(ptAng);
                        const htA: any = rh3 * Math.cos(ptAng);
                        const dxA: any = rw3 * Math.cos(Math.atan2(wtA, htA));
                        const dyA: any = rh3 * Math.sin(Math.atan2(wtA, htA));
                        const xA: any = hc + dxA;
                        const yA: any = vc + dyA;
                        const wtE: any = rw1 * Math.sin(stAng);
                        const htE: any = rh1 * Math.cos(stAng);
                        const dxE: any = rw1 * Math.cos(Math.atan2(wtE, htE));
                        const dyE: any = rh1 * Math.sin(Math.atan2(wtE, htE));
                        const xE: any = hc + dxE;
                        const yE: any = vc + dyE;
                        const wtD: any = rw2 * Math.sin(stAng);
                        const htD: any = rh2 * Math.cos(stAng);
                        const dxD: any = rw2 * Math.cos(Math.atan2(wtD, htD));
                        const dyD: any = rh2 * Math.sin(Math.atan2(wtD, htD));
                        const xD: any = hc + dxD;
                        const yD: any = vc + dyD;
                        const dxG: any = thh * Math.cos(ptAng);
                        const dyG: any = thh * Math.sin(ptAng);
                        const xG: any = xH + dxG;
                        const yG: any = yH + dyG;
                        const dxB: any = thh * Math.cos(ptAng);
                        const dyB: any = thh * Math.sin(ptAng);
                        const xB: any = xH - dxB;
                        const yB: any = yH - dyB;
                        const sx1: any = xB - hc;
                        const sy1: any = yB - vc;
                        const sx2: any = xG - hc;
                        const sy2: any = yG - vc;
                        const rO: any = (rw1 < rh1) ? rw1 : rh1;
                        const x1O: any = sx1 * rO / rw1;
                        const y1O: any = sy1 * rO / rh1;
                        const x2O: any = sx2 * rO / rw1;
                        const y2O: any = sy2 * rO / rh1;
                        const dxO: any = x2O - x1O;
                        const dyO: any = y2O - y1O;
                        const dO: any = Math.sqrt(dxO * dxO + dyO * dyO);
                        const q1: any = x1O * y2O;
                        const q2: any = x2O * y1O;
                        const DO: any = q1 - q2;
                        const q3: any = rO * rO;
                        const q4: any = dO * dO;
                        const q5: any = q3 * q4;
                        const q6: any = DO * DO;
                        const q7: any = q5 - q6;
                        const q8: any = (q7 > 0) ? q7 : 0;
                        const sdelO: any = Math.sqrt(q8);
                        const ndyO: any = dyO * -1;
                        const sdyO: any = (ndyO > 0) ? -1 : 1;
                        const q9: any = sdyO * dxO;
                        const q10: any = q9 * sdelO;
                        const q11: any = DO * dyO;
                        const dxF1: any = (q11 + q10) / q4;
                        const q12: any = q11 - q10;
                        const dxF2: any = q12 / q4;
                        const adyO: any = Math.abs(dyO);
                        const q13: any = adyO * sdelO;
                        const q14: any = DO * dxO / -1;
                        const dyF1: any = (q14 + q13) / q4;
                        const q15: any = q14 - q13;
                        const dyF2: any = q15 / q4;
                        const q16: any = x2O - dxF1;
                        const q17: any = x2O - dxF2;
                        const q18: any = y2O - dyF1;
                        const q19: any = y2O - dyF2;
                        const q20: any = Math.sqrt(q16 * q16 + q18 * q18);
                        const q21: any = Math.sqrt(q17 * q17 + q19 * q19);
                        const q22: any = q21 - q20;
                        const dxF: any = (q22 > 0) ? dxF1 : dxF2;
                        const dyF: any = (q22 > 0) ? dyF1 : dyF2;
                        const sdxF: any = dxF * rw1 / rO;
                        const sdyF: any = dyF * rh1 / rO;
                        const xF: any = hc + sdxF;
                        const yF: any = vc + sdyF;
                        const x1I: any = sx1 * rI / rw2;
                        const y1I: any = sy1 * rI / rh2;
                        const x2I: any = sx2 * rI / rw2;
                        const y2I: any = sy2 * rI / rh2;
                        const dxI: any = x2I - x1I;
                        const dyI: any = y2I - y1I;
                        const dI: any = Math.sqrt(dxI * dxI + dyI * dyI);
                        const v1: any = x1I * y2I;
                        const v2: any = x2I * y1I;
                        const DI: any = v1 - v2;
                        const v3: any = rI * rI;
                        const v4: any = dI * dI;
                        const v5: any = v3 * v4;
                        const v6: any = DI * DI;
                        const v7: any = v5 - v6;
                        const v8: any = (v7 > 0) ? v7 : 0;
                        const sdelI: any = Math.sqrt(v8);
                        const v9: any = sdyO * dxI;
                        const v10: any = v9 * sdelI;
                        const v11: any = DI * dyI;
                        const dxC1: any = (v11 + v10) / v4;
                        const v12: any = v11 - v10;
                        const dxC2: any = v12 / v4;
                        const adyI: any = Math.abs(dyI);
                        const v13: any = adyI * sdelI;
                        const v14: any = DI * dxI / -1;
                        const dyC1: any = (v14 + v13) / v4;
                        const v15: any = v14 - v13;
                        const dyC2: any = v15 / v4;
                        const v16: any = x1I - dxC1;
                        const v17: any = x1I - dxC2;
                        const v18: any = y1I - dyC1;
                        const v19: any = y1I - dyC2;
                        const v20: any = Math.sqrt(v16 * v16 + v18 * v18);
                        const v21: any = Math.sqrt(v17 * v17 + v19 * v19);
                        const v22: any = v21 - v20;
                        const dxC: any = (v22 > 0) ? dxC1 : dxC2;
                        const dyC: any = (v22 > 0) ? dyC1 : dyC2;
                        const sdxC: any = dxC * rw2 / rI;
                        const sdyC: any = dyC * rh2 / rI;
                        const xC: any = hc + sdxC;
                        const yC: any = vc + sdyC;
                        const ist0: any = Math.atan2(sdyC, sdxC);
                        const ist1: any = ist0 + rdAngVal3;
                        const istAng0: any = (ist0 > 0) ? ist0 : ist1;
                        const isw1: any = stAng - istAng0;
                        const isw2: any = isw1 + rdAngVal3;
                        const iswAng0: any = (isw1 > 0) ? isw1 : isw2;
                        const istAng: any = istAng0 + iswAng0;
                        const iswAng: any = -iswAng0;
                        const p1: any = xF - xC;
                        const p2: any = yF - yC;
                        const p3: any = Math.sqrt(p1 * p1 + p2 * p2);
                        const p4: any = p3 / 2;
                        const p5: any = p4 - thh;
                        const xGp: any = (p5 > 0) ? xF : xG;
                        const yGp: any = (p5 > 0) ? yF : yG;
                        const xBp: any = (p5 > 0) ? xC : xB;
                        const yBp: any = (p5 > 0) ? yC : yB;
                        const en0: any = Math.atan2(sdyF, sdxF);
                        const en1: any = en0 + rdAngVal3;
                        const en2: any = (en0 > 0) ? en0 : en1;
                        const sw0: any = en2 - stAng;
                        const sw1: any = sw0 - rdAngVal3;
                        const swAng: any = (sw0 > 0) ? sw1 : sw0;
                        const stAng0: any = stAng + swAng;

                        const strtAng: any = stAng0 * 180 / Math.PI;
                        const endAng: any = stAng * 180 / Math.PI;
                        const stiAng: any = istAng * 180 / Math.PI;
                        const swiAng: any = iswAng * 180 / Math.PI;
                        const ediAng: any = stiAng + swiAng;

                        d_val = "M" + xE + "," + yE +
                            " L" + xD + "," + yD +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L") +
                            " L" + xBp + "," + yBp +
                            " L" + xA + "," + yA +
                            " L" + xGp + "," + yGp +
                            " L" + xF + "," + yF +
                            PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L") +
                            " z";
                        fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor, imgFillFlg, grndFillFlg, shpId);
                        result += "<path d='" + d_val + "' fill='" + fillAttr +
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";

                        break;
                    case "leftRightCircularArrow":
                    case "chartPlus":
                    case "chartStar":
                    case "chartX":
                    case "cornerTabs":
                    case "flowChartOfflineStorage":
                    case "folderCorner":
                    case "funnel":
                    case "lineInv":
                    case "nonIsoscelesTrapezoid":
                    case "plaqueTabs":
                    case "squareTabs":
                    case "upDownArrowCallout":
                        break;
                    case undefined:
                    default:
                }

                result += "</svg>";

                result += "<div class='block " + PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content
                    " " + PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    " z-index: " + order + `;transform: rotate(` + ((txtRotate !== undefined) ? txtRotate : 0) + `deg);'>`;

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        const type: any = "shape";
                    }
                    result += PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable); //type='shape'
                }
                result += "</div>";
            } else if (custShapType !== undefined) {
                //custGeom here - Amir ///////////////////////////////////////////////////////
                //http://officeopenxml.com/drwSp-custGeom.php
                const pathLstNode = PPTXUtils.getTextByPathList(custShapType, ["a:pathLst"]);
                const pathNodes = PPTXUtils.getTextByPathList(pathLstNode, ["a:path"]);
                //const pathNode = PPTXUtils.getTextByPathList(pathLstNode, ["a:path", "attrs"]);
                const maxX = parseInt(pathNodes["attrs"]["w"]);// * slideFactor;
                const maxY = parseInt(pathNodes["attrs"]["h"]);// * slideFactor;
                const cX = (1 / maxX) * w;
                const cY = (1 / maxY) * h;
                //console.log("w = "+w+"\nh = "+h+"\nmaxX = "+maxX +"\nmaxY = " + maxY);
                //cheke if it is close shape

                //console.log("custShapType : ", custShapType, ", pathLstNode: ", pathLstNode, ", node: ", node);//, ", y:", y, ", w:", w, ", h:", h);

                let moveToNode = PPTXUtils.getTextByPathList(pathNodes, ["a:moveTo"]);
                const total_shapes = moveToNode.length;

                const lnToNodes = pathNodes["a:lnTo"]; //total a:pt : 1
                let cubicBezToNodes = pathNodes["a:cubicBezTo"]; //total a:pt : 3
                const arcToNodes = pathNodes["a:arcTo"]; //total a:pt : 0?1? ; attrs: ~4 ()
                let closeNode = PPTXUtils.getTextByPathList(pathNodes, ["a:close"]); //total a:pt : 0
                //quadBezTo //total a:pt : 2
                //console.log("ia moveToNode array: ", Array.isArray(moveToNode))
                if (!Array.isArray(moveToNode)) {
                    const moveToNode: any = [moveToNode];
                }
                //console.log("ia moveToNode array: ", Array.isArray(moveToNode))

                const multiSapeAry = [];
                let ptObj, ptOrdr;
                if (moveToNode.length > 0) {
                    //a:moveTo
                    Object.keys(moveToNode).forEach(function (key) {
                        const moveToPtNode = moveToNode[key]["a:pt"];
                        if (moveToPtNode !== undefined) {
                            Object.keys(moveToPtNode).forEach(function (key2) {
                                const ptObj: any = {};
                                const moveToNoPt = moveToPtNode[key2];
                                const spX = moveToNoPt["attrs", "x"];//parseInt(moveToNoPt["attrs", "x"]) * slideFactor;
                                const spY = moveToNoPt["attrs", "y"];//parseInt(moveToNoPt["attrs", "y"]) * slideFactor;
                                const ptOrdr = moveToNoPt["attrs", "order"];
                                ptObj.type = "movto";
                                ptObj.order = ptOrdr;
                                ptObj.x = spX;
                                ptObj.y = spY;
                                multiSapeAry.push(ptObj);
                                //console.log(key2, lnToNoPt);

                            });
                        }
                    });
                    //a:lnTo
                    if (lnToNodes !== undefined) {
                        Object.keys(lnToNodes).forEach(function (key) {
                            const lnToPtNode = lnToNodes[key]["a:pt"];
                            if (lnToPtNode !== undefined) {
                                Object.keys(lnToPtNode).forEach(function (key2) {
const ptObj: any = {};
                                    const lnToNoPt = lnToPtNode[key2];
                                    const ptX = lnToNoPt["attrs", "x"];
                                    const ptY = lnToNoPt["attrs", "y"];
const ptOrdr: any = lnToNoPt["attrs", "order"];
                                    ptObj.type = "lnto";
                                    ptObj.order = ptOrdr;
                                    ptObj.x = ptX;
                                    ptObj.y = ptY;
                                    multiSapeAry.push(ptObj);
                                    //console.log(key2, lnToNoPt);
                                });
                            }
                        });
                    }
                    //a:cubicBezTo
                    if (cubicBezToNodes !== undefined) {

                        const cubicBezToPtNodesAry = [];
                        //console.log("cubicBezToNodes: ", cubicBezToNodes, ", is arry: ", Array.isArray(cubicBezToNodes))
                        if (!Array.isArray(cubicBezToNodes)) {
                            const cubicBezToNodes: any = [cubicBezToNodes];
                        }
                        Object.keys(cubicBezToNodes).forEach(function (key) {
                            //console.log("cubicBezTo[" + key + "]:");
                            cubicBezToPtNodesAry.push(cubicBezToNodes[key]["a:pt"]);
                        });

                        //console.log("cubicBezToNodes: ", cubicBezToPtNodesAry)
                        cubicBezToPtNodesAry.forEach(function (key2) {
                            //console.log("cubicBezToPtNodesAry: key2 : ", key2)
                            const nodeObj = {};
                            nodeObj.type = "cubicBezTo";
                            nodeObj.order = key2[0]["attrs"]["order"];
                            const pts_ary = [];
                            key2.forEach(function (pt) {
                                const pt_obj = {
                                    x: pt["attrs"]["x"],
                                    y: pt["attrs"]["y"]
                                }
                                pts_ary.push(pt_obj)
                            })
                            nodeObj.cubBzPt = pts_ary;//key2;
                            multiSapeAry.push(nodeObj);
                        });
                    }
                    //a:arcTo
                    if (arcToNodes !== undefined) {
                        const arcToNodesAttrs = arcToNodes["attrs"];
                        const arcOrder = arcToNodesAttrs["order"];
                        const hR = arcToNodesAttrs["hR"];
                        const wR = arcToNodesAttrs["wR"];
                        const stAng = arcToNodesAttrs["stAng"];
                        const swAng = arcToNodesAttrs["swAng"];
                        let shftX = 0;
                        let shftY = 0;
                        const arcToPtNode = PPTXUtils.getTextByPathList(arcToNodes, ["a:pt", "attrs"]);
                        if (arcToPtNode !== undefined) {
                            const shftX: any = arcToPtNode["x"];
                            const shftY: any = arcToPtNode["y"];
                            //console.log("shftX: ",shftX," shftY: ",shftY)
                        }
const ptObj: any = {};
                        ptObj.type = "arcTo";
                        ptObj.order = arcOrder;
                        ptObj.hR = hR;
                        ptObj.wR = wR;
                        ptObj.stAng = stAng;
                        ptObj.swAng = swAng;
                        ptObj.shftX = shftX;
                        ptObj.shftY = shftY;
                        multiSapeAry.push(ptObj);

                    }
                    //a:quadBezTo

                    //a:close
                    if (closeNode !== undefined) {

                        if (!Array.isArray(closeNode)) {
                            const closeNode: any = [closeNode];
                        }
                        // Object.keys(closeNode).forEach(function (key) {
                        //     //console.log("cubicBezTo[" + key + "]:");
                        //     cubicBezToPtNodesAry.push(closeNode[key]["a:pt"]);
                        // });
                        Object.keys(closeNode).forEach(function (key) {
                            //console.log("custShapType >> closeNode: key: ", key);
                            const clsAttrs = closeNode[key]["attrs"];
                            //const clsAttrs = closeNode["attrs"];
                            const clsOrder = clsAttrs["order"];
const ptObj: any = {};
                            ptObj.type = "close";
                            ptObj.order = clsOrder;
                            multiSapeAry.push(ptObj);

                        });

                    }

                    // console.log("custShapType >> multiSapeAry: ", multiSapeAry);

                    multiSapeAry.sort(function (a, b) {
                        return a.order - b.order;
                    });

                    //console.log("custShapType >>sorted  multiSapeAry: ");
                    //console.log(multiSapeAry);
                    let k = 0;
                    let d_val = "";
                    let spX, spY, hR, wR, stAng, swAng;
                    while (k < multiSapeAry.length) {

                        if (multiSapeAry[k].type == "movto") {
                            //start point
                            const spX: any = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                            const spY: any = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                            // if (d == "") {
                            //     d = "M" + spX + "," + spY;
                            // } else {
                            //     //shape without close : then close the shape and start new path
                            //     result += "<path d='" + d + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            //         "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                            //     result += "/>";

                            //     if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            //         result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
                            //     }
                            //     if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            //         result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
                            //     }
                            //     result += "/>";

                            //     d = "M" + spX + "," + spY;
                            //     isClose = true;
                            // }

                            d_val += " M" + spX + "," + spY;

                        } else if (multiSapeAry[k].type == "lnto") {
                            const Lx = parseInt(multiSapeAry[k].x) * cX;//slideFactor;
                            const Ly = parseInt(multiSapeAry[k].y) * cY;//slideFactor;
                            d_val += " L" + Lx + "," + Ly;

                        } else if (multiSapeAry[k].type == "cubicBezTo") {
                            const Cx1 = parseInt(multiSapeAry[k].cubBzPt[0].x) * cX;//slideFactor;
                            const Cy1 = parseInt(multiSapeAry[k].cubBzPt[0].y) * cY;//slideFactor;
                            const Cx2 = parseInt(multiSapeAry[k].cubBzPt[1].x) * cX;//slideFactor;
                            const Cy2 = parseInt(multiSapeAry[k].cubBzPt[1].y) * cY;//slideFactor;
                            const Cx3 = parseInt(multiSapeAry[k].cubBzPt[2].x) * cX;//slideFactor;
                            const Cy3 = parseInt(multiSapeAry[k].cubBzPt[2].y) * cY;//slideFactor;
                            d_val += " C" + Cx1 + "," + Cy1 + " " + Cx2 + "," + Cy2 + " " + Cx3 + "," + Cy3;
                        } else if (multiSapeAry[k].type == "arcTo") {
                            const hR: any = parseInt(multiSapeAry[k].hR) * cX;//slideFactor;
                            const wR: any = parseInt(multiSapeAry[k].wR) * cY;//slideFactor;
                            const stAng: any = parseInt(multiSapeAry[k].stAng) / 60000;
                            const swAng: any = parseInt(multiSapeAry[k].swAng) / 60000;
                            //const shftX = parseInt(multiSapeAry[k].shftX) * slideFactor;
                            //const shftY = parseInt(multiSapeAry[k].shftY) * slideFactor;
                            const endAng = stAng + swAng;

                            d_val += PPTXShapeUtils.shapeArc(wR, hR, wR, hR, stAng, endAng, false);
                        } else if (multiSapeAry[k].type == "quadBezTo") {

                        } else if (multiSapeAry[k].type == "close") {
                            // result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                            //     "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                            // result += "/>";
                            // d_val = "";
                            // isClose = true;

                            d_val += "z";
                        }
                        k++;
                    }
                    //if (!isClose) {
                    //only one "moveTo" and no "close"
                    result += "<path d='" + d_val + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") +
                        "' stroke='" + ((border === undefined) ? "" : border.color) + "' stroke-width='" + ((border === undefined) ? "" : border.width) + "' stroke-dasharray='" + ((border === undefined) ? "" : border.strokeDasharray) + "' ";
                    result += "/>";
                    //console.log(result);
                }

                result += "</svg>";
                result += "<div class='block " + PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) + //block content 
                    " " + PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    " z-index: " + order + `;transform: rotate(` + ((txtRotate !== undefined) ? txtRotate : 0) + `deg);'>`;

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        const type: any = "shape";
                    }
                    result += PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable); //type=shape
                }
                result += "</div>";

                // result = "";
            } else {

                result += "<div class='block " + PPTXTextStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +//block content 
                    " " + PPTXTextStyleUtils.getContentDir(node, type, warpObj) +
                    "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                    "' style='" +
                    PPTXUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType) +
                    PPTXUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) +
                    PPTXStyleManager.getBorder(node, pNode, false, "shape", warpObj) +
                    PPTXShapeFillsUtils.getShapeFill(node, pNode, false, warpObj, source) +
                    " z-index: " + order + `;transform: rotate(` + ((txtRotate !== undefined) ? txtRotate : 0) + `deg);'>`;

                // TextBody
                if (node["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    result += PPTXTextElementUtils.genTextBody(node["p:txBody"], node, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, undefined, styleTable);
                }
                result += "</div>";

            }
            //console.log("div block result:\n", result)
            return result;
        }

export { genShape };
