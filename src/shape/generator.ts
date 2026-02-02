/**
 * pptx-shape-generator.ts
 * 形状生成器模块
 *
 * 这个模块包含了 genShape 函数，用于生成形状的SVG HTML表示
 * 由于 genShape 函数非常庞大（超过5000行），它处理所有形状类型的生成逻辑
 */

import { PPTXUtils } from '../core/utils';
import { PPTXConstants } from '../core/constants';
import { PPTXShapePropertyExtractor } from './property-extractor';
import { PPTXShapeFillsUtils } from './fills';
import { PPTXStyleManager } from '../core/style-manager';
import { PPTXColorUtils } from '../core/color';
import { PPTXTextStyleUtils } from '../text/style';
import { PPTXTextElementUtils } from '../text/element';
import { PPTXBasicShapes } from './basic';
import { PPTXStarShapes } from './star';
import { PPTXFlowchartShapes } from './flowchart';
import { PPTXActionButtonShapes } from './actionbutton';
import { PPTXArrowShapes } from './arrow';
import { PPTXCalloutShapes } from './callout';
import { PPTXShapeContainer } from './container';
import { PPTXShapeUtils } from './shape';
import { PPTXMathShapes } from './math';

const slideFactor = PPTXConstants.SLIDE_FACTOR;

// 定义形状相关的类型接口
interface ShapeNode {
  [key: string]: any;
}

interface BorderStyle {
  color: string;
  width: number;
  strokeDasharray: string;
}

interface FillColor {
  color?: any[];
  rot?: number;
  img?: string;
}

interface StyleTable {
  [key: string]: {
    name: string;
    text: string;
  };
}

interface PositionProps {
  slideXfrmNode: any;
  shapType: string | undefined;
  custShapType: string | undefined;
  rotate: number;
  flip: string;
  txtRotate: number;
  shpId: string;
  w: number;
  h: number;
  x: number;
  y: number;
  slideLayoutXfrmNode: any;
  slideMasterXfrmNode: any;
}

/**
 * 生成形状的SVG HTML表示
 * 
 * @param node - 形状节点对象
 * @param pNode - 父节点对象
 * @param slideLayoutSpNode - 幻灯片布局中的形状节点
 * @param slideMasterSpNode - 幻灯片母版中的形状节点  
 * @param id - 形状ID
 * @param name - 形状名称
 * @param idx - 形状索引
 * @param type - 形状类型
 * @param order - 显示顺序
 * @param warpObj - 包装对象，包含解析上下文
 * @param isUserDrawnBg - 是否用户绘制的背景
 * @param sType - 形状类型标识
 * @param source - 来源标识
 * @param styleTable - 样式表
 * @returns 形状的SVG HTML字符串
 */
export async function genShape(
  node: ShapeNode, 
  pNode: ShapeNode, 
  slideLayoutSpNode: ShapeNode, 
  slideMasterSpNode: ShapeNode, 
  id: string, 
  name: string, 
  idx: string, 
  type: string, 
  order: number, 
  warpObj: any, 
  isUserDrawnBg: boolean, 
  sType: string, 
  source: string, 
  styleTable: StyleTable = {}
): Promise<string> {
  let result = "";
  let dVal: any;
  // 使用属性提取器获取形状属性
  const props: PositionProps = PPTXShapePropertyExtractor.extractShapeProperties(node, slideFactor, pNode, slideLayoutSpNode, slideMasterSpNode);
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
  let fillColor: FillColor | string | undefined;
  let border: BorderStyle | undefined;
  let oShadowSvgUrlStr: string | undefined;
  let headEndNodeAttrs: any;
  let tailEndNodeAttrs: any;

  if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {
    // 创建SVG容器
    const svgCssName = "_svg_css_" + (Object.keys(styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
    const effectsClassName = svgCssName + "_effects";
    let svgStyle = PPTXUtils.getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
        PPTXUtils.getSize(slideXfrmNode, undefined, undefined) +
        " z-index: " + order + `;transform: rotate(` + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + `;`;
    // 如果有阴影效果，添加到SVG样式中
    if (oShadowSvgUrlStr && oShadowSvgUrlStr !== "") {
        svgStyle += oShadowSvgUrlStr.replace('filter:', '');
    }
    result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + `' style='` + svgStyle + `'>`;
    result += '<defs>';
  
    // 获取填充和边框
    fillColor = PPTXShapeFillsUtils.getShapeFill(node, pNode, true, warpObj, source);
    border = PPTXStyleManager.getBorder(node, pNode, true, "shape", warpObj);
    
    // 检查填充类型
    let clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    if (clrFillType == "GROUP_FILL") {
      clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
    }
    // if (clrFillType == "") {
    //     const clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:style","a:fillRef"]));
    // }
    /////////////////////////////////////////                    
    if (clrFillType == "GRADIENT_FILL") {
        grndFillFlg = true;
        const color_arry = (fillColor as FillColor).color;
        const angl = (fillColor as FillColor).rot + 90;
        const svgGrdnt = PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, shpId);
        //fill="url(#linGrd)"
        //console.log("genShape: svgGrdnt: ", svgGrdnt)
        result += svgGrdnt;

    } else if (clrFillType == "PIC_FILL") {
        imgFillFlg = true;
        // 提取图片 URL（fillColor 可能是对象或字符串）
        const imgFill = typeof fillColor === 'object' && (fillColor as FillColor).img ? (fillColor as FillColor).img : fillColor;
        const svgBgImg = PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, shpId, warpObj);
        //fill="url(#imgPtrn)"
        //console.log(svgBgImg)
        result += svgBgImg;
    } else if (clrFillType == "PATTERN_FILL") {
        let styleText = fillColor as string;
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
        const triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border!.color + "' fill='" + border!.color +
            "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
        result += triangleMarker;
    }
    
    result += '</defs>';

    // 处理不同形状类型
    switch (shapType) {
      case "rect": {
        result += PPTXBasicShapes.genRectWithDecoration(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!, oShadowSvgUrlStr, shapType);
        break;
      }
      case "flowChartProcess":
      case "flowChartPredefinedProcess":
      case "flowChartInternalStorage":
      case "actionButtonBlank": {
        result += PPTXBasicShapes.genRectWithDecoration(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!, oShadowSvgUrlStr, shapType);
        break;
      }
      case "flowChartCollate": {
        result += PPTXFlowchartShapes.genFlowChartCollate(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      }
      case "flowChartDocument": {
        result += PPTXFlowchartShapes.genFlowChartDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      }
      case "flowChartMultidocument": {
        result += PPTXFlowchartShapes.genFlowChartMultidocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      }
      case "actionButtonBackPrevious":
        result += PPTXActionButtonShapes.genActionButtonBackPrevious(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonBeginning":
        result += PPTXActionButtonShapes.genActionButtonBeginning(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonDocument":
        result += PPTXActionButtonShapes.genActionButtonDocument(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonEnd":
        result += PPTXActionButtonShapes.genActionButtonEnd(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonForwardNext":
        result += PPTXActionButtonShapes.genActionButtonForwardNext(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonHelp":
        result += PPTXActionButtonShapes.genActionButtonHelp(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonHome":
        result += PPTXActionButtonShapes.genActionButtonHome(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonInformation":
        result += PPTXActionButtonShapes.genActionButtonInformation(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonMovie":
        result += PPTXActionButtonShapes.genActionButtonMovie(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonReturn":
        result += PPTXActionButtonShapes.genActionButtonReturn(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "actionButtonSound":
        result += PPTXActionButtonShapes.genActionButtonSound(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        break;
      case "irregularSeal1":
      case "irregularSeal2": {
        let d: string;
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
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartTerminator": {
        const x1 = w * 3475 / 21600;
        const x2 = w * 18125 / 21600;
        const y1 = h * 10800 / 21600;
        const cd2 = 180;
        const c3d4 = 270;
        const cd4 = 90;
        const d = "M" + x1 + "," + 0 +
          " L" + x2 + "," + 0 +
          PPTXShapeUtils.shapeArc(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
          " L" + x1 + "," + h +
          PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace("M", "L") +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartPunchedTape": {
        const x1 = w * 5 / 20;
        const y1 = h * 2 / 20;
        const y2 = h * 18 / 20;
        const cd2 = 180;
        const d = "M" + 0 + "," + y1 +
          PPTXShapeUtils.shapeArc(x1, y1, x1, y1, cd2, 0, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace("M", "L") +
          " L" + w + "," + y2 +
          PPTXShapeUtils.shapeArc(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(x1, y2, x1, y1, 0, cd2, false).replace("M", "L") +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartOnlineStorage": {
        const x1 = w * 1 / 6;
        const y1 = h * 3 / 6;
        const c3d4 = 270;
        const cd4 = 90;
        const d = "M" + x1 + "," + 0 +
          " L" + w + "," + 0 +
          PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, 90, false).replace("M", "L") +
          " L" + x1 + "," + h +
          PPTXShapeUtils.shapeArc(x1, h / 2, x1, y1, cd4, 270, false).replace("M", "L") +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartDisplay": {
        const x1 = w * 1 / 6;
        const x2 = w * 5 / 6;
        const y1 = h * 3 / 6;
        const c3d4 = 270;
        const cd2 = 180;
        const d = "M" + 0 + "," + y1 +
          " L" + x1 + "," + 0 +
          " L" + x2 + "," + 0 +
          PPTXShapeUtils.shapeArc(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace("M", "L") +
          " L" + x1 + "," + h +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartDelay": {
        const wd2 = w / 2;
        const hd2 = h / 2;
        const cd2 = 180;
        const c3d4 = 270;
        const cd4 = 90;
        const d = "M" + 0 + "," + 0 +
          " L" + wd2 + "," + 0 +
          PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L") +
          " L" + 0 + "," + h +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartMagneticTape": {
        const wd2 = w / 2;
        const hd2 = h / 2;
        const cd2 = 180;
        const c3d4 = 270;
        const cd4 = 90;
        const idy = hd2 * Math.sin(Math.PI / 4);
        const ib = hd2 + idy;
        const ang1 = Math.atan(h / w);
        const ang1Dg = ang1 * 180 / Math.PI;
        const d = "M" + wd2 + "," + h +
          PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd4, cd2, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, 360, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace("M", "L") +
          " L" + w + "," + ib +
          " L" + w + "," + h +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "ellipse":
      case "flowChartConnector":
      case "flowChartSummingJunction":
      case "flowChartOr": {
        result += PPTXBasicShapes.genEllipse(w, h, imgFillFlg, grndFillFlg, shpId, fillColor as string, border!);
        if (shapType == "flowChartOr") {
          result += " <polyline points='" + w / 2 + " " + 0 + "," + w / 2 + " " + h + "' fill='none' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
          result += " <polyline points='" + 0 + " " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        } else if (shapType == "flowChartSummingJunction") {
          const hc = w / 2;
          const vc = h / 2;
          const wd2 = w / 2;
          const hd2 = h / 2;
          const angVal = Math.PI / 4;
          const iDx = wd2 * Math.cos(angVal);
          const idy = hd2 * Math.sin(angVal);
          const il = hc - iDx;
          const ir = hc + iDx;
          const it = vc - idy;
          const ib = vc + idy;
          result += " <polyline points='" + il + " " + it + "," + ir + " " + ib + "' fill='none' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
          result += " <polyline points='" + ir + " " + it + "," + il + " " + ib + "' fill='none' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        }
        break;
      }
      case "roundRect":
      case "round1Rect":
      case "round2DiagRect":
      case "round2SameRect":
      case "snip1Rect":
      case "snip2DiagRect":
      case "snip2SameRect":
      case "flowChartAlternateProcess":
      case "flowChartPunchedCard": {
        let sAdj1_val = 0.33334;
        let sAdj2_val = 0;
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj2_val = parseInt(sAdj.substr(4)) / 50000;
            }
          }
        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
          const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
          sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
          sAdj2_val = 0;
        }
        
        let shpTyp: string, adjTyp: string;
        let tranglRott = "";
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
        const d_val = PPTXShapeUtils.shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path " + tranglRott + "  d='" + d_val + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "snipRoundRect": {
        let sAdj1_val = 0.33334;
        let sAdj2_val = 0.33334;
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj2_val = parseInt(sAdj.substr(4)) / 50000;
            }
          }
        }
        const d_val = "M0," + h + " L" + w + "," + h + " L" + w + "," + (h / 2) * sAdj2_val +
          " L" + (w / 2 + (w / 2) * (1 - sAdj2_val)) + ",0 L" + (w / 2) * sAdj1_val + ",0 Q0,0 0," + (h / 2) * sAdj1_val + " z";

        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d_val + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "bentConnector2": {
        const d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
        result += "<path d='" + d + "' stroke='" + border!.color +
          "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' fill='none' ";
        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
          result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
        }
        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
          result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
        }
        result += "/>";
        break;
      }
      case "rtTriangle": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='0 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "triangle":
      case "flowChartExtract":
      case "flowChartMerge": {
        let shapAdjst_val = 0.5;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          shapAdjst_val = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let tranglRott = "";
        if (shapType == "flowChartMerge") {
          tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
        }
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon " + tranglRott + " points='" + (w * shapAdjst_val) + " 0,0 " + h + "," + w + " " + h + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "diamond":
      case "flowChartDecision":
      case "flowChartSort": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + (w / 2) + " 0,0 " + (h / 2) + "," + (w / 2) + " " + h + "," + w + " " + (h / 2) + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        if (shapType == "flowChartSort") {
          result += " <polyline points='0 " + h / 2 + "," + w + " " + h / 2 + "' fill='none' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        }
        break;
      }
      case "trapezoid":
      case "flowChartManualOperation":
      case "flowChartManualInput": {
        let adjst_val = 0.2;
        const max_adj_const = 0.7407;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          const adjst = parseInt(shapAdjst.substr(4)) * slideFactor;
          adjst_val = (adjst * 0.5) / max_adj_const;
        }
        let cnstVal = 0;
        let tranglRott = "";
        if (shapType == "flowChartManualOperation") {
          tranglRott = "transform='rotate(180 " + w / 2 + "," + h / 2 + ")'";
        }
        if (shapType == "flowChartManualInput") {
          adjst_val = 0;
          cnstVal = h / 5;
        }
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon " + tranglRott + " points='" + (w * adjst_val) + " " + cnstVal + ",0 " + h + "," + w + " " + h + "," + (1 - adjst_val) * w + " 0' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "parallelogram":
      case "flowChartInputOutput": {
        let adjst_val = 0.25;
        let max_adj_const: number;
        if (w > h) {
          max_adj_const = w / h;
        } else {
          max_adj_const = h / w;
        }
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          const adjst = parseInt(shapAdjst.substr(4)) / 100000;
          adjst_val = adjst / max_adj_const!;
        }
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + adjst_val * w + " 0,0 " + h + "," + (1 - adjst_val) * w + " " + h + "," + w + " 0' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "pentagon": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + (0.5 * w) + " 0,0 " + (0.375 * h) + "," + (0.15 * w) + " " + h + "," + 0.85 * w + " " + h + "," + w + " " + 0.375 * h + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "hexagon":
      case "flowChartPreparation": {
        let adj = 25000 * slideFactor;
        const vf = 115470 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const angVal1 = 60 * Math.PI / 180;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        const vc = h / 2;
        const hd2 = h / 2;
        const ss = Math.min(w, h);
        const maxAdj = cnstVal1 * w / ss;
        const a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const shd2 = hd2 * vf / cnstVal2;
        const x1 = ss * a / cnstVal2;
        const x2 = w - x1;
        const dy1 = shd2 * Math.sin(angVal1);
        const y1 = vc - dy1;
        const y2 = vc + dy1;

        const d = "M" + 0 + "," + vc +
          " L" + x1 + "," + y1 +
          " L" + x2 + "," + y1 +
          " L" + w + "," + vc +
          " L" + x2 + "," + y2 +
          " L" + x1 + "," + y2 +
          " z";

        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "heptagon": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + (0.5 * w) + " 0," + w / 8 + " " + h / 4 + ",0 " + (5 / 8) * h + "," + w / 4 + " " + h + "," + (3 / 4) * w + " " + h + "," +
          w + " " + (5 / 8) * h + "," + (7 / 8) * w + " " + h / 4 + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "octagon": {
        let adj1 = 0.25;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj1 = parseInt(shapAdjst.substr(4)) / 100000;
        }
        const adj2 = (1 - adj1);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + adj1 * w + " 0,0 " + adj1 * h + ",0 " + adj2 * h + "," + adj1 * w + " " + h + "," + adj2 * w + " " + h + "," +
          w + " " + adj2 * h + "," + w + " " + adj1 * h + "," + adj2 * w + " 0' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";

        break;
      }
      case "decagon": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + h / 2 + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
          (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + h / 2 + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "dodecagon": {
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='" + (3 / 8) * w + " 0," + w / 8 + " " + h / 8 + ",0 " + (3 / 8) * h + ",0 " + (5 / 8) * h + "," + w / 8 + " " + (7 / 8) * h + "," + (3 / 8) * w + " " + h + "," +
          (5 / 8) * w + " " + h + "," + (7 / 8) * w + " " + (7 / 8) * h + "," + w + " " + (5 / 8) * h + "," + w + " " + (3 / 8) * h + "," + (7 / 8) * w + " " + h / 8 + "," + (5 / 8) * w + " 0' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star4": {
        const d = PPTXStarShapes.genStar4(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star5": {
        const d = PPTXStarShapes.genStar5(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star6": {
        const d = PPTXStarShapes.genStar6(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star7": {
        const d = PPTXStarShapes.genStar7(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star8": {
        const d = PPTXStarShapes.genStar8(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star10": {
        const d = PPTXStarShapes.genStar10(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star12": {
        const d = PPTXStarShapes.genStar12(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star16": {
        const d = PPTXStarShapes.genStar16(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star24": {
        const d = PPTXStarShapes.genStar24(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "star32": {
        const d = PPTXStarShapes.genStar32(w, h, node, slideFactor);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "pie":
      case "pieWedge":
      case "arc": {
        let adj1: number;
        let adj2: number;
        let H: number;
        let isClose: boolean;
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
        
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst !== undefined) {
          let shapAdjst1 = PPTXUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
          let shapAdjst2 = shapAdjst1;
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
        const pieVals = PPTXShapeUtils.shapePie(H, w, adj1, adj2, isClose);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + pieVals[0] + "' transform='" + pieVals[1] + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "chord": {
        let sAdj1_val = 45;
        let sAdj2_val = 270;
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj1_val = parseInt(sAdj.substr(4)) / 60000;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj2_val = parseInt(sAdj.substr(4)) / 60000;
            }
          }
        }
        const hR = h / 2;
        const wR = w / 2;
        const d_val = PPTXShapeUtils.shapeArc(wR, hR, wR, hR, sAdj1_val, sAdj2_val, true);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d_val + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "frame": {
        let adj1 = 12500 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a1: number;
        let x1: number;
        let x4: number;
        let y4: number;
        if (adj1 < 0) a1 = 0
        else if (adj1 > cnstVal1) a1 = cnstVal1
        else a1 = adj1
        x1 = Math.min(w, h) * a1 / cnstVal2;
        x4 = w - x1;
        y4 = h - x1;
        const d = "M" + 0 + "," + 0 +
          " L" + w + "," + 0 +
          " L" + w + "," + h +
          " L" + 0 + "," + h +
          ` zM` + x1 + "," + x1 +
          " L" + x1 + "," + y4 +
          " L" + x4 + "," + y4 +
          " L" + x4 + "," + x1 +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "donut": {
        let adj = 25000 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        let dr: number;
        let iwd2: number;
        let ihd2: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        const d = "M" + 0 + "," + h / 2 +
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
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "noSmoking": {
        let adj = 18750 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        let dr: number;
        let iwd2: number;
        let ihd2: number;
        let ang: number;
        let ct: number;
        let st: number;
        let m: number;
        let n: number;
        let drd2: number;
        let dang: number;
        let dang2: number;
        let swAng: number;
        let stAng1: number;
        let stAng2: number;
        let ct1: number;
        let st1: number;
        let m1: number;
        let n1: number;
        let dx1: number;
        let dy1: number;
        let x1: number;
        let y1: number;
        let x2: number;
        let y2: number;
        let stAng1deg: number;
        let stAng2deg: number;
        let swAng2deg: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        ang = Math.atan(h / w);
        ct = ihd2 * Math.cos(ang);
        st = iwd2 * Math.sin(ang);
        m = Math.sqrt(ct * ct + st * st); //"mod ct st 0"
        n = iwd2 * ihd2 / m;
        drd2 = dr / 2;
        dang = Math.atan(drd2 / n);
        dang2 = dang * 2;
        swAng = -Math.PI + dang2;
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
        stAng1deg = stAng1 * 180 / Math.PI;
        stAng2deg = stAng2 * 180 / Math.PI;
        swAng2deg = swAng * 180 / Math.PI;
        const d = "M" + 0 + "," + h / 2 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
          ` zM` + x1 + "," + y1 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
          ` zM` + x2 + "," + y2 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "mathPlus":
        result += PPTXMathShapes.genMathPlus(w, h, node, slideFactor);
        break;
      case "mathMinus":
        result += PPTXMathShapes.genMathMinus(w, h, node, slideFactor);
        break;
      case "mathMultiply":
        result += PPTXMathShapes.genMathMultiply(w, h, node, slideFactor);
        break;
      case "mathDivide": {
        // 从JavaScript版本复制的复杂逻辑
        let adj1 = 0;
        let adj2 = 0;
        let adj3 = 0;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          if (Array.isArray(shapAdjst_ary)) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
              const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
              if (sAdj_name == "adj1") {
                const sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj1 = parseInt(sAdj1.substr(4));
              } else if (sAdj_name == "adj2") {
                const sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj2 = parseInt(sAdj2.substr(4));
              } else if (sAdj_name == "adj3") {
                const sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj3 = parseInt(sAdj3.substr(4));
              }
            }
          } else {
            const sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            adj1 = parseInt(sAdj1.substr(4));
          }
        }
        
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const cnstVal3 = 200000 * slideFactor;
        const cnstVal4 = 1000 * slideFactor;
        const cnstVal5 = 36745 * slideFactor;
        const cnstVal6 = 73490 * slideFactor;
        
        if (shapAdjst_ary === undefined) {
          adj1 = 23520 * slideFactor;
          adj2 = 5880 * slideFactor;
          adj3 = 11760 * slideFactor;
        } else {
          adj1 = adj1 * slideFactor;
          adj2 = adj2 * slideFactor;
          adj3 = adj3 * slideFactor;
        }
        
        let a1, ma1, ma3h, ma3w, maxAdj3_val, a3, m4a3, maxAdj2_val, a2, dy1, yg, rad, dx1,
            y3, y4, a, y2, y1, y5, x1, x3, x2;
        
        a1 = (adj1 < cnstVal4) ? cnstVal4 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        ma1 = -a1;
        ma3h = (cnstVal6 + ma1) / 4;
        ma3w = cnstVal5 * w / h;
        maxAdj3_val = (ma3h < ma3w) ? ma3h : ma3w;
        a3 = (adj3 < cnstVal4) ? cnstVal4 : (adj3 > maxAdj3_val) ? maxAdj3_val : adj3;
        m4a3 = -4 * a3;
        maxAdj2_val = cnstVal6 + m4a3 - a1;
        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2_val) ? maxAdj2_val : adj2;
        
        dy1 = h * a1 / cnstVal3;
        yg = h * a2 / cnstVal2;
        rad = h * a3 / cnstVal2;
        dx1 = w * cnstVal6 / cnstVal3;
        
        y3 = h / 2 - dy1;
        y4 = h / 2 + dy1;
        a = yg + rad;
        y2 = y3 - a;
        y1 = y2 - rad;
        y5 = h - y1;
        
        x1 = w / 2 - dx1;
        x3 = w / 2 + dx1;
        x2 = w / 2 - rad;
        
        const cd4 = 90, c3d4 = 270;
        const cX1 = w / 2 - Math.cos(c3d4 * Math.PI / 180) * rad;
        const cY1 = y1 - Math.sin(c3d4 * Math.PI / 180) * rad;
        const cX2 = w / 2 - Math.cos(Math.PI / 2) * rad;
        const cY2 = y5 - Math.sin(Math.PI / 2) * rad;
        
        const dVal = "M" + w / 2 + "," + y1 +
            PPTXShapeUtils.shapeArc(cX1, cY1, rad, rad, c3d4, c3d4 + 360, false).replace("M", "L") +
            ` z M` + w / 2 + "," + y5 +
            PPTXShapeUtils.shapeArc(cX2, cY2, rad, rad, cd4, cd4 + 360, false).replace("M", "L") +
            ` z M` + x1 + "," + y3 +
            " L" + x3 + "," + y3 +
            " L" + x3 + "," + y4 +
            " L" + x1 + "," + y4 +
            " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + dVal + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "mathEqual":
        result += PPTXMathShapes.genMathEqual(w, h, node, slideFactor);
        break;
      case "mathNotEqual": {
        // 从JavaScript版本复制的复杂逻辑
        let adj1 = 0;
        let adj2 = 110 * Math.PI / 180;
        let adj3 = 0;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          if (Array.isArray(shapAdjst_ary)) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
              const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
              if (sAdj_name == "adj1") {
                const sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj1 = parseInt(sAdj1.substr(4));
              } else if (sAdj_name == "adj2") {
                const sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
              } else if (sAdj_name == "adj3") {
                const sAdj3 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                adj3 = parseInt(sAdj3.substr(4));
              }
            }
          } else {
            const sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
            adj1 = parseInt(sAdj1.substr(4));
          }
        }
        
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const cnstVal3 = 200000 * slideFactor;
        const cnstVal4 = 73490 * slideFactor;
        
        if (shapAdjst_ary === undefined) {
          adj1 = 23520 * slideFactor;
          adj2 = 110 * Math.PI / 180;
          adj3 = 11760 * slideFactor;
        } else {
          adj1 = adj1 * slideFactor;
          adj2 = adj2; // adj2 already converted to radians above
          adj3 = adj3 * slideFactor;
        }
        
        let a1, crAng, a2a1, maxAdj3, a3, dy1, dy2, dx1, x1, x8, y2, y3, y1, y4,
            cadj2, xadj2, len, bhw, bhw2, x7, dx67, x6, dx57, x5, dx47, x4, dx37,
            x3, dx27, x2, rx7, rx6, rx5, rx4, rx3, rx2, dx7, rxt, lxt, rx, lx,
            dy3, dy4, ry, ly, dlx, drx, dly, dry, xC1, xC2, yC1, yC2, yC3, yC4;
        
        const angVal1 = 70 * Math.PI / 180, angVal2 = 110 * Math.PI / 180;
        
        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        crAng = (adj2 < angVal1) ? angVal1 : (adj2 > angVal2) ? angVal2 : adj2;
        a2a1 = a1 * 2;
        maxAdj3 = cnstVal2 - a2a1;
        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
        
        dy1 = h * a1 / cnstVal2;
        dy2 = h * a3 / cnstVal3;
        dx1 = w * cnstVal4 / cnstVal3;
        
        x1 = w / 2 - dx1;
        x8 = w / 2 + dx1;
        y2 = h / 2 - dy2;
        y3 = h / 2 + dy2;
        y1 = y2 - dy1;
        y4 = y3 + dy1;
        
        cadj2 = crAng - Math.PI / 2;
        xadj2 = (h / 2) * Math.tan(cadj2);
        len = Math.sqrt(xadj2 * xadj2 + (h / 2) * (h / 2));
        bhw = len * dy1 / (h / 2);
        bhw2 = bhw / 2;
        
        x7 = w / 2 + xadj2 - bhw2;
        dx67 = xadj2 * y1 / (h / 2);
        x6 = x7 - dx67;
        dx57 = xadj2 * y2 / (h / 2);
        x5 = x7 - dx57;
        dx47 = xadj2 * y3 / (h / 2);
        x4 = x7 - dx47;
        dx37 = xadj2 * y4 / (h / 2);
        x3 = x7 - dx37;
        dx27 = xadj2 * 2;
        x2 = x7 - dx27;
        
        rx7 = x7 + bhw;
        rx6 = x6 + bhw;
        rx5 = x5 + bhw;
        rx4 = x4 + bhw;
        rx3 = x3 + bhw;
        rx2 = x2 + bhw;
        
        dx7 = dy1 * (h / 2) / len;
        rxt = x7 + dx7;
        lxt = rx7 - dx7;
        rx = (cadj2 > 0) ? rxt : rx7;
        lx = (cadj2 > 0) ? x7 : lxt;
        
        dy3 = dy1 * xadj2 / len;
        dy4 = -dy3;
        ry = (cadj2 > 0) ? dy3 : 0;
        ly = (cadj2 > 0) ? 0 : dy4;
        
        dlx = w - rx;
        drx = w - lx;
        dly = h - ry;
        dry = h - ly;
        
        xC1 = (rx + lx) / 2;
        xC2 = (drx + dlx) / 2;
        yC1 = (ry + ly) / 2;
        yC2 = (y1 + y2) / 2;
        yC3 = (y3 + y4) / 2;
        yC4 = (dry + dly) / 2;
        
        const dVal = "M" + x1 + "," + y1 +
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
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + dVal + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "can": {
        let adj = 25000 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        let dr: number;
        let iwd2: number;
        let ihd2: number;
        let ang: number;
        let ct: number;
        let st: number;
        let m: number;
        let n: number;
        let drd2: number;
        let dang: number;
        let dang2: number;
        let swAng: number;
        let stAng1: number;
        let stAng2: number;
        let ct1: number;
        let st1: number;
        let m1: number;
        let n1: number;
        let dx1: number;
        let dy1: number;
        let x1: number;
        let y1: number;
        let x2: number;
        let y2: number;
        let stAng1deg: number;
        let stAng2deg: number;
        let swAng2deg: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        ang = Math.atan(h / w);
        ct = ihd2 * Math.cos(ang);
        st = iwd2 * Math.sin(ang);
        m = Math.sqrt(ct * ct + st * st); //"mod ct st 0"
        n = iwd2 * ihd2 / m;
        drd2 = dr / 2;
        dang = Math.atan(drd2 / n);
        dang2 = dang * 2;
        swAng = -Math.PI + dang2;
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
        stAng1deg = stAng1 * 180 / Math.PI;
        stAng2deg = stAng2 * 180 / Math.PI;
        swAng2deg = swAng * 180 / Math.PI;
        const d = "M" + 0 + "," + h / 2 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
          ` zM` + x1 + "," + y1 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
          ` zM` + x2 + "," + y2 +
          PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "cube": {
        let sAdj1_val = 25000 * slideFactor;
        let sAdj2_val = 25000 * slideFactor;
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj1_val = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj2_val = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        const offX = sAdj1_val * w;
        const offY = sAdj2_val * h;
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += " <polygon points='0 0," + offX + " " + offY + "," + offX + " " + h + ",0 " + h + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />"; // left face
        result += " <polygon points='0 0," + w + " 0," + (w - offX) + " " + offY + "," + offX + " " + offY + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />"; // top face
        result += " <polygon points='" + offX + " " + offY + "," + (w - offX) + " " + offY + "," + (w - offX) + " " + (h - offY) + "," + offX + " " + (h - offY) + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />"; // right face
        break;
      }
      case "bevel": {
        let sAdj1_val = 12500 * slideFactor;
        let sAdj2_val = 12500 * slideFactor;
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj1_val = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              sAdj2_val = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        const wd2 = w / 2;
        const hd2 = h / 2;
        const wd8 = w / 8;
        const hd8 = h / 8;
        const wd32 = w / 32;
        const hd32 = h / 32;
        const adj1_calc = Math.min(wd8, hd8) * sAdj1_val;
        const adj2_calc = Math.min(wd8, hd8) * sAdj2_val;
        const d_val = "M0," + adj1_calc + " L" + adj1_calc + ",0 L" + (w - adj2_calc) + ",0 L" + w + "," + adj2_calc + " L" + w + "," + (h - adj1_calc) + " L" + (w - adj1_calc) + "," + h + " L" + adj2_calc + "," + h + " L0," + (h - adj2_calc) + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d_val + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "torus": {
        let adj = 25000 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        let dr: number;
        let iwd2: number;
        let ihd2: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        dr = Math.min(w, h) * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        const d = "M" + 0 + "," + h / 2 +
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
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "teardrop": {
        const d = "M0," + h + " C" + w + ",0 " + w + "," + h + " 0," + h + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "homePlate": {
        const d = "M0," + h + " L0," + (h / 2) + " L" + w + ",0 L" + w + "," + h + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "corner": {
        let adj = 50000 * slideFactor;
        const cnstVal1 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        const xAdj = w * a / cnstVal1;
        const yAdj = h * a / cnstVal1;
        const d = "M0,0 L" + xAdj + ",0 L" + xAdj + "," + yAdj + " L0," + yAdj + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "diagonalStripe": {
        let adj = 50000 * slideFactor;
        const cnstVal1 = 100000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        const halfH = h / 2;
        const diagH = h * a / cnstVal1;
        const d = "M0," + (halfH - diagH / 2) + " L" + w + "," + (halfH + diagH / 2) + " L" + w + "," + (halfH + diagH / 2 + 10) + " L0," + (halfH - diagH / 2 + 10) + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "foldedCorner": {
        let adj = 16667 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        let a: number;
        if (adj < 0) a = 0
        else if (adj > cnstVal1) a = cnstVal1
        else a = adj
        const foldX = w * a / cnstVal1;
        const foldY = h * a / cnstVal1;
        const d = "M0,0 L" + w + ",0 L" + w + "," + h + " L0," + h + " z M0,0 L" + foldX + ",0 L" + foldX + "," + foldY + " L0," + foldY + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "verticalScroll":
      case "horizontalScroll": {
        const scrollRadius = Math.min(w, h) * 0.1;
        const scrollCenterY = h / 2;
        const scrollStartX = scrollRadius;
        const scrollEndX = w - scrollRadius;
        const d = "M" + scrollStartX + "," + (scrollCenterY - scrollRadius) +
          " A" + scrollRadius + "," + scrollRadius + " 0 1,1 " + scrollStartX + "," + (scrollCenterY + scrollRadius) +
          " L" + scrollEndX + "," + (scrollCenterY + scrollRadius) +
          " A" + scrollRadius + "," + scrollRadius + " 0 1,1 " + scrollEndX + "," + (scrollCenterY - scrollRadius) +
          " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "wedgeRectCallout": {
        // 简化的楔形矩形标注形状
        const d = "M0,0 L" + w + ",0 L" + w + "," + h + " L" + (w / 4) + "," + h + " L0," + (h * 3 / 4) + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "wedgeRoundRectCallout": {
        // 简化的楔形圆角矩形标注形状
        const roundRadius = Math.min(w, h) * 0.1;
        const d = "M" + roundRadius + ",0 L" + (w - roundRadius) + ",0 " +
          "A" + roundRadius + "," + roundRadius + " 0 0,1 " + w + "," + roundRadius + " " +
          "L" + w + "," + (h - roundRadius) + " " +
          "A" + roundRadius + "," + roundRadius + " 0 0,1 " + (w - roundRadius) + "," + h + " " +
          "L" + (w / 4) + "," + h + " " +
          "L0," + (h * 3 / 4) + " " +
          "A" + roundRadius + "," + roundRadius + " 0 0,1 " + roundRadius + ",0" + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "wedgeEllipseCallout": {
        // 简化的楔形椭圆形标注形状
        const d = "M" + w / 2 + ",0 " +
          "A" + w / 2 + "," + h / 2 + " 0 1,1 " + w / 2 + "," + h + " " +
          "L" + (w / 4) + "," + h + " " +
          "L0," + (h * 3 / 4) + " " +
          "A" + w / 2 + "," + h / 2 + " 0 1,0 " + w / 2 + ",0" + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "cloudCallout": {
        // 简化的云形标注形状
        const cloudRadius = Math.min(w, h) * 0.2;
        const cloudCenterX = w / 2;
        const cloudCenterY = h / 2;
        const d = "M" + (cloudCenterX - cloudRadius) + "," + cloudCenterY + " " +
          "C" + (cloudCenterX - cloudRadius * 1.5) + "," + (cloudCenterY - cloudRadius * 0.5) + " " +
          (cloudCenterX - cloudRadius * 0.5) + "," + (cloudCenterY - cloudRadius) + " " +
          cloudCenterX + "," + (cloudCenterY - cloudRadius * 0.5) + " " +
          "C" + (cloudCenterX + cloudRadius * 0.5) + "," + (cloudCenterY - cloudRadius) + " " +
          (cloudCenterX + cloudRadius * 1.5) + "," + (cloudCenterY - cloudRadius * 0.5) + " " +
          (cloudCenterX + cloudRadius) + "," + cloudCenterY + " " +
          "C" + (cloudCenterX + cloudRadius * 1.5) + "," + (cloudCenterY + cloudRadius * 0.5) + " " +
          (cloudCenterX + cloudRadius * 0.5) + "," + (cloudCenterY + cloudRadius) + " " +
          cloudCenterX + "," + (cloudCenterY + cloudRadius * 0.5) + " " +
          "C" + (cloudCenterX - cloudRadius * 0.5) + "," + (cloudCenterY + cloudRadius) + " " +
          (cloudCenterX - cloudRadius * 1.5) + "," + (cloudCenterY + cloudRadius * 0.5) + " " +
          (cloudCenterX - cloudRadius) + "," + cloudCenterY + " z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "borderCallout1":
      case "borderCallout2":
      case "borderCallout3":
      case "accentCallout1":
      case "accentCallout2":
      case "accentCallout3":
      case "accentBorderCallout1":
      case "accentBorderCallout2":
      case "accentBorderCallout3":
      case "callout1":
      case "callout2":
      case "callout3":
        // 标注形状处理
        result += PPTXCalloutShapes.genCallout(w, h, shapType, node, slideFactor, fillColor as string, border!);
        break;
      case "leftArrow":
      case "rightArrow":
      case "upArrow":
      case "downArrow":
      case "leftRightArrow":
      case "upDownArrow":
      case "quadArrow":
      case "leftRightUpArrow":
      case "bentArrow":
      case "uturnArrow":
      case "leftUpArrow":
      case "bentUpArrow":
      case "curvedRightArrow":
      case "curvedLeftArrow":
      case "curvedUpArrow":
      case "curvedDownArrow":
      case "stripedRightArrow":
      case "notchedRightArrow":
      case "chevron":
      case "rightArrowCallout":
      case "leftArrowCallout":
      case "upArrowCallout":
      case "downArrowCallout":
      case "leftRightArrowCallout":
      case "UpDownArrowCallout":
      case "QuadArrowCallout":
      case "circularArrow":
      case "leftCircularArrow":
      case "leftRightCircularArrow":
      case "swooshArrow":
        result += PPTXArrowShapes.genArrow(w, h, shapType, node, slideFactor, fillColor as string, border!);
        break;
      case "bentConnector3":
      case "bentConnector4":
      case "bentConnector5":
      case "curvedConnector2":
      case "curvedConnector3":
      case "curvedConnector4":
      case "curvedConnector5":
        // 对于连接器形状，生成相应的连线
        result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border!.color +
          "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' ";
        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
          result += "marker-start='url(#markerTriangle_" + shpId + ")' ";
        }
        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
          result += "marker-end='url(#markerTriangle_" + shpId + ")' ";
        }
        result += "/>";
        break;
      case "line":
      case "lineInv":
      case "chartPlus":
      case "chartStar":
      case "chartX":
        // 对于线条，使用基本线条生成函数
        if (shapType === "line") {
          result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        }
        // 其他形状（lineInv, chartPlus, chartStar, chartX）不渲染任何内容
        break;
      case "blockArc": {
        // 块状弧形处理
        let adj1 = 0;
        let adj2 = 10800000;
        const maxAdj = 21599999;
        
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst !== undefined) {
          const adj1Value = PPTXUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
          if (adj1Value !== undefined) {
            adj1 = parseInt(adj1Value.substr(4));
          }
          
          const adj2Value = PPTXUtils.getTextByPathList(shapAdjst, ["attrs", "fmla"]);
          if (adj2Value !== undefined) {
            adj2 = parseInt(adj2Value.substr(4));
          }
        }
        
        // 限制角度范围
        if (adj1 < 0) adj1 = 0;
        if (adj1 > maxAdj) adj1 = maxAdj;
        if (adj2 < 0) adj2 = 0;
        if (adj2 > 21599999) adj2 = 21599999; // 一圈半
        
        // 将角度转换为度数
        const startAngle = adj1 / 60000;
        const endAngle = adj2 / 60000;
        
        const d = PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, startAngle, endAngle, true);
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "bracePair": {
        // 大括号对处理
        const margin = Math.min(w, h) * 0.1;
        const curveHeight = h * 0.4;
        
        const d = "M" + margin + "," + h / 2 + 
              " C" + margin + "," + (h / 2 - curveHeight) + " " + w / 2 + "," + (h / 2 - curveHeight) + " " + w / 2 + "," + h / 2 +
              " C" + margin + "," + (h / 2 + curveHeight) + " " + w / 2 + "," + (h / 2 + curveHeight) + " " + w / 2 + "," + h / 2;
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "bracketPair": {
        // 方括号对处理
        const bracketSize = Math.min(w * 0.1, h * 0.2);
        
        const d = "M0," + bracketSize + " L" + bracketSize + "," + bracketSize + " L" + bracketSize + "," + (h - bracketSize) + 
              " L0," + (h - bracketSize) + 
              " M" + w + "," + bracketSize + " L" + (w - bracketSize) + "," + bracketSize + 
              " L" + (w - bracketSize) + "," + (h - bracketSize) + " L" + w + "," + (h - bracketSize);
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "gear6": {
        // 6齿齿轮处理
        const centerX = w / 2;
        const centerY = h / 2;
        const outerRadius = Math.min(w, h) * 0.4;
        const innerRadius = outerRadius * 0.6;
        const toothDepth = outerRadius * 0.2;
        
        let d = "M" + (centerX + outerRadius) + "," + centerY;
        
        // 画6个齿
        for (let i = 1; i <= 12; i++) { // 12段：6个外侧+6个内侧
          const angle = (i * Math.PI) / 6; // 每30度一段
          let radius = (i % 2 === 0) ? innerRadius : outerRadius + toothDepth;
          
          const x = centerX + Math.cos(angle) * radius;
          const y = centerY + Math.sin(angle) * radius;
          d += " L" + x + "," + y;
        }
        
        d += " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "gear9": {
        // 9齿齿轮形状处理
        const d = PPTXShapeUtils.shapeGear(w, h / 3.5, 9);
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' transform='rotate(20," + (3 / 7) * h + "," + (3 / 7) * h + ")' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "cloud": {
        // 云朵形状处理
        result += PPTXCalloutShapes.genCloudCallout(w, h, node, slideFactor);
        break;
      }
      case "smileyFace": {
        // 笑脸处理
        let adj = 4653 * slideFactor;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const cnstVal3 = 4653 * slideFactor;
        
        // 限制调整值
        let a = (adj < -cnstVal3) ? -cnstVal3 : (adj > cnstVal3) ? cnstVal3 : adj;
        
        const x1 = w * 4969 / 21699;
        const x2 = w * 6215 / 21600;
        const x3 = w * 13135 / 21600;
        const x4 = w * 16640 / 21600;
        const y1 = h * 7570 / 21600;
        const y3 = h * 16515 / 21600;
        const dy2 = h * a / cnstVal2;
        const y2 = y3 - dy2;
        const y4 = y3 + dy2;
        const dy3 = h * a / cnstVal1;
        const y5 = y4 + dy3;
        const wR = w * 1125 / 21600;
        const hR = h * 1125 / 21600;
        
        const wd2 = w / 2;
        const hd2 = h / 2;
        
        const cX1 = x2 - wR * Math.cos(Math.PI);
        const cY1 = y1 - hR * Math.sin(Math.PI);
        const cX2 = x3 - wR * Math.cos(Math.PI);
        
        const d = //眼睛
              PPTXShapeUtils.shapeArc(cX1, cY1, wR, hR, 180, 540, false) +
              PPTXShapeUtils.shapeArc(cX2, cY1, wR, hR, 180, 540, false) +
              //嘴巴
              " M" + x1 + "," + y2 +
              " Q" + wd2 + "," + y5 + " " + x4 + "," + y2 +
              " Q" + wd2 + "," + y5 + " " + x1 + "," + y2 +
              //头部
              " M" + 0 + "," + hd2 +
              PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, 180, 540, false).replace("M", "L") +
              " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "plus": {
        // 加号形状处理
        let adj1 = 0.25;
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj1 = parseInt(shapAdjst.substr(4)) / 100000;
        }
        
        const adj2 = (1 - adj1);
        const d = " M" + adj1 * w + " 0 L" + adj1 * w + " " + adj1 * h + " L0 " + adj1 * h + " L0 " + adj2 * h + " L" + 
              adj1 * w + " " + adj2 * h + " L" + adj1 * w + " " + h + " L" + adj2 * w + " " + h + " L" + adj2 * w + " " + adj2 * h + " L" + w + " " + adj2 * h + " L" +
              w + " " + adj1 * h + " L" + adj2 * w + " " + adj1 * h + " L" + adj2 * w + " 0 Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "wave": {
        // 波浪形状处理
        let adj1 = 12500 * slideFactor;
        let adj2 = 0;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj1 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj2 = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        
        const cnstVal2 = -10000 * slideFactor;
        const cnstVal3 = 50000 * slideFactor;
        const cnstVal4 = 100000 * slideFactor;
        const cnstVal5 = 20000 * slideFactor;
        
        const t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
        
        let a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
        let a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
        
        const y1 = h * a1 / cnstVal4;
        const dy2 = y1 * 10 / 3;
        const y2 = y1 - dy2;
        const y3 = y1 + dy2;
        const y4 = b - y1;
        const y5 = y4 - dy2;
        const y6 = y4 + dy2;
        const of2 = w * a2 / cnstVal3;
        let dx2 = (of2 > 0) ? 0 : of2;
        const x2 = l - dx2;
        let dx5 = (of2 > 0) ? of2 : 0;
        const x5 = r - dx5;
        const dx3 = (dx2 + x5) / 3;
        const x3 = x2 + dx3;
        const x4 = (x3 + x5) / 2;
        const x6 = l + dx5;
        const x10 = r + dx2;
        const x7 = x6 + dx3;
        const x8 = (x7 + x10) / 2;
        
        const d = "M" + x2 + "," + y1 +
              " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
              " L" + x10 + "," + y4 +
              " C" + x8 + "," + y6 + " " + x7 + "," + y5 + " " + x6 + "," + y4 +
              " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "heart": {
        // 心形处理
        const dx1 = w * 49 / 48;
        const dx2 = w * 10 / 48;
        const x1 = w / 2 - dx1;
        const x2 = w / 2 - dx2;
        const x3 = w / 2 + dx2;
        
        const d = "M" + x1 + "," + h / 4 + " C" + x1 + "," + 0 + " " + x2 + "," + 0 + " " + w / 2 + "," + h / 4 +
              " C" + w / 2 + "," + 0 + " " + x3 + "," + 0 + " " + x3 + "," + h / 4 +
              " C" + w / 2 + "," + h / 2 + " " + w + "," + h + " " + w / 2 + "," + h +
              " C0," + h + " " + w / 2 + "," + h / 2 + " " + x1 + "," + h / 4 + " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "leftRightRibbon": {
        // 左右丝带形状处理
        let adj1 = 50000 * slideFactor;
        let adj2 = 50000 * slideFactor;
        let adj3 = 16667 * slideFactor;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj1 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj2 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj3") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj3 = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        
        const cnstVal1 = 33333 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        const cnstVal3 = 200000 * slideFactor;
        const cnstVal4 = 400000 * slideFactor;
        
        const ss = Math.min(w, h);
        
        let a3 = adj3 < 0 ? 0 : (adj3 > cnstVal1 ? cnstVal1 : adj3);
        const maxAdj1 = cnstVal2 - a3;
        let a1 = adj1 < 0 ? 0 : (adj1 > maxAdj1 ? maxAdj1 : adj1);
        const w1 = w / 2 - w / 32;
        const maxAdj2 = cnstVal2 * w1 / ss;
        let a2 = adj2 < 0 ? 0 : (adj2 > maxAdj2 ? maxAdj2 : adj2);
        const x1 = ss * a2 / cnstVal2;
        const x4 = w - x1;
        const dy1 = h * a1 / cnstVal3;
        const dy2 = h * a3 / -cnstVal3;
        const ly1 = h / 2 + dy2 - dy1;
        const ry4 = h / 2 + dy1 - dy2;
        const ly2 = ly1 + dy1;
        const ry3 = h - ly2;
        const ly4 = ly2 * 2;
        const ry1 = h - ly4;
        const ly3 = ly4 - ly1;
        const ry2 = h - ly3;
        const hR = a3 * ss / cnstVal4;
        const x2 = w / 2 - w / 32;
        const x3 = w / 2 + w / 32;
        const y1 = ly1 + hR;
        const y2 = ry2 - hR;
        
        const d = "M" + 0 + "," + ly2 +
            "L" + x1 + "," + 0 +
            "L" + x1 + "," + ly1 +
            "L" + w / 2 + "," + ly1 +
            PPTXShapeUtils.shapeArc(w / 2, y1, w / 32, hR, 270, 450, false).replace("M", "L") +
            PPTXShapeUtils.shapeArc(w / 2, y2, w / 32, hR, 270, 90, false).replace("M", "L") +
            "L" + x4 + "," + ry2 +
            "L" + x4 + "," + ry1 +
            "L" + w + "," + ry3 +
            "L" + x4 + "," + h +
            "L" + x4 + "," + ry4 +
            "L" + w / 2 + "," + ry4 +
            PPTXShapeUtils.shapeArc(w / 2, ry4 - hR, w / 32, hR, 90, 180, false).replace("M", "L") +
            "L" + x2 + "," + ly3 +
            "L" + x1 + "," + ly3 +
            "L" + x1 + "," + ly4 +
            ` zM` + x3 + "," + y1 +
            "L" + x3 + "," + ry2 +
            "M" + x2 + "," + y2 +
            "L" + x2 + "," + ly3;
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "ribbon":
      case "ribbon2": {
        // 丝带形状处理
        let adj1 = (shapType == "ribbon2") ? 16667 * slideFactor : 16667 * slideFactor;
        let adj2 = 50000 * slideFactor;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj1 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj2 = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        
        const cnstVal1 = 25000 * slideFactor;
        const cnstVal2 = 33333 * slideFactor;
        const cnstVal3 = 75000 * slideFactor;
        const cnstVal4 = 100000 * slideFactor;
        const cnstVal5 = 200000 * slideFactor;
        const cnstVal6 = 400000 * slideFactor;
        
        const t = 0, l = 0, b = h, r = w;
        const wd8 = w / 8;
        const wd32 = w / 32;
        
        let a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
        let a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
        
        const x10 = r - wd8;
        const dx2 = w * a2 / cnstVal5;
        const x2 = w / 2 - dx2;
        const x9 = w / 2 + dx2;
        const x3 = x2 + wd32;
        const x8 = x9 - wd32;
        const x5 = x2 + wd8;
        const x6 = x9 - wd8;
        const x4 = x5 - wd32;
        const x7 = x6 + wd32;
        const hR = h * a1 / cnstVal6;
        
        let d = "";
        if (shapType == "ribbon2") {
          const dy1 = h * a1 / cnstVal5;
          const y1 = b - dy1;
          const dy2 = h * a1 / cnstVal4;
          const y2 = b - dy2;
          const y4 = t + dy2;
          const y3 = (y4 + b) / 2;
          const y6 = b - hR;
          const y7 = y1 - hR;
          
          d = "M" + l + "," + b +
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
        } else {
          const y1 = h * a1 / cnstVal5;
          const y2 = h * a1 / cnstVal4;
          const y4 = b - y2;
          const y3 = y4 / 2;
          const y5 = b - hR;
          const y6 = y2 - hR;
          
          d = "M" + l + "," + t +
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
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "moon": {
        // 月牙形状处理
        const radius = Math.min(w, h) * 0.4;
        const shift = radius * 0.5; // 月亮的偏移量
        
        const d = "M" + w / 2 + "," + (h / 2 - radius) + // 顶部点
              PPTXShapeUtils.shapeArc(w / 2, h / 2, radius, radius, 270, 630, false).replace("M", "L") + // 外弧
              " L" + (w / 2 + shift) + "," + h / 2 + // 移动到内圆起点
              PPTXShapeUtils.shapeArc(w / 2 + shift, h / 2, radius * 0.7, radius * 0.7, 270, -90, false).replace("M", "L") + // 内弧
              " Z"; // 关闭路径
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "sun": {
        // 太阳形状处理
        let adj1 = 25000 * slideFactor;
        const cnstVal1 = 12500 * slideFactor;
        const cnstVal2 = 46875 * slideFactor;
        
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        let a1;
        if (adj1 < cnstVal1) a1 = cnstVal1
        else if (adj1 > cnstVal2) a1 = cnstVal2
        else a1 = adj1
        
        const cnstVa3 = 50000 * slideFactor;
        const cnstVa4 = 100000 * slideFactor;
        const g0 = cnstVa3 - a1;
        const g1 = g0 * (30274 * slideFactor) / (32768 * slideFactor);
        const g2 = g0 * (12540 * slideFactor) / (32768 * slideFactor);
        const g3 = g1 + cnstVa3;
        const g4 = g2 + cnstVa3;
        const g5 = cnstVa3 - g1;
        const g6 = cnstVa3 - g2;
        const g7 = g0 * (23170 * slideFactor) / (32768 * slideFactor);
        const g8 = cnstVa3 + g7;
        const g9 = cnstVa3 - g7;
        const g10 = g5 * 3 / 4;
        const g11 = g6 * 3 / 4;
        const g12 = g10 + 3662 * slideFactor;
        const g13 = g11 + 36620 * slideFactor;
        const g14 = g11 + 12500 * slideFactor;
        const g15 = cnstVa4 - g10;
        const g16 = cnstVa4 - g12;
        const g17 = cnstVa4 - g13;
        const g18 = cnstVa4 - g14;
        const ox1 = w * (18436 * slideFactor) / (21600 * slideFactor);
        const oy1 = h * (3163 * slideFactor) / (21600 * slideFactor);
        const ox2 = w * (3163 * slideFactor) / (21600 * slideFactor);
        const oy2 = h * (18436 * slideFactor) / (21600 * slideFactor);
        const x8 = w * g8 / cnstVa4;
        const x9 = w * g9 / cnstVa4;
        const x10 = w * g10 / cnstVa4;
        const x12 = w * g12 / cnstVa4;
        const x13 = w * g13 / cnstVa4;
        const x14 = w * g14 / cnstVa4;
        const x15 = w * g15 / cnstVa4;
        const x16 = w * g16 / cnstVa4;
        const x17 = w * g17 / cnstVa4;
        const x18 = w * g18 / cnstVa4;
        const x19 = w * a1 / cnstVa4;
        const wR = w * g0 / cnstVa4;
        const hR = h * g0 / cnstVa4;
        const y8 = h * g8 / cnstVa4;
        const y9 = h * g9 / cnstVa4;
        const y10 = h * g10 / cnstVa4;
        const y12 = h * g12 / cnstVa4;
        const y13 = h * g13 / cnstVa4;
        const y14 = h * g14 / cnstVa4;
        const y15 = h * g15 / cnstVa4;
        const y16 = h * g16 / cnstVa4;
        const y17 = h * g17 / cnstVa4;
        const y18 = h * g18 / cnstVa4;
        
        const d = "M" + w + "," + h / 2 +
              " L" + x15 + "," + y18 +
              " L" + x15 + "," + y14 +
              "z M" + ox1 + "," + oy1 +
              " L" + x16 + "," + y17 +
              " L" + x13 + "," + y12 +
              "z M" + w / 2 + "," + 0 +
              " L" + x18 + "," + y10 +
              " L" + x14 + "," + y10 +
              "z M" + ox2 + "," + oy1 +
              " L" + x17 + "," + y12 +
              " L" + x12 + "," + y17 +
              "z M" + 0 + "," + h / 2 +
              " L" + x10 + "," + y14 +
              " L" + x10 + "," + y18 +
              "z M" + ox2 + "," + oy2 +
              " L" + x12 + "," + y13 +
              " L" + x17 + "," + y16 +
              "z M" + w / 2 + "," + h +
              " L" + x14 + "," + y15 +
              " L" + x18 + "," + y15 +
              "z M" + ox1 + "," + oy2 +
              " L" + x13 + "," + y16 +
              " L" + x16 + "," + y13 +
              " z M" + x19 + "," + h / 2 +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L") +
              " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "custom":
        if (custShapType !== undefined) {
          // 对于自定义形状，使用适当的处理方法
          result += "<!-- Custom shape: " + custShapType + " -->";
        }
        break;
      case "flowChartMagneticDisk":
      case "flowChartMagneticDrum":
      case "flowChartMagneticTape": {
        // 磁盘、磁鼓、磁带流程图形状
        let adj = 50000 * slideFactor; // 对于磁盘和磁鼓，使用固定值
        
        if (shapType !== "flowChartMagneticDisk" && shapType !== "flowChartMagneticDrum") {
          // 对于磁带，使用默认值并可能从节点获取
          adj = 25000 * slideFactor;
          const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
          if (shapAdjst !== undefined) {
            adj = parseInt(shapAdjst.substr(4)) * slideFactor;
          }
        }
        
        const ss = Math.min(w, h);
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 200000 * slideFactor;
        
        let a, dr, iwd2, ihd2, ang, ct, st, m, n, drd2, dang, dang2, swAng, stAng1, stAng2, ct1, st1, m1, n1, dx1, dy1, x1, y1, x2, y2, stAng1deg, stAng2deg, swAng2deg;
        
        if (adj < 0) a = 0;
        else if (adj > cnstVal1) a = cnstVal1;
        else a = adj;
        
        dr = ss * a / cnstVal2;
        iwd2 = w / 2 - dr;
        ihd2 = h / 2 - dr;
        ang = Math.atan(h / w);
        ct = ihd2 * Math.cos(ang);
        st = iwd2 * Math.sin(ang);
        m = Math.sqrt(ct * ct + st * st); // "mod ct st 0"
        n = iwd2 * ihd2 / m;
        drd2 = dr / 2;
        dang = Math.atan(drd2 / n);
        dang2 = dang * 2;
        swAng = -Math.PI + dang2;
        stAng1 = ang - dang;
        stAng2 = stAng1 - Math.PI;
        ct1 = ihd2 * Math.cos(stAng1);
        st1 = iwd2 * Math.sin(stAng1);
        m1 = Math.sqrt(ct1 * ct1 + st1 * st1); // "mod ct1 st1 0"
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
        
        const d = "M" + 0 + "," + h / 2 +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L") +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L") +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L") +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L") +
              ` zM` + x1 + "," + y1 +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace("M", "L") +
              ` zM` + x2 + "," + y2 +
              PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace("M", "L") +
              " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "'  fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartOfflineStorage": {
        // 离线存储流程图形状
        const d = "M" + w * 0.1 + "," + h + " L0,0 L" + w * 0.9 + "," + 0 + " L" + w + "," + h + " Z";
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "flowChartOffpageConnector": {
        // 跨页连接器流程图形状
        const points = PPTXArrowShapes.genDownArrow(w, h, node, slideFactor).replace("polygon points='", "").replace("'", "");
        result += " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor as string) : "url(#imgPtrn_" + shpId + ")") +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "lightningBolt": {
        // 闪电 bolt 形状
        const centerX = w / 2;
        const centerY = h / 2;
        
        // 简化闪电形状
        const d = "M" + w * 0.1 + "," + h * 0.1 + 
              " L" + w * 0.5 + "," + h * 0.1 + 
              " L" + w * 0.4 + "," + h * 0.5 + 
              " L" + w * 0.9 + "," + h * 0.2 + 
              " L" + w * 0.6 + "," + h * 0.9 + 
              " L" + w * 0.5 + "," + h * 0.5 + 
              " L" + w * 0.1 + "," + h * 0.8 + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "plaque": {
        // 盾牌形状
        let adj1 = 16667 * slideFactor;
        const cnstVal1 = 50000 * slideFactor;
        const cnstVal2 = 100000 * slideFactor;
        
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj1 = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        let a1, x1, x2, y2;
        if (adj1 < 0) a1 = 0;
        else if (adj1 > cnstVal1) a1 = cnstVal1;
        else a1 = adj1;
        
        x1 = a1 * (Math.min(w, h)) / cnstVal2;
        x2 = w - x1;
        y2 = h - x1;
        
        const d = "M0," + x1 +
              PPTXShapeUtils.shapeArc(0, 0, x1, x1, 90, 0, false).replace("M", "L") +
              " L" + x2 + "," + 0 +
              PPTXShapeUtils.shapeArc(w, 0, x1, x1, 180, 90, false).replace("M", "L") +
              " L" + w + "," + y2 +
              PPTXShapeUtils.shapeArc(w, h, x1, x1, 270, 180, false).replace("M", "L") +
              " L" + x1 + "," + h +
              PPTXShapeUtils.shapeArc(0, h, x1, x1, 0, -90, false).replace("M", "L") + " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path   d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "nonIsoscelesTrapezoid": {
        // 非等腰梯形
        let adj1 = 12500 * slideFactor;
        let adj2 = 12500 * slideFactor;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj1 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj2 = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        
        const x1 = adj1 * w / (100000 * slideFactor);
        const x2 = w - adj2 * w / (100000 * slideFactor);
        
        const d = "M0,0 L" + x1 + "," + h + " L" + x2 + "," + h + " L" + w + ",0 Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "folderCorner": {
        // 文件夹角标形状
        const cornerSize = Math.min(w * 0.2, h * 0.2);
        
        const d = "M0,0 L" + cornerSize + ",0 L0," + cornerSize + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "funnel": {
        // 漏斗形状
        const neckWidth = w * 0.3;
        const neckHeight = h * 0.3;
        
        const d = "M0,0 L" + w + ",0 L" + (w - neckWidth) / 2 + "," + (h - neckHeight) + 
              " L" + (w - neckWidth) / 2 + "," + h + " L" + (w + neckWidth) / 2 + "," + h + 
              " L" + (w + neckWidth) / 2 + "," + (h - neckHeight) + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "halfFrame": {
        // 半框架形状
        const frameWidth = Math.min(w * 0.1, h * 0.1);
        
        const d = "M0,0 L" + w + ",0 L" + w + "," + h + " L0," + h + " L0," + frameWidth + 
              " L" + frameWidth + "," + frameWidth + " L" + frameWidth + "," + (h - frameWidth) + 
              " L" + (w - frameWidth) + "," + (h - frameWidth) + " L" + (w - frameWidth) + "," + frameWidth + 
              " L" + frameWidth + "," + frameWidth + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "diagStripe": {
        // 对角条纹形状
        let adj = 50000 * slideFactor;
        const cnstVal1 = 100000 * slideFactor;
        
        const shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
        if (shapAdjst !== undefined) {
          adj = parseInt(shapAdjst.substr(4)) * slideFactor;
        }
        
        let a;
        if (adj < 0) a = 0;
        else if (adj > cnstVal1) a = cnstVal1;
        else a = adj;
        
        const stripeWidth = w * 0.2; // 固定条纹宽度
        const d = "M0," + (h / 2 - stripeWidth / 2) + " L" + w + "," + (h / 2 + stripeWidth / 2) + 
              " L" + w + "," + (h / 2 + stripeWidth / 2 + 10) + " L0," + (h / 2 - stripeWidth / 2 + 10) + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "cornerTabs": {
        // 角标签形状
        const tabSize = Math.min(w * 0.2, h * 0.2);
        
        const d = "M0,0 L" + tabSize + ",0 L" + tabSize + "," + tabSize + " L0," + tabSize + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "plaqueTabs": {
        // 盾形标签形状
        const tabHeight = h * 0.1;
        
        const d = "M0,0 L" + w + ",0 L" + w + "," + h + " L0," + h + " L0," + tabHeight + 
              " L" + w + "," + tabHeight + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "squareTabs": {
        // 方形标签形状
        const tabHeight = h * 0.1;
        const tabSpacing = w / 4;
        
        const d = "M0,0 L" + w + ",0 L" + w + "," + h + " L0," + h + " Z" + 
              "M0," + (h - tabHeight) + " L" + tabSpacing + "," + (h - tabHeight) + " L" + tabSpacing + "," + h + 
              "M" + (tabSpacing * 2) + "," + (h - tabHeight) + " L" + (tabSpacing * 3) + "," + (h - tabHeight) + " L" + (tabSpacing * 3) + "," + h + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "doubleWave": {
        // 双波浪形状
        let adj1 = 6250 * slideFactor;
        let adj2 = 0;
        
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
          for (let i = 0; i < shapAdjst_ary.length; i++) {
            const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj1 = parseInt(sAdj.substr(4)) * slideFactor;
            } else if (sAdj_name == "adj2") {
              const sAdj = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
              adj2 = parseInt(sAdj.substr(4)) * slideFactor;
            }
          }
        }
        
        const cnstVal1 = 12500 * slideFactor;
        const cnstVal2 = -10000 * slideFactor;
        const cnstVal3 = 50000 * slideFactor;
        const cnstVal4 = 100000 * slideFactor;
        
        const t = 0, l = 0, b = h, r = w, wd8 = w / 8, wd32 = w / 32;
        
        let a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
        let a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
        
        const y1 = h * a1 / cnstVal4;
        const dy2 = y1 * 10 / 3;
        const y2 = y1 - dy2;
        const y3 = y1 + dy2;
        const y4 = b - y1;
        const y5 = y4 - dy2;
        const y6 = y4 + dy2;
        const of2 = w * a2 / cnstVal3;
        let dx2 = (of2 > 0) ? 0 : of2;
        const x2 = l - dx2;
        let dx8 = (of2 > 0) ? of2 : 0;
        const x8 = r - dx8;
        const dx3 = (dx2 + x8) / 6;
        const x3 = x2 + dx3;
        const dx4 = (dx2 + x8) / 3;
        const x4 = x2 + dx4;
        const x5 = (x2 + x8) / 2;
        const x6 = x5 + dx3;
        const x7 = (x6 + x8) / 2;
        const x9 = l + dx8;
        const x15 = r + dx2;
        const x10 = x9 + dx3;
        const x11 = x9 + dx4;
        const x12 = (x9 + x15) / 2;
        const x13 = x12 + dx3;
        const x14 = (x13 + x15) / 2;
        
        const d = "M" + x2 + "," + y1 +
              " C" + x3 + "," + y2 + " " + x4 + "," + y3 + " " + x5 + "," + y1 +
              " C" + x6 + "," + y2 + " " + x7 + "," + y3 + " " + x8 + "," + y1 +
              " L" + x15 + "," + y4 +
              " C" + x14 + "," + y6 + " " + x13 + "," + y5 + " " + x12 + "," + y4 +
              " C" + x11 + "," + y6 + " " + x10 + "," + y5 + " " + x9 + "," + y4 +
              " z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      case "ellipseRibbon":
      case "ellipseRibbon2": {
        // 椭圆形丝带形状
        const ribbonWidth = w * 0.3;
        const ribbonPosition = (w - ribbonWidth) / 2;
        
        const d = "M0,0 A" + w / 2 + "," + h / 2 + " 0 1,1 " + w + ",0 A" + w / 2 + "," + h / 2 + " 0 1,1 0,0" +
              " M" + ribbonPosition + "," + (h * 0.2) + " L" + (ribbonPosition + ribbonWidth) + "," + (h * 0.2) + 
              " L" + (ribbonPosition + ribbonWidth) + "," + (h * 0.8) + " L" + ribbonPosition + "," + (h * 0.8) + " Z";
        
        const fillAttr = PPTXShapeContainer.getFillAttrFromFlags(fillColor as string, imgFillFlg, grndFillFlg, shpId);
        result += "<path d='" + d + "' fill='" + fillAttr +
          "' stroke='" + border!.color + "' stroke-width='" + border!.width + "' stroke-dasharray='" + border!.strokeDasharray + "' />";
        break;
      }
      default:
        // 处理未识别的形状类型
        console.warn(`Unknown shape type: ${shapType}`);
        break;
    }
  } else if (custShapType !== undefined) {
    // 处理自定义形状
    const svgCssName = "_svg_css_" + (Object.keys(styleTable).length + 1) + "_"  + Math.floor(Math.random() * 1001);
    const effectsClassName = svgCssName + "_effects";
    let svgStyle = PPTXUtils.getPosition(slideXfrmNode, pNode, undefined, undefined, sType) +
        PPTXUtils.getSize(slideXfrmNode, undefined, undefined) +
        " z-index: " + order + `;transform: rotate(` + ((rotate !== undefined) ? rotate : 0) + "deg)" + flip + `;`;
    // 如果有阴影效果，添加到SVG样式中
    if (oShadowSvgUrlStr && oShadowSvgUrlStr !== "") {
        svgStyle += oShadowSvgUrlStr.replace('filter:', '');
    }
    result += "<svg class='drawing " + svgCssName + " " + effectsClassName + " ' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name + `' style='` + svgStyle + `'>`;
    
    // 获取填充和边框
    fillColor = PPTXShapeFillsUtils.getShapeFill(node, pNode, true, warpObj, source);
    border = PPTXStyleManager.getBorder(node, pNode, true, "shape", warpObj);
    
    // 检查填充类型
    let clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(node, ["p:spPr"]));
    if (clrFillType == "GROUP_FILL") {
      clrFillType = PPTXColorUtils.getFillType(PPTXUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
    }
    
    if (clrFillType == "GRADIENT_FILL") {
      grndFillFlg = true;
    } else if (clrFillType == "PIC_FILL") {
      imgFillFlg = true;
    } else if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
        (custShapType == "arc" ||
          custShapType == "bracketPair" ||
          custShapType == "bracePair" ||
          custShapType == "leftBracket" ||
          custShapType == "leftBrace" ||
          custShapType == "rightBrace" ||
          custShapType == "rightBracket")) { 
      fillColor = "none";
    }
    
    result += '<defs>';
    
    if (grndFillFlg) {
      const color_arry = (fillColor as FillColor).color;
      const angl = (fillColor as FillColor).rot + 90;
      const svgGrdnt = PPTXShapeFillsUtils.getSvgGradient(w, h, angl, color_arry, shpId);
      result += svgGrdnt;
    } else if (imgFillFlg) {
      const imgFill = typeof fillColor === 'object' && (fillColor as FillColor).img ? (fillColor as FillColor).img : fillColor;
      const svgBgImg = PPTXShapeFillsUtils.getSvgImagePattern(node, imgFill, shpId, warpObj);
      result += svgBgImg;
    }
    
    // 处理箭头标记
    headEndNodeAttrs = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
    tailEndNodeAttrs = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
    
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
      const triangleMarker = "<marker id='markerTriangle_" + shpId + "' viewBox='0 0 10 10' refX='1' refY='5' markerWidth='5' markerHeight='5' stroke='" + border!.color + "' fill='" + border!.color +
        "' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>";
      result += triangleMarker;
    }
    
    result += '</defs>';
    
    result += "<!-- Custom shape: " + custShapType + " -->";
    result += "</svg>";
  }

  result += "</svg>";

  return result;
}