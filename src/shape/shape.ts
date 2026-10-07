/**
 * 形状渲染主模块
 * 
 * 这是 PPTX 形状渲染的核心模块，负责处理所有 PowerPoint 预设形状的 SVG 生成。
 * 
 * 模块职责:
 * - 坐标变换和尺寸计算
 * - 形状类型识别和路由
 * - 基础几何形状的 SVG 生成（矩形、圆形、三角形等）
 * - 协调各子模块（箭头、星形、括号、饼图等）
 * 
 * 结构说明:
 * - 该模块是一个 IIFE，导出 PPTXShapeUtils 对象
 * - genShape() 是主入口函数，处理单个形状的完整渲染流程
 * - 使用大量的 switch-case 语句处理不同形状类型
 * - 复杂形状已拆分到独立子模块（arrow-shapes.js, star-shapes.js 等）
 * 
 * 注意事项:
 * - 代码量较大（4875行），包含约 208 个形状类型
 * - 使用 ES5 语法以保持兼容性
 * - 变量命名使用匈牙利命名法（如 shpId, imgFillFlg, grndFillFlg）
 * 
 * @module shape/shape
 */

import { PPTXXmlUtils } from '../utils/xml';
import { PPTXStyleUtils } from '../utils/style';
import { PPTXTextUtils } from '../utils/text';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR, SHADOW_SIGMA_RATIO, GLOW_DILATE_RATIO, GLOW_SIGMA_RATIO } from '../core/constants';
import {
    polarToCartesian,
    shapeArc,
    shapeArcAlt,
    shapeSnipRoundRect,
    shapeSnipRoundRectAlt,
    shapePie,
    shapeGear
} from './path-generators';
import { renderCustomShape } from './custom-shape';
import { renderStar, isStar } from './star-shapes';
import { renderMathSymbol, isMathSymbol } from './math-symbols';
import { renderBracket, isBracket } from './bracket-shapes';
import { renderMiscShape, isMiscShape } from './misc-shapes';
import { renderPieShape, isPieShape } from './pie-shapes';
import { renderArrow, isArrow } from './arrow-shapes';
import {
    RECT_SHAPES,
    ROUND_RECT_SHAPES,
    SNIP_RECT_SHAPES,
    FLOWCHART_SHAPES,
    ACTION_BUTTONS,
    BASIC_SHAPES,
    STAR_SHAPES,
    ARROW_SHAPES,
    CALLOUT_SHAPES,
    BRACKET_SHAPES,
    SPECIAL_SHAPES,
    getShapeCategory,
    isComplexShape
} from './shape-categories';
import { renderActionButton, isActionButton } from './action-buttons';
import type { XmlNode, WarpObject, ParseSettings } from '../core/types';

/** PPTXShapeUtils 模块对外接口 */
interface ShapeUtilsModule {
    /** 重新导出的路径生成函数 */
    shapeArc: typeof shapeArc;
    shapeArcAlt: typeof shapeArcAlt;
    shapePie: typeof shapePie;
    shapeGear: typeof shapeGear;
    shapeSnipRoundRect: typeof shapeSnipRoundRect;
    shapeSnipRoundRectAlt: typeof shapeSnipRoundRectAlt;
    polarToCartesian: typeof polarToCartesian;
    /** 核心形状生成函数（定义在 IIFE 内部，此处显式声明签名） */
    genShape: (node: XmlNode | undefined, pNode: XmlNode | undefined, slideLayoutSpNode: XmlNode | undefined, slideMasterSpNode: XmlNode | undefined, id: string | number | undefined, name: string | undefined, idx: number | undefined, type: string, order: string | number | undefined, warpObj: WarpObject, isUserDrawnBg: boolean | undefined, sType: string, source: string, settings: ParseSettings) => Promise<string>;
}

export const PPTXShapeUtils: ShapeUtilsModule = (function() {
    /**
     * 辅助函数：生成形状的 data- 属性字符串
     * @param {Object} node - 节点
     * @param {Object} slideXfrmNode - 变换节点
     * @param {string} id - ID
     * @param {string} name - 名称
     * @param {string} idx - 索引
     * @param {string} type - 类型
     * @param {number} rotate - 旋转角度
     * @param {string} sType - 形状类型
     * @returns {string} data- 属性字符串
     */
    function genShapeDataAttributes(node: XmlNode | undefined, slideXfrmNode: XmlNode | undefined, id: string | number | undefined, name: string | undefined, idx: number | undefined, type: string, rotate: number | undefined, sType: string) {
        let dataAttrs = '';
        
        // 提取位置和尺寸信息
        let offX = 0, offY = 0, extCx = 0, extCy = 0, rot = 0, flipH = 0, flipV = 0;
        if (slideXfrmNode !== undefined) {
            if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                offX = slideXfrmNode['a:off'].attrs.x || 0;
                offY = slideXfrmNode['a:off'].attrs.y || 0;
            }
            if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                extCx = slideXfrmNode['a:ext'].attrs.cx || 0;
                extCy = slideXfrmNode['a:ext'].attrs.cy || 0;
            }
            if (slideXfrmNode['attrs']) {
                rot = slideXfrmNode['attrs'].rot || 0;
                flipH = slideXfrmNode['attrs'].flipH || '0';
                flipV = slideXfrmNode['attrs'].flipV || '0';
            }
        }
        
        // 获取形状类型
        const shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
        
        // 构建 data- 属性
        dataAttrs += ` data-node-id="${id || ''}"`;
        dataAttrs += ` data-node-name="${name || ''}"`;
        dataAttrs += ` data-node-idx="${idx || ''}"`;
        dataAttrs += ` data-node-type="${type || ''}"`;
        dataAttrs += ` data-shape-type="${sType || ''}"`;
        dataAttrs += ` data-off-x="${offX}"`;
        dataAttrs += ` data-off-y="${offY}"`;
        dataAttrs += ` data-ext-cx="${extCx}"`;
        dataAttrs += ` data-ext-cy="${extCy}"`;
        dataAttrs += ` data-rotate="${rotate || 0}"`;
        dataAttrs += ` data-flip-h="${flipH}"`;
        dataAttrs += ` data-flip-v="${flipV}"`;
        if (shapType) {
            dataAttrs += ` data-geom-type="${shapType}"`;
        }
        
        return dataAttrs;
    }

    async function genShape(node: XmlNode | undefined, pNode: XmlNode | undefined, slideLayoutSpNode: XmlNode | undefined, slideMasterSpNode: XmlNode | undefined, id: string | number | undefined, name: string | undefined, idx: number | undefined, type: string, order: string | number | undefined, warpObj: WarpObject, isUserDrawnBg: boolean | undefined, sType: string, source: string, settings: ParseSettings) {
            //var dltX = 0;
            //var dltY = 0;
            const xfrmList = ["p:spPr", "a:xfrm"];
            const slideXfrmNode = PPTXXmlUtils.getTextByPathList(node, xfrmList);
            const slideLayoutXfrmNode = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
            const slideMasterXfrmNode = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, xfrmList);

            let result = "";
            const shpId = PPTXXmlUtils.getTextByPathList(node, ["attrs", "order"]);
            const shapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);

            // 初始化3D变换样式
            let transform3dStyle = "";
            //custGeom - Amir
            const custShapType = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);

            let isFlipV = false;
            let isFlipH = false;
            let flip = "";
            const flipVAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]);
            const flipHAttr = PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]);
            if (flipVAttr === "1" || flipVAttr === "true") {
                isFlipV = true;
            }
            if (flipHAttr === "1" || flipHAttr === "true") {
                isFlipH = true;
            }
            if (isFlipH && !isFlipV) {
                flip = " scale(-1,1)"
            } else if (!isFlipH && isFlipV) {
                flip = " scale(1,-1)"
            } else if (isFlipH && isFlipV) {
                flip = " scale(-1,-1)"
            }
            /////////////////////////Amir////////////////////////
            //rotate
            const rotate = PPTXXmlUtils.angleToDegrees(PPTXXmlUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"]));


            let txtRotate;
            const txtXframeNode = PPTXXmlUtils.getTextByPathList(node, ["p:txXfrm"]);
            if (txtXframeNode !== undefined) {
                const txtXframeRot = PPTXXmlUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
                if (txtXframeRot !== undefined) {
                    txtRotate = PPTXXmlUtils.angleToDegrees(txtXframeRot) + 90;
                }
            } else {
                txtRotate = 0;
            }
            
            // Adjust text rotation to compensate for shape flip
            let txtFlip = "";
            if (isFlipV) {
                txtFlip = " scale(1,-1)";
            }
            if (isFlipH) {
                txtFlip = " scale(-1,1)";
            }
            if (isFlipH && isFlipV) {
                txtFlip = " scale(-1,-1)";
            }
            //////////////////////////////////////////////////

            // 处理组合缩放 - 当形状在group-abs类型组合中时需要应用缩放
            let workingXfrmNode = slideXfrmNode;
            let drawW, drawH; // 原始未缩放的尺寸，用于SVG内部绘图

            // 先保存原始尺寸（如果存在）
            if (slideXfrmNode && slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
                const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
                drawW = (originalCx !== undefined && originalCy !== undefined)
                    ? originalCx * SLIDE_FACTOR
                    : undefined;
                drawH = (originalCx !== undefined && originalCy !== undefined)
                    ? originalCy * SLIDE_FACTOR
                    : undefined;
            }

            if (sType === 'group-abs' && warpObj.currentGroupScale && slideXfrmNode) {
                const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

                // 创建缩放后的xfrmNode
                workingXfrmNode = JSON.parse(JSON.stringify(slideXfrmNode));

                // 缩放尺寸
                if (slideXfrmNode['a:ext'] && slideXfrmNode['a:ext'].attrs) {
                    const originalCx = parseInt(slideXfrmNode['a:ext'].attrs.cx);
                    const originalCy = parseInt(slideXfrmNode['a:ext'].attrs.cy);
                    workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                    workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
                }

                // 调整位置(相对于childX/childY)
                if (slideXfrmNode['a:off'] && slideXfrmNode['a:off'].attrs) {
                    const originalOffX = parseInt(slideXfrmNode['a:off'].attrs.x);
                    const originalOffY = parseInt(slideXfrmNode['a:off'].attrs.y);

                    // childX和childY已经是像素值，需要转换为EMU单位来计算
                    const childXEmu = childX / SLIDE_FACTOR;
                    const childYEmu = childY / SLIDE_FACTOR;

                    // 计算相对于childOff的偏移（EMU单位）
                    const relativeX = originalOffX - childXEmu;
                    const relativeY = originalOffY - childYEmu;

                    // 应用缩放，结果为EMU单位
                    workingXfrmNode['a:off'].attrs.x = Math.round(childXEmu + relativeX * scaleX);
                    workingXfrmNode['a:off'].attrs.y = Math.round(childYEmu + relativeY * scaleY);
                }
            }

            if (shapType !== undefined || custShapType !== undefined /*&& slideXfrmNode !== undefined*/) {
                // 使用 workingXfrmNode 而不是 slideXfrmNode，以正确处理 group-abs 情况
                const off = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:off", "attrs"]);
                var x = (off !== undefined) ? parseInt(off["x"]) * SLIDE_FACTOR : 0;
                var y = (off !== undefined) ? parseInt(off["y"]) * SLIDE_FACTOR : 0;

                let ext = PPTXXmlUtils.getTextByPathList(workingXfrmNode, ["a:ext", "attrs"]);

                // Fallback to slideLayoutXfrmNode if workingXfrmNode is undefined or ext is undefined
                if (ext === undefined && slideLayoutXfrmNode !== undefined) {
                    ext = PPTXXmlUtils.getTextByPathList(slideLayoutXfrmNode, ["a:ext", "attrs"]);
                }
                // Fallback to slideMasterXfrmNode if still undefined
                if (ext === undefined && slideMasterXfrmNode !== undefined) {
                    ext = PPTXXmlUtils.getTextByPathList(slideMasterXfrmNode, ["a:ext", "attrs"]);
                }

                var w: any = (ext !== undefined && ext["cx"] !== undefined) ? parseInt(ext["cx"]) * SLIDE_FACTOR : 100;
                var h: any = (ext !== undefined && ext["cy"] !== undefined) ? parseInt(ext["cy"]) * SLIDE_FACTOR : 100;
                w = isNaN(w) ? 100 : w;
                h = isNaN(h) ? 100 : h;

                // 如果drawW/drawH未定义（非group-abs情况），则使用w和h
                if (drawW === undefined) drawW = w;
                if (drawH === undefined) drawH = h;

                // 对于连接器类型，需要特殊处理
                const isConnector = (shapType === 'straightConnector1' || shapType === 'bentConnector2' ||
                                   shapType === 'bentConnector3' || shapType === 'bentConnector4' ||
                                   shapType === 'bentConnector5' || shapType === 'curvedConnector2' ||
                                   shapType === 'curvedConnector3' || shapType === 'curvedConnector4' ||
                                   shapType === 'curvedConnector5');





                const svgCssName = `_svg_css_${(Object.keys(warpObj.styleTable).length + 1)}_${Math.floor(Math.random() * 1001)}`;

                let hasCssEffect = false; // Track if there's a CSS effect (like shadow)
                const effectsClassName = `${svgCssName}_effects`;

                // 对于连接器，当width或height为0时，需要设置最小尺寸
                let svgSizeStyle = "";

                if (isConnector && (w === 0 || h === 0)) {
                    // 设置最小尺寸为strokeWidth的2倍（或至少4px），确保线条可见
                    const strokeWidth = 1.5; // 默认stroke-width，实际可以从border获取
                    const minSize = Math.max(strokeWidth * 2, 4);
                    // SVG容器的尺寸至少为minSize
                    const svgW = (w === 0 || w < minSize) ? minSize : w;
                    const svgH = (h === 0 || h < minSize) ? minSize : h;
                    svgSizeStyle = `width:${svgW}px; height:${svgH}px; overflow: visible;`;
                    // 更新w和h为SVG容器尺寸，这样后续代码会使用正确的尺寸
                    w = svgW;
                    h = svgH;

                } else {
                    svgSizeStyle = `${PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined)} overflow: visible;`;
                }

                // 如果形状在组合中被缩放，SVG内容需要应用缩放
                let svgTransform = `transform: rotate(${((rotate !== undefined) ? rotate : 0)}deg)${flip};`;
                if (sType === 'group-abs' && warpObj.currentGroupScale) {
                    // 对于自定义形状，我们已经在 renderCustomShape 中使用了缩放后的尺寸
                    // 所以不需要在这里应用 SVG transform scale()
                    // 预置形状仍然需要使用 transform scale()
                    if (custShapType === undefined) {
                        const { scaleX, scaleY } = warpObj.currentGroupScale;
                        svgTransform = `transform: rotate(${(rotate !== undefined) ? rotate : 0}deg)${flip} scale(${scaleX},${scaleY});`;
                    }
                }

                const svgTag = `<svg class='drawing ${svgCssName}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}'' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, undefined, undefined, sType)}${svgSizeStyle} z-index: ${order};${svgTransform}'>`;
                result += svgTag;
                result += '<defs>'
                // Fill Color
                var fillColor = await PPTXStyleUtils.getShapeFill(node, pNode, true, warpObj, source);

                var grndFillFlg: any = false;
                var imgFillFlg: any = false;
                let clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
                if (clrFillType == "GROUP_FILL") {
                    clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr"]));
                }
                // if (clrFillType == "") {
                //     var clrFillType = PPTXStyleUtils.getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:style","a:fillRef"]));
                // }

                /////////////////////////////////////////                    
                if (clrFillType == "GRADIENT_FILL") {
                    grndFillFlg = true;
                    const color_arry = fillColor.color;
                    // getGradientFill 返回的 rot 是给 CSS linear-gradient() 用的（rot = a:lin/@ang + 90）；
                    // SVGangle 需要 a:lin/@ang 的原始角度（顺时针、0° 指向右）
                    const angl = fillColor.rot - 90;
                    const svgGrdnt = PPTXStyleUtils.getSvgGradient(w, h, angl, color_arry, shpId);
                    //fill="url(#linGrd)"
                    //console.log("genShape: svgGrdnt: ", svgGrdnt)
                    result += svgGrdnt;

                } else if (clrFillType == "PIC_FILL") {
                    imgFillFlg = true;
                    const svgBgImg = PPTXStyleUtils.getSvgImagePattern(node, fillColor, shpId, warpObj);
                    //fill="url(#imgPtrn)"
                    //console.log(svgBgImg)
                    result += svgBgImg;
                } else if (clrFillType == "PATTERN_FILL") {
                    // 图案填充：生成 SVG <pattern> 平铺定义，形状用 fill="url(#pattPtrn_xx)" 引用。
                    // 不再走 styleTable 的 CSS background —— SVG 元素的 background 不会按路径裁剪，
                    // 且以图案文本为 key 会让多个形状互相覆盖（后写入者覆盖前者）。
                    // 图案可能来自形状自身，也可能继承自所属组合（a:grpFill → p:grpSpPr）。
                    const ownPattFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:pattFill"]);
                    const pattHost = (ownPattFill !== undefined || pNode === undefined)
                        ? node
                        : { "a:pattFill": PPTXXmlUtils.getTextByPathList(pNode, ["p:grpSpPr", "a:pattFill"]) };
                    const svgPattFill = PPTXStyleUtils.getSvgPatternFill(pattHost, shpId, warpObj);
                    result += svgPattFill;
                    fillColor = (svgPattFill === "") ? "none" : `url(#pattPtrn_${shpId})`;
                } else {
                    if (clrFillType != "SOLID_FILL" && clrFillType != "PATTERN_FILL" &&
                        (shapType == "arc" ||
                            shapType == "bracketPair" ||
                            shapType == "bracePair" ||
                            shapType == "leftBracket" ||
                            shapType == "leftBrace" ||
                            shapType == "rightBrace" ||
                            shapType == "rightBracket")) { //Temp. solution  - TODO
                        fillColor = "none";
                    }
                }
                // Border Color
                var border = PPTXStyleUtils.getBorder(node, pNode, true, "shape", warpObj);

                var headEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
                var tailEndNodeAttrs = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
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
                
                // 处理3D效果
                const scene3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
                const sp3d = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
                
                if (scene3d || sp3d) {
                    transform3dStyle = process3DEffects(scene3d, sp3d);
                }
                // Check if there's an effectRef in p:style
                const effectRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:effectRef"]);
                let effectStyleNode = undefined;
                
                if (effectRefNode !== undefined) {
                    const effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
                    if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                        // Access the effect style from the theme
                        let effectStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:effectStyleLst"]["a:effectStyle"];
                        if (effectStyleLst !== undefined) {
                            // Ensure effectStyleLst is an array
                            if (!Array.isArray(effectStyleLst)) {
                                effectStyleLst = [effectStyleLst];
                            }
                            // Convert effectIdx to number and use as array index
                            // effectRef idx is 0-based (idx="0" refers to first effectStyle)
                            idx = Number(effectIdx);
                            // Handle idx out of range
                            if (effectStyleLst.length > 0) {
                                if (idx >= 0 && idx < effectStyleLst.length) {
                                    effectStyleNode = effectStyleLst[idx];
                                } else {
                                    // When idx is out of range, try to find an effectStyle with shadow
                                    // Start from the end of the list and work backwards
                                    for (var i = effectStyleLst.length - 1; i >= 0; i--) {
                                        const testEffectStyle = effectStyleLst[i];
                                        const hasShadow = PPTXXmlUtils.getTextByPathList(testEffectStyle, ["a:effectLst", "a:outerShdw"]);
                                        if (hasShadow !== undefined) {
                                            effectStyleNode = testEffectStyle;
                                            break;
                                        }
                                    }
                                    // If no shadow found, use modulo to wrap around
                                    if (effectStyleNode === undefined) {
                                        idx = idx % effectStyleLst.length;
                                        if (idx < 0) idx += effectStyleLst.length;
                                        effectStyleNode = effectStyleLst[idx];
                                    }
                                }
                            }
                        }
                    }
                }
                
                //////////////////////////////outerShdw///////////////////////////////////////////
                //not support sizing the shadow
                let outerShdwNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:outerShdw"]);
                
                // If no direct outerShdw, check from effectStyle
                if (outerShdwNode === undefined && effectStyleNode !== undefined) {
                    outerShdwNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:outerShdw"]);
                }

                var oShadowSvgUrlStr: any = ""
                // Check if outerShdwNode exists and has valid shadow attributes
                // A valid shadow should have at least dist defined with a non-zero value
                let hasOuterShadow = false;
                if (outerShdwNode && typeof outerShdwNode === 'object' && !Array.isArray(outerShdwNode)) {
                    // Check that outerShdwNode is not empty and is actually a valid outerShdw node
                    const nodeKeys = Object.keys(outerShdwNode);
                    if (nodeKeys.length > 0) {
                        var attrs = outerShdwNode.attrs;
                        // A valid outerShdw node should have an attrs object with shadow properties
                        if (attrs && typeof attrs === 'object') {
                            const { dist: distVal, blurRad: blurRadVal } = attrs;
                            // Only consider it a valid shadow if dist is defined and non-zero
                            // Also check if at least one of the required shadow attributes is present
                            const hasShadowAttrs = (distVal !== undefined || blurRadVal !== undefined ||
                                                 attrs.dir !== undefined || attrs.sx !== undefined ||
                                                 attrs.sy !== undefined || attrs.algn !== undefined);
                            hasOuterShadow = hasShadowAttrs && (distVal !== undefined && distVal !== "" && distVal !== "0" && distVal !== 0);
                        }
                    }
                }

                // Check if shape has 3D effects (sp3d or scene3d)
                // Only disable shadow if 3D effects are defined directly on the shape (p:spPr)
                // 3D effects from effectStyle (via effectRef) should not disable the shadow
                const sp3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:sp3d"]);
                const scene3dNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:scene3d"]);
                // Check if outerShdw is from effectStyle
                const shadowFromEffectStyle = (outerShdwNode !== undefined && effectStyleNode !== undefined);
                if ((sp3dNode !== undefined || scene3dNode !== undefined) && !shadowFromEffectStyle) {
                    // Disable shadow when 3D effects are present on the shape itself
                    hasOuterShadow = false;
                }


                // 阴影与发光统一用一个 SVG 滤镜实现：
                //  · CSS drop-shadow 的 blur 半径各浏览器对 σ 的解释不一致（实测 Chromium 近似 σ ≈ 半径），
                //    改用 feGaussianBlur 的 stdDeviation 表达才可控且跨浏览器一致；
                //  · 两者基于同一份 SourceAlpha 计算，避免链式 filter 让发光把阴影也算进去。
                const fxId = `fx_${svgCssName}`;
                let fxPasses = '';
                const fxLayers: string[] = [];
                let fxMargin = 0;

                // 颜色可能带 a:alpha（getSolidFill 返回 8 位 hex），需拆成 flood-color + flood-opacity
                const splitAlphaColor = (rawClr: string) => {
                    const isHex8 = /^[0-9a-fA-F]{8}$/.test(rawClr);
                    return {
                        color: isHex8 ? rawClr.slice(0, 6) : rawClr,
                        opacity: isHex8 ? (parseInt(rawClr.slice(6, 8), 16) / 255).toFixed(3) : '1'
                    };
                };

                if (hasOuterShadow) {
                    const chdwClrNode = PPTXStyleUtils.getSolidFill(outerShdwNode, undefined, undefined, warpObj) || '000000';
                    const outerShdwAttrs = outerShdwNode["attrs"];

                    //var algn = outerShdwAttrs["algn"];
                    var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                    var dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR;//(px) //* (3 / 4); //(pt)
                    //var rotWithShape = outerShdwAttrs["rotWithShape"];
                    var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR) : 0;
                    //var sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                    //var sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                    const vx = dist * Math.sin(dir * Math.PI / 180);
                    const hx = dist * Math.cos(dir * Math.PI / 180);
                    // OOXML 未规定模糊核：σ 按 SHADOW_SIGMA_RATIO 换算（对 WPS 剖面拟合所得）
                    const shdwSigma = blurRad * SHADOW_SIGMA_RATIO;
                    const shdwClr = splitAlphaColor(chdwClrNode);
                    fxPasses += `<feGaussianBlur in="SourceAlpha" stdDeviation="${shdwSigma}" result="shdwBlur"/>`;
                    fxPasses += `<feOffset in="shdwBlur" dx="${hx}" dy="${vx}" result="shdwOff"/>`;
                    fxPasses += `<feFlood flood-color="#${shdwClr.color}" flood-opacity="${shdwClr.opacity}" result="shdwColor"/>`;
                    fxPasses += `<feComposite in="shdwColor" in2="shdwOff" operator="in" result="shdwLayer"/>`;
                    fxLayers.push('shdwLayer');
                    fxMargin = Math.max(fxMargin, dist + shdwSigma * 3);
                }

                //////////////////////////////glow///////////////////////////////////////////////
                let glowNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:glow"]);

                // If no direct glow, check from effectStyle
                if (glowNode === undefined && effectStyleNode !== undefined) {
                    glowNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:glow"]);
                }
                if (glowNode !== undefined) {
                    const glowAttrs = glowNode["attrs"] || {};
                    const glowRad = glowAttrs["rad"] ? (parseInt(glowAttrs["rad"]) * SLIDE_FACTOR) : 0;
                    if (glowRad > 0) {
                        const glowClr = splitAlphaColor(PPTXStyleUtils.getSolidFill(glowNode, undefined, undefined, warpObj) || '000000');
                        // 按 PowerPoint/WPS 的发光算法：轮廓先向外膨胀出一段实心光边，再高斯柔化外缘。
                        // 只对轮廓做高斯模糊的话，尖角处强度只有 ~0.25、整体明显偏淡。
                        const glowDilate = glowRad * GLOW_DILATE_RATIO;
                        const glowSigma = glowRad * GLOW_SIGMA_RATIO;
                        fxPasses += `<feMorphology in="SourceAlpha" operator="dilate" radius="${glowDilate}" result="glowDil"/>`;
                        fxPasses += `<feGaussianBlur in="glowDil" stdDeviation="${glowSigma}" result="glowBlur"/>`;
                        fxPasses += `<feFlood flood-color="#${glowClr.color}" flood-opacity="${glowClr.opacity}" result="glowColor"/>`;
                        fxPasses += `<feComposite in="glowColor" in2="glowBlur" operator="in" result="glowLayer"/>`;
                        // 发光在最底层，阴影次之，形状本体最上
                        fxLayers.unshift('glowLayer');
                        fxMargin = Math.max(fxMargin, glowDilate + glowSigma * 3);
                    }
                }

                // 阴影/发光注册为一条 filter CSS 规则（作用于该形状 SVG 元素）
                if (fxLayers.length > 0) {
                    const margin = Math.ceil(fxMargin) + 2;
                    // 显式指定 sRGB：SVG 滤镜默认在 linearRGB 空间运算，与 PowerPoint/WPS/CSS 的 sRGB 不一致。
                    // 滤镜区域用 userSpaceOnUse 精确给出（按 bbox 百分比在细长图形上会裁掉光晕/阴影）。
                    let fxFilter = `<filter id="${fxId}" filterUnits="userSpaceOnUse" x="${-margin}" y="${-margin}"` +
                        ` width="${w + margin * 2}" height="${h + margin * 2}" color-interpolation-filters="sRGB">`;
                    fxFilter += fxPasses;
                    fxFilter += `<feMerge>${fxLayers.map(layer => `<feMergeNode in="${layer}"/>`).join('')}<feMergeNode in="SourceGraphic"/></feMerge>`;
                    fxFilter += `</filter>`;
                    result += fxFilter;

                    let effectsCss = `filter:url(#${fxId});`;
                    if (effectsCss in warpObj.styleTable) {
                        effectsCss += `do-nothing: ${svgCssName};`;
                    }

                    warpObj.styleTable[effectsCss] = {
                        "name": effectsClassName,
                        "text": effectsCss
                    };

                    // Add effectsClassName to the SVG tag since we have a CSS shadow effect
                    hasCssEffect = true;
                    result = result.replace(`class='drawing ${svgCssName}'`, `class='drawing ${svgCssName} ${effectsClassName}'`);
                }

                //////////////////////////////softEdge///////////////////////////////////////////
                // Soft edge effect - creates a blurred/feathered edge
                let softEdgeNode = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:effectLst", "a:softEdge"]);
                
                // If no direct softEdge, check from effectStyle
                if (softEdgeNode === undefined && effectStyleNode !== undefined) {
                    softEdgeNode = PPTXXmlUtils.getTextByPathList(effectStyleNode, ["a:effectLst", "a:softEdge"]);
                }
                
                var softEdgeFilterStr: any = ""
                if (softEdgeNode !== undefined) {
                    const softEdgeAttrs = softEdgeNode["attrs"];
                    const rad = (softEdgeAttrs["rad"]) ? (parseInt(softEdgeAttrs["rad"]) * SLIDE_FACTOR) : 0;
                    
                    // softEdge effect according to Office Open XML specification:
                    // Applies a Gaussian blur to the edges of the shape
                    // The radius determines how far the blur extends from the edge
                    const softEdgeId = `softedge_${shpId}`;
                    // 同样显式指定 sRGB：滤镜默认在 linearRGB 空间模糊，会让柔化后的颜色偏亮发灰
                    let softEdgeFilter = `<filter id="${softEdgeId}" x="-20%" y="-20%" width="140%" height="140%" color-interpolation-filters="sRGB">`;
                    // Blur the source to create soft edge
                    softEdgeFilter += `<feGaussianBlur in="SourceGraphic" stdDeviation="${rad}" />`;
                    softEdgeFilter += '</filter>';
                    result += softEdgeFilter;
                    softEdgeFilterStr = `filter="url(#${softEdgeId})"`;
                } 
                ////////////////////////////////////////////////////////////////////////////////////////
                if ((headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) ||
                    (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow"))) {
                    // 箭头标记：refX=10 表示箭头尖端与线条端点对齐
                    const triangleMarker = `<marker id='markerTriangle_${shpId}' viewBox='0 0 10 10' refX='10' refY='5' markerWidth='5' markerHeight='5' stroke='${border.color}' fill='${border.color}' orient='auto-start-reverse' markerUnits='strokeWidth'><path d='M 0 0 L 10 5 L 0 10 z' /></marker>`;
                    result += triangleMarker;
                }
                result += '</defs>'
            }
            if (shapType !== undefined && custShapType === undefined) {
                //console.log("shapType: ", shapType)
                switch (shapType) {
                    case "rect":
                    case "flowChartProcess":
                    case "flowChartPredefinedProcess":
                    case "flowChartInternalStorage":
                    case "actionButtonBlank": {
                        result += `<rect x='0' y='0' width='${w}' height='${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr}  />`;

                        if (shapType == "flowChartPredefinedProcess") {
                            result += `<rect x='${w * (1 / 8)}' y='0' width='${w * (6 / 8)}' height='${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        } else if (shapType == "flowChartInternalStorage") {
                            result += ` <polyline points='${w * (1 / 8)} 0,${w * (1 / 8)} ${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                            result += ` <polyline points='0 ${h * (1 / 8)},${w} ${h * (1 / 8)}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        }
                        break;
                    }
                    case "flowChartCollate": {
                        var d: any = `M 0,0 L${w},${0} L${0},${h} L${w},${h} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr} ${softEdgeFilterStr} />`;

                        break;
                    }
                    case "flowChartDocument": {
                        var y1, y2: any, y3, x1;
                        x1 = w * 10800 / 21600;
                        y1 = h * 17322 / 21600;
                        y2 = h * 20172 / 21600;
                        y3 = h * 23922 / 21600;
                        var d: any = `M${0},${0} L${w},${0} L${w},${y1} C${x1},${y1} ${x1},${y3} ${0},${y2} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartMultidocument": {
                        var y1, y2: any, y3, y4, y5, y6, y7, y8, y9, x1, x2: any, x3, x4, x5, x6, x7;
                        y1 = h * 18022 / 21600;
                        y2 = h * 3675 / 21600;
                        y3 = h * 23542 / 21600;
                        y4 = h * 1815 / 21600;
                        y5 = h * 16252 / 21600;
                        y6 = h * 16352 / 21600;
                        y7 = h * 14392 / 21600;
                        y8 = h * 20782 / 21600;
                        y9 = h * 14467 / 21600;
                        x1 = w * 1532 / 21600;
                        x2 = w * 20000 / 21600;
                        x3 = w * 9298 / 21600;
                        x4 = w * 19298 / 21600;
                        x5 = w * 18595 / 21600;
                        x6 = w * 2972 / 21600;
                        x7 = w * 20800 / 21600;
                        var d: any = `M${0},${y2} L${x5},${y2} L${x5},${y1} C${x3},${y1} ${x3},${y3} ${0},${y8} zM${x1},${y2} L${x1},${y4} L${x2},${y4} L${x2},${y5} C${x4},${y5} ${x5},${y6} ${x5},${y6}M${x6},${y4} L${x6},${0} L${w},${0} L${w},${y7} C${x7},${y7} ${x2},${y9} ${x2},${y9}`;

                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "actionButtonBackPrevious":
                    case "actionButtonBeginning":
                    case "actionButtonDocument":
                    case "actionButtonEnd":
                    case "actionButtonForwardNext":
                    case "actionButtonHelp":
                    case "actionButtonHome":
                    case "actionButtonInformation":
                    case "actionButtonMovie":
                    case "actionButtonReturn":
                    case "actionButtonSound": {
                        result += renderActionButton(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt);
                        break;
                    }
                    case "irregularSeal1":
                    case "irregularSeal2": {
                        if (shapType == "irregularSeal1") {
                            var d: any = `M${w * 10800 / 21600},${h * 5800 / 21600} L${w * 14522 / 21600},${0} L${w * 14155 / 21600},${h * 5325 / 21600} L${w * 18380 / 21600},${h * 4457 / 21600} L${w * 16702 / 21600},${h * 7315 / 21600} L${w * 21097 / 21600},${h * 8137 / 21600} L${w * 17607 / 21600},${h * 10475 / 21600} L${w},${h * 13290 / 21600} L${w * 16837 / 21600},${h * 12942 / 21600} L${w * 18145 / 21600},${h * 18095 / 21600} L${w * 14020 / 21600},${h * 14457 / 21600} L${w * 13247 / 21600},${h * 19737 / 21600} L${w * 10532 / 21600},${h * 14935 / 21600} L${w * 8485 / 21600},${h} L${w * 7715 / 21600},${h * 15627 / 21600} L${w * 4762 / 21600},${h * 17617 / 21600} L${w * 5667 / 21600},${h * 13937 / 21600} L${w * 135 / 21600},${h * 14587 / 21600} L${w * 3722 / 21600},${h * 11775 / 21600} L${0},${h * 8615 / 21600} L${w * 4627 / 21600},${h * 7617 / 21600} L${w * 370 / 21600},${h * 2295 / 21600} L${w * 7312 / 21600},${h * 6320 / 21600} L${w * 8352 / 21600},${h * 2295 / 21600} z`;
                        } else if (shapType == "irregularSeal2") {
                            var d: any = `M${w * 11462 / 21600},${h * 4342 / 21600} L${w * 14790 / 21600},${0} L${w * 14525 / 21600},${h * 5777 / 21600} L${w * 18007 / 21600},${h * 3172 / 21600} L${w * 16380 / 21600},${h * 6532 / 21600} L${w},${h * 6645 / 21600} L${w * 16985 / 21600},${h * 9402 / 21600} L${w * 18270 / 21600},${h * 11290 / 21600} L${w * 16380 / 21600},${h * 12310 / 21600} L${w * 18877 / 21600},${h * 15632 / 21600} L${w * 14640 / 21600},${h * 14350 / 21600} L${w * 14942 / 21600},${h * 17370 / 21600} L${w * 12180 / 21600},${h * 15935 / 21600} L${w * 11612 / 21600},${h * 18842 / 21600} L${w * 9872 / 21600},${h * 17370 / 21600} L${w * 8700 / 21600},${h * 19712 / 21600} L${w * 7527 / 21600},${h * 18125 / 21600} L${w * 4917 / 21600},${h} L${w * 4805 / 21600},${h * 18240 / 21600} L${w * 1285 / 21600},${h * 17825 / 21600} L${w * 3330 / 21600},${h * 15370 / 21600} L${0},${h * 12877 / 21600} L${w * 3935 / 21600},${h * 11592 / 21600} L${w * 1172 / 21600},${h * 8270 / 21600} L${w * 5372 / 21600},${h * 7817 / 21600} L${w * 4502 / 21600},${h * 3625 / 21600} L${w * 8550 / 21600},${h * 6382 / 21600} L${w * 9722 / 21600},${h * 1887 / 21600} z`;
                        }
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartTerminator": {
                        var x1, x2: any, y1, cd2: any = 180, cd4: any = 90, c3d4 = 270;
                        x1 = w * 3475 / 21600;
                        x2 = w * 18125 / 21600;
                        y1 = h * 10800 / 21600;
                        //path attrs: w = 21600; h = 21600; 
                        var d: any = `M${x1},${0} L${x2},${0}${PPTXShapeUtils.shapeArcAlt(x2, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace!("M", "L")} L${x1},${h}${PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, cd4 + cd2, false).replace!("M", "L")} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartPunchedTape": {
                        var x1, x1, y1, y2: any, cd2: any = 180;
                        x1 = w * 5 / 20;
                        y1 = h * 2 / 20;
                        y2 = h * 18 / 20;
                        var d: any = `M${0},${y1}${PPTXShapeUtils.shapeArcAlt(x1, y1, x1, y1, cd2, 0, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y1, x1, y1, cd2, 360, false).replace!("M", "L")} L${w},${y2}${PPTXShapeUtils.shapeArcAlt(w * (3 / 4), y2, x1, y1, 0, -cd2, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(x1, y2, x1, y1, 0, cd2, false).replace!("M", "L")} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartOnlineStorage": {
                        var x1, y1, c3d4 = 270, cd4: any = 90;
                        x1 = w * 1 / 6;
                        y1 = h * 3 / 6;
                        var d: any = `M${x1},${0} L${w},${0}${PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, 90, false).replace!("M", "L")} L${x1},${h}${PPTXShapeUtils.shapeArcAlt(x1, h / 2, x1, y1, cd4, 270, false).replace!("M", "L")} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartDisplay": {
                        var x1, x2: any, y1, c3d4 = 270, cd2: any = 180;
                        x1 = w * 1 / 6;
                        x2 = w * 5 / 6;
                        y1 = h * 3 / 6;
                        //path attrs: w = 6; h = 6; 
                        var d: any = `M${0},${y1} L${x1},${0} L${x2},${0}${PPTXShapeUtils.shapeArcAlt(w, h / 2, x1, y1, c3d4, c3d4 + cd2, false).replace!("M", "L")} L${x1},${h} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartDelay": {
                        var wd2: any = w / 2, hd2: any = h / 2, cd2: any = 180, c3d4 = 270, cd4: any = 90;
                        var d: any = `M${0},${0} L${wd2},${0}${PPTXShapeUtils.shapeArc(wd2, hd2, wd2, hd2, c3d4, c3d4 + cd2, false).replace("M", "L")} L${0},${h} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartMagneticTape": {
                        var wd2: any = w / 2, hd2: any = h / 2, cd2: any = 180, c3d4 = 270, cd4: any = 90;
                        let idy, ib, ang1;
                        idy = hd2 * Math.sin(Math.PI / 4);
                        ib = hd2 + idy;
                        ang1 = Math.atan(h / w);
                        const ang1Dg = ang1 * 180 / Math.PI;
                        var d: any = `M${wd2},${h}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd4, cd2, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, cd2, c3d4, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, c3d4, 360, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(wd2, hd2, wd2, hd2, 0, ang1Dg, false).replace!("M", "L")} L${w},${ib} L${w},${h} z`;
                        result += `<path d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "ellipse":
                    case "flowChartConnector":
                    case "flowChartSummingJunction":
                    case "flowChartOr": {
                        result += `<ellipse cx='${(w / 2)}' cy='${(h / 2)}' rx='${(w / 2)}' ry='${(h / 2)}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        if (shapType == "flowChartOr") {
                            result += ` <polyline points='${w / 2} ${0},${w / 2} ${h}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                            result += ` <polyline points='${0} ${h / 2},${w} ${h / 2}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        } else if (shapType == "flowChartSummingJunction") {
                            var iDx, idy, il, ir, it, ib, hc = w / 2, vc = h / 2, wd2: any = w / 2, hd2: any = h / 2;
                            const angVal = Math.PI / 4;
                            iDx = wd2 * Math.cos(angVal);
                            idy = hd2 * Math.sin(angVal);
                            il = hc - iDx;
                            ir = hc + iDx;
                            it = vc - idy;
                            ib = vc + idy;
                            result += ` <polyline points='${il} ${it},${ir} ${ib}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                            result += ` <polyline points='${ir} ${it},${il} ${ib}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
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
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, sAdj1_val;// = 0.33334;
                        let sAdj2, sAdj2_val;// = 0.33334;
                        let shpTyp, adjTyp;
                        if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor === Array) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        } else if (shapAdjst_ary !== undefined && shapAdjst_ary.constructor !== Array) {
                            const sAdj = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                            sAdj1_val = parseInt(sAdj.substr(4)) / 50000;
                            sAdj2_val = 0;
                        }
                        //console.log("shapType: ",shapType,",node: ",node )
                        let tranglRott = "";
                        switch (shapType) {
                            case "roundRect":
                            case "flowChartAlternateProcess": {
                                shpTyp = "round";
                                adjTyp = "cornrAll";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            }
                            case "round1Rect": {
                                shpTyp = "round";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                break;
                            }
                            case "round2DiagRect": {
                                shpTyp = "round";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                            case "round2SameRect": {
                                shpTyp = "round";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                            case "snip1Rect":
                            case "flowChartPunchedCard": {
                                shpTyp = "snip";
                                adjTyp = "cornr1";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                sAdj2_val = 0;
                                if (shapType == "flowChartPunchedCard") {
                                    tranglRott = `transform='translate(${w},0) scale(-1,1)'`;
                                }
                                break;
                            }
                            case "snip2DiagRect": {
                                shpTyp = "snip";
                                adjTyp = "diag";
                                if (sAdj1_val === undefined) sAdj1_val = 0;
                                if (sAdj2_val === undefined) sAdj2_val = 0.33334;
                                break;
                            }
                            case "snip2SameRect": {
                                shpTyp = "snip";
                                adjTyp = "cornr2";
                                if (sAdj1_val === undefined) sAdj1_val = 0.33334;
                                if (sAdj2_val === undefined) sAdj2_val = 0;
                                break;
                            }
                        }
                        let d_val = PPTXShapeUtils.shapeSnipRoundRectAlt(w, h, sAdj1_val!, sAdj2_val!, shpTyp!, adjTyp!);
                        result += `<path ${tranglRott}  d='${d_val}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "snipRoundRect": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, sAdj1_val = 0.33334;
                        let sAdj2, sAdj2_val = 0.33334;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) / 50000;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) / 50000;
                                }
                            }
                        }
                        /**
                         * snipRoundRect: 混合形状，有两个角是圆角，有两个角是缺角
                         *
                         * 形状说明：
                         * - 左上角：凹进去的缺角（直线斜切）
                         * - 右上角：凹进去的缺角（直线斜切）
                         * - 右下角：凹进去的圆角
                         * - 左下角：凹进去的圆角
                         *
                         * 参数说明：
                         * - adj1: 控制圆角的半径（用于右下角和左下角）
                         * - adj2: 控制缺角的大小（用于左上角和右上角）
                         */
                        const radius = Math.min(w, h) * sAdj1_val;     // 圆角半径
                        const snipSize = Math.min(w, h) * sAdj2_val;   // 缺角大小

                        // 生成路径：从左下角开始，逆时针绘制
                        let d_val = `M0,${(h - radius)} Q0,${h} ${radius},${h} L${w},${h} Q${w},${h} ${w},${(h - radius)} L${w},${snipSize} L${(w - snipSize)},0 L${snipSize},0 L0,${(h - snipSize)} z`;

                        result += `<path   d='${d_val}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "bentConnector2": {
                        var d: any = "";
                        // 使用drawW和drawH（原始尺寸）
                        const bendW = (drawW !== undefined) ? drawW : w;
                        const bendH = (drawH !== undefined) ? drawH : h;
                        // 路径方向（SVG容器会通过flip变换处理翻转）
                        d = `M ${bendW} 0 L ${bendW} ${bendH} L 0 ${bendH}`;
                        result += `<path d='${d}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' fill='none' `;
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-start='url(#markerTriangle_${shpId})' `;
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-end='url(#markerTriangle_${shpId})' `;
                        }
                        result += "/>";
                        break;
                    }
                    case "rtTriangle": {
                        result += ` <polygon points='0 0,0 ${h},${w} ${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "triangle":
                    case "flowChartExtract":
                    case "flowChartMerge": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let shapAdjst_val = 0.5;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                            //console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nshapAdjst_val: "+shapAdjst_val);
                        }
                        let tranglRott = "";
                        if (shapType == "flowChartMerge") {
                            tranglRott = `transform='rotate(180 ${w / 2},${h / 2})'`;
                        }
                        result += ` <polygon ${tranglRott} points='${(w * shapAdjst_val)} 0,0 ${h},${w} ${h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "diamond":
                    case "flowChartDecision":
                    case "flowChartSort": {
                        result += ` <polygon points='${(w / 2)} 0,0 ${(h / 2)},${(w / 2)} ${h},${w} ${(h / 2)}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        if (shapType == "flowChartSort") {
                            result += ` <polyline points='0 ${h / 2},${w} ${h / 2}' fill='none' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        }
                        break;
                    }
                    case "trapezoid":
                    case "flowChartManualOperation":
                    case "flowChartManualInput": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adjst_val = 0.2;
                        let max_adj_const = 0.7407;
                        if (shapAdjst !== undefined) {
                            const adjst = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                            adjst_val = (adjst * 0.5) / max_adj_const;
                            // console.log("w: "+w+"\nh: "+h+"\nshapAdjst: "+shapAdjst+"\nadjst_val: "+adjst_val);
                        }
                        var cnstVal: any = 0;
                        let tranglRott = "";
                        if (shapType == "flowChartManualOperation") {
                            tranglRott = `transform='rotate(180 ${w / 2},${h / 2})'`;
                        }
                        if (shapType == "flowChartManualInput") {
                            adjst_val = 0;
                            cnstVal = h / 5;
                        }
                        result += ` <polygon ${tranglRott} points='${(w * adjst_val)} ${cnstVal},0 ${h},${w} ${h},${(1 - adjst_val) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "parallelogram":
                    case "flowChartInputOutput": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adjst_val = 0.25;
                        let max_adj_const;
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
                        result += ` <polygon points='${adjst_val * w} 0,0 ${h},${(1 - adjst_val) * w} ${h},${w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "pentagon": {
                        result += ` <polygon points='${(0.5 * w)} 0,0 ${(0.375 * h)},${(0.15 * w)} ${h},${0.85 * w} ${h},${w} ${0.375 * h}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "hexagon":
                    case "flowChartPreparation": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 25000 * SLIDE_FACTOR;
                        const vf = 115470 * SLIDE_FACTOR;;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const angVal1 = 60 * Math.PI / 180;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var maxAdj, a, shd2, x1, x2: any, dy1, y1, y2: any, vc = h / 2, hd2: any = h / 2;
                        const ss = Math.min(w, h);
                        maxAdj = cnstVal1 * w / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        shd2 = hd2 * vf / cnstVal2;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        dy1 = shd2 * Math.sin(angVal1);
                        y1 = vc - dy1;
                        y2 = vc + dy1;

                        var d: any = `M${0},${vc} L${x1},${y1} L${x2},${y1} L${w},${vc} L${x2},${y2} L${x1},${y2} z`;

                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "heptagon": {
                        result += ` <polygon points='${(0.5 * w)} 0,${w / 8} ${h / 4},0 ${(5 / 8) * h},${w / 4} ${h},${(3 / 4) * w} ${h},${w} ${(5 / 8) * h},${(7 / 8) * w} ${h / 4}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "octagon": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        let adj2 = (1 - adj1);
                        //console.log("adj1: "+adj1+"\nadj2: "+adj2);
                        result += ` <polygon points='${adj1 * w} 0,0 ${adj1 * h},0 ${adj2 * h},${adj1 * w} ${h},${adj2 * w} ${h},${w} ${adj2 * h},${w} ${adj1 * h},${adj2 * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "decagon": {
                        result += ` <polygon points='${(3 / 8) * w} 0,${w / 8} ${h / 8},0 ${h / 2},${w / 8} ${(7 / 8) * h},${(3 / 8) * w} ${h},${(5 / 8) * w} ${h},${(7 / 8) * w} ${(7 / 8) * h},${w} ${h / 2},${(7 / 8) * w} ${h / 8},${(5 / 8) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "dodecagon": {
                        result += ` <polygon points='${(3 / 8) * w} 0,${w / 8} ${h / 8},0 ${(3 / 8) * h},0 ${(5 / 8) * h},${w / 8} ${(7 / 8) * h},${(3 / 8) * w} ${h},${(5 / 8) * w} ${h},${(7 / 8) * w} ${(7 / 8) * h},${w} ${(5 / 8) * h},${w} ${(3 / 8) * h},${(7 / 8) * w} ${h / 8},${(5 / 8) * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "star4":
                    case "star5":
                    case "star6":
                    case "star7":
                    case "star8":
                    case "star10":
                    case "star12":
                    case "star16":
                    case "star24":
                    case "star32": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderStar(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArcAlt, node!);
                        break;
                    }
                    case "pie":
                    case "pieWedge":
                    case "arc":
                    case "chord": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderPieShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node!, oShadowSvgUrlStr);
                        break;
                    }
                    case "frame": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj1 = 12500 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var a1, x1, x4, y4;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        x1 = Math.min(w, h) * a1 / cnstVal2;
                        x4 = w - x1;
                        y4 = h - x1;
                        var d: any = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} zM${x1},${x1} L${x1},${y4} L${x4},${y4} L${x4},${x1} z`;
                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "donut": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        let a, dr, iwd2, ihd2;
                        if (adj < 0) a = 0
                        else if (adj > cnstVal1) a = cnstVal1
                        else a = adj
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        var d: any = `M${0},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L")} zM${dr},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 180, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 90, 0, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 0, -90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, iwd2, ihd2, 270, 180, false).replace("M", "L")} z`;
                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' ${oShadowSvgUrlStr} />`;
                        break;
                    }
                    case "noSmoking": {
                        /**
                         * noSmoking: 禁止符形状
                         * 参考 pptxjs.js 实现
                         *
                         * 形状说明：
                         * - 一个完整的圆圈
                         * - 中间有一条从左上到右下的斜杠（带圆角）
                         *
                         * 参数说明：
                         * - adj: 控制斜杠的粗细 (范围: 0-50000)
                         */
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 18750 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }

                        // 计算调整值
                        let a, dr, iwd2, ihd2, ang, ct, st, m, n;
                        if (adj < 0) a = 0;
                        else if (adj > cnstVal1) a = cnstVal1;
                        else a = adj;

                        // 斜杠宽度
                        dr = Math.min(w, h) * a / cnstVal2;
                        iwd2 = w / 2 - dr;
                        ihd2 = h / 2 - dr;
                        ang = Math.atan(h / w);
                        ct = ihd2 * Math.cos(ang);
                        st = iwd2 * Math.sin(ang);
                        m = Math.sqrt(ct * ct + st * st);
                        n = iwd2 * ihd2 / m;
                        const drd2 = dr / 2;
                        const dang = Math.atan(drd2 / n);
                        let dang2 = dang * 2;
                        let swAng = -Math.PI + dang2;

                        // 绘制路径（参考 pptxjs.js 使用圆弧方式）
                        const stAng1 = ang - dang;
                        let stAng2 = stAng1 - Math.PI;
                        const stAng1deg = stAng1 * 180 / Math.PI;
                        const stAng2deg = stAng2 * 180 / Math.PI;
                        const swAng2deg = swAng * 180 / Math.PI;

                        let dx1 = n * Math.cos(stAng1);
                        let dy1 = n * Math.sin(stAng1);
                        var x1: any = w / 2 + dx1;
                        var y1: any = h / 2 + dy1;
                        var x2: any = w / 2 - dx1;
                        var y2: any = h / 2 - dy1;

                        var d: any = `M${0},${h / 2}${shapeArcAlt!(w / 2, h / 2, w / 2, h / 2, 180, 270, false).replace!("M", "L")}${shapeArcAlt!(w / 2, h / 2, w / 2, h / 2, 270, 360, false).replace!("M", "L")}${shapeArcAlt!(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace!("M", "L")}${shapeArcAlt!(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace!("M", "L")} zM${x1},${y1}${shapeArcAlt!(w / 2, h / 2, iwd2, ihd2, stAng1deg, (stAng1deg + swAng2deg), false).replace!("M", "L")} zM${x2},${y2}${shapeArcAlt!(w / 2, h / 2, iwd2, ihd2, stAng2deg, (stAng2deg + swAng2deg), false).replace!("M", "L")} z`;

                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "halfFrame": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, sAdj1_val = 3.5;
                        let sAdj2, sAdj2_val = 3.5;
                        const cnsVal = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        const minWH = Math.min(w, h);
                        let maxAdj2 = (cnsVal * w) / minWH;
                        let a1, a2;
                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        var x1: any = (minWH * a2) / cnsVal;
                        const g1 = h * x1 / w;
                        let g2 = h - g1;
                        let maxAdj1 = (cnsVal * g2) / minWH;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val
                        var y1: any = minWH * a1 / cnsVal;
                        var dx2: any = y1 * w / h;
                        var x2: any = w - dx2;
                        let dy2 = x1 * h / w;
                        var y2: any = h - dy2;
                        var d: any = `M0,0 L${w},${0} L${x2},${y1} L${x1},${y1} L${x1},${y2} L0,${h} z`;

                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        //console.log("w: ",w,", h: ",h,", sAdj1_val: ",sAdj1_val,", sAdj2_val: ",sAdj2_val,",maxAdj1: ",maxAdj1,",maxAdj2: ",maxAdj2)
                        break;
                    }
                    case "bracePair":
                    case "bracketPair":
                    case "leftBrace":
                    case "leftBracket":
                    case "rightBrace":
                    case "rightBracket": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderBracket(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node!);
                        break;
                    }
                    case "moon": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 0.5;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) / 100000;//*96/914400;;
                        }
                        var hd2: any, cd2: any, cd4: any;

                        hd2 = h / 2;
                        cd2 = 180;
                        cd4 = 90;

                        let adj2 = (1 - adj) * w;
                        var d: any = `M${w},${h}${PPTXShapeUtils.shapeArc(w, hd2, w, hd2, cd4, (cd4 + cd2), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w, hd2, adj2, hd2, (cd4 + cd2), cd4, false).replace("M", "L")} z`;
                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "corner": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, sAdj1_val = 50000 * SLIDE_FACTOR;
                        let sAdj2, sAdj2_val = 50000 * SLIDE_FACTOR;
                        const cnsVal = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj1_val = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    sAdj2_val = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        const minWH = Math.min(w, h);
                        let maxAdj1 = cnsVal * h / minWH;
                        let maxAdj2 = cnsVal * w / minWH;
                        var a1, a2, x1, dy1, y1;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > maxAdj1) a1 = maxAdj1
                        else a1 = sAdj1_val

                        if (sAdj2_val < 0) a2 = 0
                        else if (sAdj2_val > maxAdj2) a2 = maxAdj2
                        else a2 = sAdj2_val
                        x1 = minWH * a2 / cnsVal;
                        dy1 = minWH * a1 / cnsVal;
                        y1 = h - dy1;

                        var d: any = `M0,0 L${x1},${0} L${x1},${y1} L${w},${y1} L${w},${h} L0,${h} z`;

                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "diagStripe": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let sAdj1_val = 50000 * SLIDE_FACTOR;
                        const cnsVal = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            sAdj1_val = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var a1, x2: any, y2: any;
                        if (sAdj1_val < 0) a1 = 0
                        else if (sAdj1_val > cnsVal) a1 = cnsVal
                        else a1 = sAdj1_val
                        x2 = w * a1 / cnsVal;
                        y2 = h * a1 / cnsVal;
                        var d: any = `M${0},${y2} L${x2},${0} L${w},${0} L${0},${h} z`;

                        result += `<path   d='${d}'  fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "gear6":
                    case "gear9": {
                        txtRotate = 0;
                        var gearNum = shapType.substr(4), d: any;
                        if (gearNum == "6") {
                            d = shapeGear(w, h / 3.5, parseInt(gearNum));
                        } else { //gearNum=="9"
                            d = shapeGear(w, h / 3.5, parseInt(gearNum));
                        }
                        result += `<path   d='${d}' transform='rotate(20,${(3 / 7) * h},${(3 / 7) * h})' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "bentConnector3": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let shapAdjst_val = 0.5;
                        // 使用drawW和drawH（原始尺寸）
                        const connectorW = (drawW !== undefined) ? drawW : w;
                        const connectorH = (drawH !== undefined) ? drawH : h;
                        if (shapAdjst !== undefined) {
                            shapAdjst_val = parseInt(shapAdjst.substr(4)) / 100000;
                            // 路径方向（SVG容器会通过flip变换处理翻转）
                            result += ` <polyline points='0 0,${(shapAdjst_val) * connectorW} 0,${(shapAdjst_val) * connectorW} ${connectorH},${connectorW} ${connectorH}' fill='transparent'' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' `;
                            if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                                result += `marker-start='url(#markerTriangle_${shpId})' `;
                            }
                            if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                                result += `marker-end='url(#markerTriangle_${shpId})' `;
                            }
                            result += "/>";
                        }
                        break;
                    }
                    case "plus": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj1 = 0.25;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) / 100000;

                        }
                        let adj2 = (1 - adj1);
                        result += ` <polygon points='${adj1 * w} 0,${adj1 * w} ${adj1 * h},0 ${adj1 * h},0 ${adj2 * h},${adj1 * w} ${adj2 * h},${adj1 * w} ${h},${adj2 * w} ${h},${adj2 * w} ${adj2 * h},${w} ${adj2 * h},${+w} ${adj1 * h},${adj2 * w} ${adj1 * h},${adj2 * w} 0' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "teardrop": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj1 = 100000 * SLIDE_FACTOR;
                        const cnsVal1 = adj1;
                        const cnsVal2 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var a1, r2, tw, th, sw, sh, dx1, dy1, x1, y1, x2: any, y2: any, rd45;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnsVal2) a1 = cnsVal2
                        else a1 = adj1
                        r2 = Math.sqrt(2);
                        tw = r2 * (w / 2);
                        th = r2 * (h / 2);
                        sw = (tw * a1) / cnsVal1;
                        sh = (th * a1) / cnsVal1;
                        rd45 = (45 * (Math.PI) / 180);
                        dx1 = sw * (Math.cos(rd45));
                        dy1 = sh * (Math.cos(rd45));
                        x1 = (w / 2) + dx1;
                        y1 = (h / 2) - dy1;
                        x2 = ((w / 2) + x1) / 2;
                        y2 = ((h / 2) + y1) / 2;

                        let d_val = `${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 180, 270, false)}Q ${x2},0 ${x1},${y1}Q ${w},${y2} ${w},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 0, 90, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(w / 2, h / 2, w / 2, h / 2, 90, 180, false).replace("M", "L")} z`;
                        result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        // console.log("shapAdjst: ",shapAdjst,", adj1: ",adj1);
                        break;
                    }
                    case "plaque": {
                        /**
                         * plaque: 凸出圆角的矩形
                         *
                         * 形状说明：
                         * - 4个角都是向外凸出的1/4圆弧
                         * - 类似凸出卡片的样式
                         * - 半圆弧的圆心在矩形的四个角上
                         *
                         * 参数说明：
                         * - adj: 控制圆角半径大小 (范围: 0-50000)
                         *
                         * 坐标系统：
                         * - (0, 0) 到 (w, h) 的矩形区域
                         * - 四个角向外延伸出1/4圆弧
                         */

                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adjVal = 25000; // 默认值
                        if (shapAdjst !== undefined) {
                            adjVal = parseInt(shapAdjst.substr(4));
                        }
                        // 限制 adj 在有效范围内 (0-50000)
                        if (adjVal < 0) adjVal = 0;
                        else if (adjVal > 50000) adjVal = 50000;

                        // 计算圆弧半径：adj/100000 * min(w, h)
                        let r: any = (adjVal / 100000) * Math.min(w, h);

                        /**
                         * 路径绘制顺序（逆时针从左上角圆弧开始）：
                         *
                         * 左上角：向外凸出的1/4圆弧，圆心在(0,0)
                         * - 起点: (r, 0)
                         * - 圆弧到: (0, r) - 1/4圆弧（0度到90度）
                         *
                         * 右上角：向外凸出的1/4圆弧，圆心在(w,0)
                         * - 线段到: (w - r, 0)
                         * - 圆弧到: (w, r) - 1/4圆弧（90度到180度）
                         *
                         * 右下角：向外凸出的1/4圆弧，圆心在(w,h)
                         * - 线段到: (w, h - r)
                         * - 圆弧到: (w - r, h) - 1/4圆弧（180度到270度）
                         *
                         * 左下角：向外凸出的1/4圆弧，圆心在(0,h)
                         * - 线段到: (r, h)
                         * - 圆弧到: (0, h - r) - 1/4圆弧（270度到360度）
                         * - 闭合: 回到起点
                         */

                        let d_val = `M${r},0A${r} ${r} 0 0 1 0,${r}L0,${(h - r)}A${r} ${r} 0 0 1 ${r},${h}L${(w - r)},${h}A${r} ${r} 0 0 1 ${w},${(h - r)}L${w},${r}A${r} ${r} 0 0 1 ${(w - r)},0 z`;

                        result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "sun": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr = SLIDE_FACTOR;
                        let adj1 = 25000 * refr;
                        const cnstVal1 = 12500 * refr;
                        const cnstVal2 = 46875 * refr;
                        if (shapAdjst !== undefined) {
                            adj1 = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        let a1;
                        if (adj1 < cnstVal1) a1 = cnstVal1
                        else if (adj1 > cnstVal2) a1 = cnstVal2
                        else a1 = adj1

                        const cnstVa3 = 50000 * refr;
                        const cnstVa4 = 100000 * refr;
                        let g0 = cnstVa3 - a1,
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
                            x8: any = w * g8 / cnstVa4,
                            x9: any = w * g9 / cnstVa4,
                            x10: any = w * g10 / cnstVa4,
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
                            y15 = h * g15 / cnstVa4,
                            y16 = h * g16 / cnstVa4,
                            y17 = h * g17 / cnstVa4,
                            y18 = h * g18 / cnstVa4;

                        let d_val = `M${w},${h / 2} L${x15},${y18} L${x15},${y14}z M${ox1},${oy1} L${x16},${y17} L${x13},${y12}z M${w / 2},${0} L${x18},${y10} L${x14},${y10}z M${ox2},${oy1} L${x17},${y12} L${x12},${y17}z M${0},${h / 2} L${x10},${y14} L${x10},${y18}z M${ox2},${oy2} L${x12},${y13} L${x17},${y16}z M${w / 2},${h} L${x14},${y15} L${x18},${y15}z M${ox1},${oy2} L${x13},${y16} L${x16},${y13} z M${x19},${h / 2}${PPTXShapeUtils.shapeArc(w / 2, h / 2, wR, hR, 180, 540, false).replace("M", "L")} z`;
                        //console.log("adj1: ",adj1,d_val);
                        result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;


                        break;
                    }
                    case "heart": {
                        var dx1, dx2: any, x1, x2: any, x3, x4, y1;
                        dx1 = w * 49 / 48;
                        dx2 = w * 10 / 48
                        x1 = w / 2 - dx1
                        x2 = w / 2 - dx2
                        x3 = w / 2 + dx2
                        x4 = w / 2 + dx1
                        y1 = -h / 3;
                        let d_val = `M${w / 2},${h / 4}C${x3},${y1} ${x4},${h / 4} ${w / 2},${h}C${x1},${h / 4} ${x2},${y1} ${w / 2},${h / 4} z`;

                        result += `<path   d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "lightningBolt": {
                        var x1: any = w * 5022 / 21600,
                            x2: any = w * 11050 / 21600,
                            x3: any = w * 8472 / 21600,
                            x4: any = w * 8757 / 21600,
                            x5: any = w * 10012 / 21600,
                            x6: any = w * 14767 / 21600,
                            x7: any = w * 12222 / 21600,
                            x8: any = w * 12860 / 21600,
                            x9: any = w * 13917 / 21600,
                            x10: any = w * 7602 / 21600,
                            x11: any = w * 16577 / 21600,
                            y1: any = h * 3890 / 21600,
                            y2: any = h * 6080 / 21600,
                            y3: any = h * 6797 / 21600,
                            y4: any = h * 7437 / 21600,
                            y5: any = h * 12877 / 21600,
                            y6: any = h * 9705 / 21600,
                            y7: any = h * 12007 / 21600,
                            y8: any = h * 13987 / 21600,
                            y9: any = h * 8382 / 21600,
                            y10 = h * 14277 / 21600,
                            y11 = h * 14915 / 21600;

                        let d_val = `M${x3},${0} L${x8},${y2} L${x2},${y3} L${x11},${y7} L${x6},${y5} L${w},${h} L${x5},${y11} L${x7},${y8} L${x1},${y6} L${x10},${y9} L${0},${y1} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "cube": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr = SLIDE_FACTOR;
                        let adj = 25000 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        let d_val;
                        const cnstVal2 = 100000 * refr;
                        const ss = Math.min(w, h);
                        var a, y1, y4, x4;
                        a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        y1 = ss * a / cnstVal2;
                        y4 = h - y1;
                        x4 = w - y1;
                        d_val = `M${0},${y1} L${y1},${0} L${w},${0} L${w},${y4} L${x4},${h} L${0},${h} zM${0},${y1} L${x4},${y1} M${x4},${y1} L${w},${0}M${x4},${y1} L${x4},${h}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "bevel": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr = SLIDE_FACTOR;
                        let adj = 12500 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        let d_val;
                        const cnstVal1 = 50000 * refr;
                        const cnstVal2 = 100000 * refr;
                        const ss = Math.min(w, h);
                        var a, x1, x2: any, y2: any;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        x1 = ss * a / cnstVal2;
                        x2 = w - x1;
                        y2 = h - x1;
                        d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${x1} L${x2},${x1} L${x2},${y2} L${x1},${y2} z M${0},${0} L${x1},${x1} M${0},${h} L${x1},${y2} M${w},${0} L${x2},${x1} M${w},${h} L${x2},${y2}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "foldedCorner": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        const refr = SLIDE_FACTOR;
                        let adj = 16667 * refr;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * refr;
                        }
                        let d_val;
                        const cnstVal1 = 50000 * refr;
                        const cnstVal2 = 100000 * refr;
                        const ss = Math.min(w, h);
                        var a, dy2, dy1, x1, x2: any, y2: any, y1;
                        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
                        dy2 = ss * a / cnstVal2;
                        dy1 = dy2 / 5;
                        x1 = w - dy2;
                        x2 = x1 + dy1;
                        y2 = h - dy2;
                        y1 = y2 + dy1;
                        d_val = `M${x1},${h} L${x2},${y1} L${w},${y2} L${x1},${h} L${0},${h} L${0},${0} L${w},${0} L${w},${y2}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "cloud":
                    case "cloudCallout": {
                        // 云形的原始设计是基于 43200x43200 的正方形
                        // 根据 Office Open XML 规范，X坐标使用w缩放，Y坐标使用h缩放

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num: any) {
                            return parseFloat(num.toFixed(2));
                        }

                        // 生成椭圆弧路径的辅助函数（使用SVG A命令）
                        // 参数：中心点(cx,cy)，半径(rx,ry)，起始角度startAngle，扫描角度sweepAngle
                        function ellipseArc(cx: any, cy: any, rx: any, ry: any, startAngle: any, sweepAngle: any) {
                            const endAngle = startAngle + sweepAngle;
                            // 计算起点和终点
                            const startX = cx + rx * Math.cos(startAngle * Math.PI / 180);
                            const startY = cy + ry * Math.sin(startAngle * Math.PI / 180);
                            const endX = cx + rx * Math.cos(endAngle * Math.PI / 180);
                            const endY = cy + ry * Math.sin(endAngle * Math.PI / 180);
                            
                            // 确定large-arc-flag和sweep-flag
                            const largeArc = Math.abs(sweepAngle) > 180 ? 1 : 0;
                            const sweep = sweepAngle > 0 ? 1 : 0;
                            
                            return {
                                start: { x: fmt(startX), y: fmt(startY) },
                                end: { x: fmt(endX), y: fmt(endY) },
                                path: `A ${fmt(rx)} ${fmt(ry)} 0 ${largeArc} ${sweep} ${fmt(endX)} ${fmt(endY)}`
                            };
                        }

                        // X坐标使用 w 缩放，Y坐标使用 h 缩放
                        const x0 = fmt(w * 3900 / 43200);
                        const y0 = fmt(h * 14370 / 43200);
                        
                        // 半径：RX使用 w 缩放，RY使用 h 缩放
                        const rX1 = fmt(w * 6753 / 43200), rY1 = fmt(h * 9190 / 43200);
                        const rX2 = fmt(w * 5333 / 43200), rY2 = fmt(h * 7267 / 43200);
                        const rX3 = fmt(w * 4365 / 43200), rY3 = fmt(h * 5945 / 43200);
                        const rX4 = fmt(w * 4857 / 43200), rY4 = fmt(h * 6595 / 43200);
                        const rY5 = fmt(h * 7273 / 43200);
                        const rX6 = fmt(w * 6775 / 43200), rY6 = fmt(h * 9220 / 43200);
                        const rX7 = fmt(w * 5785 / 43200), rY7 = fmt(h * 7867 / 43200);
                        const rX8 = fmt(w * 6752 / 43200), rY8 = fmt(h * 9215 / 43200);
                        const rX9 = fmt(w * 7720 / 43200), rY9 = fmt(h * 10543 / 43200);
                        const rX10 = fmt(w * 4360 / 43200), rY10 = fmt(h * 5918 / 43200);
                        const rX11 = fmt(w * 4345 / 43200);

                        // 角度（以度为单位）
                        const sA1 = -11429249 / 60000, wA1 = 7426832 / 60000;
                        const sA2 = -8646143 / 60000, wA2 = 5396714 / 60000;
                        const sA3 = -8748475 / 60000, wA3 = 5983381 / 60000;
                        const sA4 = -7859164 / 60000, wA4 = 7034504 / 60000;
                        const sA5 = -4722533 / 60000, wA5 = 6541615 / 60000;
                        const sA6 = -2776035 / 60000, wA6 = 7816140 / 60000;
                        const sA7 = 37501 / 60000, wA7 = 6842000 / 60000;
                        const sA8 = 1347096 / 60000, wA8 = 6910353 / 60000;
                        const sA9 = 3974558 / 60000, wA9 = 4542661 / 60000;
                        const sA10 = -16496525 / 60000, wA10 = 8804134 / 60000;
                        const sA11 = -14809710 / 60000, wA11 = 9151131 / 60000;

                        // 计算各弧线的中心点
                        // 弧线中心点 = 起点 - 半径 * cos/sin(起始角度)
                        const cX0 = fmt(x0 - rX1 * Math.cos(sA1 * Math.PI / 180));
                        const cY0 = fmt(y0 - rY1 * Math.sin(sA1 * Math.PI / 180));

                        // 生成弧线1
                        const arc1 = ellipseArc(cX0, cY0, rX1, rY1, sA1, wA1);
                        
                        // 计算弧线2的中心点（基于弧线1的终点）
                        const cX1 = fmt(arc1.end.x - rX2 * Math.cos(sA2 * Math.PI / 180));
                        const cY1 = fmt(arc1.end.y - rY2 * Math.sin(sA2 * Math.PI / 180));
                        const arc2 = ellipseArc(cX1, cY1, rX2, rY2, sA2, wA2);
                        
                        // 弧线3
                        const cX2 = fmt(arc2.end.x - rX3 * Math.cos(sA3 * Math.PI / 180));
                        const cY2 = fmt(arc2.end.y - rY3 * Math.sin(sA3 * Math.PI / 180));
                        const arc3 = ellipseArc(cX2, cY2, rX3, rY3, sA3, wA3);
                        
                        // 弧线4
                        const cX3 = fmt(arc3.end.x - rX4 * Math.cos(sA4 * Math.PI / 180));
                        const cY3 = fmt(arc3.end.y - rY4 * Math.sin(sA4 * Math.PI / 180));
                        const arc4 = ellipseArc(cX3, cY3, rX4, rY4, sA4, wA4);
                        
                        // 弧线5
                        const cX4 = fmt(arc4.end.x - rX2 * Math.cos(sA5 * Math.PI / 180));
                        const cY4 = fmt(arc4.end.y - rY5 * Math.sin(sA5 * Math.PI / 180));
                        const arc5 = ellipseArc(cX4, cY4, rX2, rY5, sA5, wA5);
                        
                        // 弧线6
                        const cX5 = fmt(arc5.end.x - rX6 * Math.cos(sA6 * Math.PI / 180));
                        const cY5 = fmt(arc5.end.y - rY6 * Math.sin(sA6 * Math.PI / 180));
                        const arc6 = ellipseArc(cX5, cY5, rX6, rY6, sA6, wA6);
                        
                        // 弧线7
                        const cX6 = fmt(arc6.end.x - rX7 * Math.cos(sA7 * Math.PI / 180));
                        const cY6 = fmt(arc6.end.y - rY7 * Math.sin(sA7 * Math.PI / 180));
                        const arc7 = ellipseArc(cX6, cY6, rX7, rY7, sA7, wA7);
                        
                        // 弧线8
                        const cX7 = fmt(arc7.end.x - rX8 * Math.cos(sA8 * Math.PI / 180));
                        const cY7 = fmt(arc7.end.y - rY8 * Math.sin(sA8 * Math.PI / 180));
                        const arc8 = ellipseArc(cX7, cY7, rX8, rY8, sA8, wA8);
                        
                        // 弧线9
                        const cX8 = fmt(arc8.end.x - rX9 * Math.cos(sA9 * Math.PI / 180));
                        const cY8 = fmt(arc8.end.y - rY9 * Math.sin(sA9 * Math.PI / 180));
                        const arc9 = ellipseArc(cX8, cY8, rX9, rY9, sA9, wA9);
                        
                        // 弧线10
                        const cX9 = fmt(arc9.end.x - rX10 * Math.cos(sA10 * Math.PI / 180));
                        const cY9 = fmt(arc9.end.y - rY10 * Math.sin(sA10 * Math.PI / 180));
                        const arc10 = ellipseArc(cX9, cY9, rX10, rY10, sA10, wA10);
                        
                        // 弧线11
                        const cX10 = fmt(arc10.end.x - rX11 * Math.cos(sA11 * Math.PI / 180));
                        const cY10 = fmt(arc10.end.y - rY3 * Math.sin(sA11 * Math.PI / 180));
                        const arc11 = ellipseArc(cX10, cY10, rX11, rY3, sA11, wA11);

                        // 构建完整路径
                        let d1 = `M${x0},${y0} ${arc1.path} ${arc2.path} ${arc3.path} ${arc4.path} ${arc5.path} ${arc6.path} ${arc7.path} ${arc8.path} ${arc9.path} ${arc10.path} ${arc11.path} z`;

                        if (shapType == "cloudCallout") {
                            const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                            const refr = SLIDE_FACTOR;
                            let sAdj1, adj1 = -20833 * refr;
                            let sAdj2, adj2 = 62500 * refr;
                            if (shapAdjst_ary !== undefined) {
                                for (const i of shapAdjst_ary.keys()){
                                    const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4)) * refr;
                                    } else if (sAdj_name == "adj2") {
                                        sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj2 = parseInt(sAdj2.substr(4)) * refr;
                                    }
                                }
                            }
                            let d_val;
                            const cnstVal2 = 100000 * refr;
                            const ss = Math.min(w, h);
                            var wd2: any = w / 2, hd2: any = h / 2;

                            let dxPos, dyPos, xPos, yPos, ht, wt, g2, g3, g4, g5, g6, g7, g8, g9, g10, g11, g12, g13, g14, g15, g16,
                                g17, g18, g19, g20, g21, g22, g23, g24, g25, g26, x23, x24, x25;

                            dxPos = w * adj1 / cnstVal2;
                            dyPos = h * adj2 / cnstVal2;
                            xPos = wd2 + dxPos;
                            yPos = hd2 + dyPos;
                            ht = hd2 * Math.cos(Math.atan(dyPos / dxPos));
                            wt = wd2 * Math.sin(Math.atan(dyPos / dxPos));
                            g2 = wd2 * Math.cos(Math.atan(wt / ht));
                            g3 = hd2 * Math.sin(Math.atan(wt / ht));
                            //console.log("adj1: ",adj1,"adj2: ",adj2)
                            if (adj1 >= 0) {
                                g4 = wd2 + g2;
                                g5 = hd2 + g3;
                            } else {
                                g4 = wd2 - g2;
                                g5 = hd2 - g3;
                            }
                            g6 = g4 - xPos;
                            g7 = g5 - yPos;
                            g8 = Math.sqrt(g6 * g6 + g7 * g7);
                            g9 = ss * 6600 / 21600;
                            g10 = g8 - g9;
                            g11 = g10 / 3;
                            g12 = ss * 1800 / 21600;
                            g13 = g11 + g12;
                            g14 = g13 * g6 / g8;
                            g15 = g13 * g7 / g8;
                            g16 = g14 + xPos;
                            g17 = g15 + yPos;
                            g18 = ss * 4800 / 21600;
                            g19 = g11 * 2;
                            g20 = g18 + g19;
                            g21 = g20 * g6 / g8;
                            g22 = g20 * g7 / g8;
                            g23 = g21 + xPos;
                            g24 = g22 + yPos;
                            g25 = ss * 1200 / 21600;
                            g26 = ss * 600 / 21600;
                            x23 = xPos + g26;
                            x24 = g16 + g25;
                            x25 = g23 + g12;

                            d_val = //" M" + x23 + "," + yPos + 
                                `${PPTXShapeUtils.shapeArc(x23 - g26, yPos, g26, g26, 0, 360, false)} z M${x24},${g17}${PPTXShapeUtils.shapeArc(x24 - g25, g17, g25, g25, 0, 360, false).replace("M", "L")} z M${x25},${g24}${PPTXShapeUtils.shapeArc(x25 - g12, g24, g12, g12, 0, 360, false).replace("M", "L")} z`;
                            d1 += d_val;
                        }
                        result += `<path d='${d1}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "smileyFace":
                    case "verticalScroll":
                    case "horizontalScroll": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderMiscShape(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node!);
                        break;
                    }
                    case "wedgeEllipseCallout": {
                        // cloudTransformAttr: 原代码引用了一个从未声明的变换属性，
                        // 运行时会导致 ReferenceError。其余 preset 形状路径均不带 transform 属性，
                        // 故此处置为空字符串，保证该分支行为正确且不崩溃。
                        const cloudTransformAttr = '';
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = -20833 * refr;
                        let sAdj2, adj2 = 62500 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        const angVal1 = 11 * Math.PI / 180;
                        const ss = Math.min(w, h);
                        var dxPos, dyPos, xPos, yPos, sdx, sdy, pang, stAng, enAng, dx1, dy1, x1, y1, dx2: any, dy2,
                            x2: any, y2: any, stAng1, enAng1, swAng1, swAng2, swAng,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        sdx = dxPos * h;
                        sdy = dyPos * w;
                        pang = Math.atan(sdy / sdx);
                        stAng = pang + angVal1;
                        enAng = pang - angVal1;
                        dx1 = hc * Math.cos(stAng);
                        dy1 = vc * Math.sin(stAng);
                        dx2 = hc * Math.cos(enAng);
                        dy2 = vc * Math.sin(enAng);
                        if (dxPos >= 0) {
                            x1 = hc + dx1;
                            y1 = vc + dy1;
                            x2 = hc + dx2;
                            y2 = vc + dy2;
                        } else {
                            x1 = hc - dx1;
                            y1 = vc - dy1;
                            x2 = hc - dx2;
                            y2 = vc - dy2;
                        }
                        /*
                        //stAng = pang+angVal1;
                        //enAng = pang-angVal1;
                        //dx1 = hc*Math.cos(stAng);
                        //dy1 = vc*Math.sin(stAng);
                        x1 = hc+dx1;
                        y1 = vc+dy1;
                        dx2 = hc*Math.cos(enAng);
                        dy2 = vc*Math.sin(enAng);
                        x2 = hc+dx2;
                        y2 = vc+dy2;
                        stAng1 = Math.atan(dy1/dx1);
                        enAng1 = Math.atan(dy2/dx2);
                        swAng1 = enAng1-stAng1;
                        swAng2 = swAng1+2*Math.PI;
                        swAng = (swAng1 > 0)?swAng1:swAng2;
                        var stAng1Dg = stAng1*180/Math.PI;
                        var swAngDg = swAng*180/Math.PI;
                        var endAng = stAng1Dg + swAngDg;
                        */
                        d_val = `M${x1},${y1} L${xPos},${yPos} L${x2},${y2}${PPTXShapeUtils.shapeArcAlt(hc, vc, hc, vc, 0, 360, true)}`;// +
                        //PPTXShapeUtils.shapeArc(hc,vc,hc,vc,stAng1Dg,stAng1Dg+swAngDg,false).replace("M","L") +
                        //" z";
                        
                        result += `<path d='${d_val}'${cloudTransformAttr} fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "wedgeRectCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = -20833 * refr;
                        let sAdj2, adj2 = 62500 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        var dxPos, dyPos, xPos, yPos, dx, dy, dq, ady, adq, dz, xg1, xg2, x1, x2: any,
                            yg1, yg2, y1, y2: any, t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        dx = xPos - hc;
                        dy = yPos - vc;
                        dq = dxPos * h / w;
                        ady = Math.abs(dyPos);
                        adq = Math.abs(dq);
                        dz = ady - adq;
                        xg1 = (dxPos > 0) ? 7 : 2;
                        xg2 = (dxPos > 0) ? 10 : 5;
                        x1 = w * xg1 / 12;
                        x2 = w * xg2 / 12;
                        yg1 = (dyPos > 0) ? 7 : 2;
                        yg2 = (dyPos > 0) ? 10 : 5;
                        y1 = h * yg1 / 12;
                        y2 = h * yg2 / 12;
                        t1 = (dxPos > 0) ? 0 : xPos;
                        xl = (dz > 0) ? 0 : t1;
                        t2 = (dyPos > 0) ? x1 : xPos;
                        xt = (dz > 0) ? t2 : x1;
                        t3 = (dxPos > 0) ? xPos : w;
                        xr = (dz > 0) ? w : t3;
                        t4 = (dyPos > 0) ? xPos : x1;
                        xb = (dz > 0) ? t4 : x1;
                        t5 = (dxPos > 0) ? y1 : yPos;
                        yl = (dz > 0) ? y1 : t5;
                        t6 = (dyPos > 0) ? 0 : yPos;
                        yt = (dz > 0) ? t6 : 0;
                        t7 = (dxPos > 0) ? yPos : y1;
                        yr = (dz > 0) ? y1 : t7;
                        t8 = (dyPos > 0) ? yPos : h;
                        yb = (dz > 0) ? t8 : h;

                        d_val = `M${0},${0} L${x1},${0} L${xt},${yt} L${x2},${0} L${w},${0} L${w},${y1} L${xr},${yr} L${w},${y2} L${w},${h} L${x2},${h} L${xb},${yb} L${x1},${h} L${0},${h} L${0},${y2} L${xl},${yl} L${0},${y1} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "wedgeRoundRectCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = -20833 * refr;
                        let sAdj2, adj2 = 62500 * refr;
                        let sAdj3, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        const ss = Math.min(w, h);
                        var dxPos, dyPos, xPos, yPos, dq, ady, adq, dz, xg1, xg2, x1, x2: any, yg1, yg2, y1, y2: any,
                            t1, xl, t2, xt, t3, xr, t4, xb, t5, yl, t6, yt, t7, yr, t8, yb, u1, u2, v2,
                            vc = h / 2, hc = w / 2;
                        dxPos = w * adj1 / cnstVal1;
                        dyPos = h * adj2 / cnstVal1;
                        xPos = hc + dxPos;
                        yPos = vc + dyPos;
                        dq = dxPos * h / w;
                        ady = Math.abs(dyPos);
                        adq = Math.abs(dq);
                        dz = ady - adq;
                        xg1 = (dxPos > 0) ? 7 : 2;
                        xg2 = (dxPos > 0) ? 10 : 5;
                        x1 = w * xg1 / 12;
                        x2 = w * xg2 / 12;
                        yg1 = (dyPos > 0) ? 7 : 2;
                        yg2 = (dyPos > 0) ? 10 : 5;
                        y1 = h * yg1 / 12;
                        y2 = h * yg2 / 12;
                        t1 = (dxPos > 0) ? 0 : xPos;
                        xl = (dz > 0) ? 0 : t1;
                        t2 = (dyPos > 0) ? x1 : xPos;
                        xt = (dz > 0) ? t2 : x1;
                        t3 = (dxPos > 0) ? xPos : w;
                        xr = (dz > 0) ? w : t3;
                        t4 = (dyPos > 0) ? xPos : x1;
                        xb = (dz > 0) ? t4 : x1;
                        t5 = (dxPos > 0) ? y1 : yPos;
                        yl = (dz > 0) ? y1 : t5;
                        t6 = (dyPos > 0) ? 0 : yPos;
                        yt = (dz > 0) ? t6 : 0;
                        t7 = (dxPos > 0) ? yPos : y1;
                        yr = (dz > 0) ? y1 : t7;
                        t8 = (dyPos > 0) ? yPos : h;
                        yb = (dz > 0) ? t8 : h;
                        u1 = ss * adj3 / cnstVal1;
                        u2 = w - u1;
                        v2 = h - u1;
                        d_val = `M${0},${u1}${PPTXShapeUtils.shapeArc(u1, u1, u1, u1, 180, 270, false).replace("M", "L")} L${x1},${0} L${xt},${yt} L${x2},${0} L${u2},${0}${PPTXShapeUtils.shapeArc(u2, u1, u1, u1, 270, 360, false).replace("M", "L")} L${w},${y1} L${xr},${yr} L${w},${y2} L${w},${v2}${PPTXShapeUtils.shapeArc(u2, v2, u1, u1, 0, 90, false).replace("M", "L")} L${x2},${h} L${xb},${yb} L${x1},${h} L${u1},${h}${PPTXShapeUtils.shapeArc(u1, v2, u1, u1, 90, 180, false).replace("M", "L")} L${0},${y2} L${xl},${yl} L${0},${y1} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
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
                    case "callout3": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = 18750 * refr;
                        let sAdj2, adj2 = -8333 * refr;
                        let sAdj3, adj3 = 18750 * refr;
                        let sAdj4, adj4 = -16667 * refr;
                        let sAdj5, adj5 = 100000 * refr;
                        let sAdj6, adj6 = -16667 * refr;
                        let sAdj7, adj7 = 112963 * refr;
                        let sAdj8, adj8 = -8333 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * refr;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * refr;
                                } else if (sAdj_name == "adj6") {
                                    sAdj6 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj6 = parseInt(sAdj6.substr(4)) * refr;
                                } else if (sAdj_name == "adj7") {
                                    sAdj7 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj7 = parseInt(sAdj7.substr(4)) * refr;
                                } else if (sAdj_name == "adj8") {
                                    sAdj8 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj8 = parseInt(sAdj8.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 100000 * refr;
                        let isBorder = true;
                        switch (shapType) {
                            case "borderCallout1":
                            case "callout1":
                                if (shapType == "borderCallout1") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2: any, x2: any;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2}`;
                                break;
                            case "borderCallout2":
                            case "callout2":
                                if (shapType == "borderCallout2") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2: any, x2: any, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x2},${y2}`;

                                break;
                            case "borderCallout3":
                            case "callout3":
                                if (shapType == "borderCallout3") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;

                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;

                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2: any, x2: any, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;

                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;

                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4} L${x3},${y3} L${x2},${y2}`;
                                break;
                            case "accentBorderCallout1":
                            case "accentCallout1":
                                if (shapType == "accentBorderCallout1") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }

                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 112500 * refr;
                                    adj4 = -38333 * refr;
                                }
                                var y1, x1, y2: any, x2: any;
                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} M${x1},${0} L${x1},${h}`;
                                break;
                            case "accentBorderCallout2":
                            case "accentCallout2":
                                if (shapType == "accentBorderCallout2") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 112500 * refr;
                                    adj6 = -46667 * refr;
                                }
                                var y1, x1, y2: any, x2: any, y3, x3;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x2},${y2} M${x1},${0} L${x1},${h}`;

                                break;
                            case "accentBorderCallout3":
                            case "accentCallout3":
                                if (shapType == "accentBorderCallout3") {
                                    isBorder = true;
                                } else {
                                    isBorder = false;
                                }
                                isBorder = true;
                                if (shapAdjst_ary === undefined) {
                                    adj1 = 18750 * refr;
                                    adj2 = -8333 * refr;
                                    adj3 = 18750 * refr;
                                    adj4 = -16667 * refr;
                                    adj5 = 100000 * refr;
                                    adj6 = -16667 * refr;
                                    adj7 = 112963 * refr;
                                    adj8 = -8333 * refr;
                                }
                                var y1, x1, y2: any, x2: any, y3, x3, y4, x4;

                                y1 = h * adj1 / cnstVal1;
                                x1 = w * adj2 / cnstVal1;
                                y2 = h * adj3 / cnstVal1;
                                x2 = w * adj4 / cnstVal1;
                                y3 = h * adj5 / cnstVal1;
                                x3 = w * adj6 / cnstVal1;
                                y4 = h * adj7 / cnstVal1;
                                x4 = w * adj8 / cnstVal1;
                                d_val = `M${0},${0} L${w},${0} L${w},${h} L${0},${h} z M${x1},${y1} L${x2},${y2} L${x3},${y3} L${x4},${y4} L${x3},${y3} L${x2},${y2} M${x1},${0} L${x1},${h}`;
                                break;
                        }

                        //console.log("shapType: ", shapType, ",isBorder:", isBorder)
                        //if(isBorder){
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        //}else{
                        //    result += "<path d='"+d_val+"' fill='" + (!imgFillFlg?(grndFillFlg?"url(#linGrd_"+shpId+")":fillColor):"url(#imgPtrn_"+shpId+")") + 
                        //        "' stroke='none' />";

                        //}
                        break;
                    }
                    case "leftRightRibbon": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = 50000 * refr;
                        let sAdj2, adj2 = 50000 * refr;
                        let sAdj3, adj3 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * refr;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 33333 * refr;
                        const cnstVal2 = 100000 * refr;
                        const cnstVal3 = 200000 * refr;
                        const cnstVal4 = 400000 * refr;
                        const ss = Math.min(w, h);
                        var a3, maxAdj1, a1, w1, maxAdj2, a2, x1, x4, dy1, dy2, ly1, ry4, ly2, ry3, ly4, ry1,
                            ly3, ry2, hR, x2: any, x3, y1, y2: any, wd32 = w / 32, vc = h / 2, hc = w / 2;

                        a3 = (adj3 < 0) ? 0 : (adj3 > cnstVal1) ? cnstVal1 : adj3;
                        maxAdj1 = cnstVal2 - a3;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        w1 = hc - wd32;
                        maxAdj2 = cnstVal2 * w1 / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        x1 = ss * a2 / cnstVal2;
                        x4 = w - x1;
                        dy1 = h * a1 / cnstVal3;
                        dy2 = h * a3 / -cnstVal3;
                        ly1 = vc + dy2 - dy1;
                        ry4 = vc + dy1 - dy2;
                        ly2 = ly1 + dy1;
                        ry3 = h - ly2;
                        ly4 = ly2 * 2;
                        ry1 = h - ly4;
                        ly3 = ly4 - ly1;
                        ry2 = h - ly3;
                        hR = a3 * ss / cnstVal4;
                        x2 = hc - wd32;
                        x3 = hc + wd32;
                        y1 = ly1 + hR;
                        y2 = ry2 - hR;

                        d_val = `M${0},${ly2}L${x1},${0}L${x1},${ly1}L${hc},${ly1}${PPTXShapeUtils.shapeArcAlt(hc, y1, wd32, hR, 270, 450, false).replace!("M", "L")}${PPTXShapeUtils.shapeArcAlt(hc, y2, wd32, hR, 270, 90, false).replace!("M", "L")}L${x4},${ry2}L${x4},${ry1}L${w},${ry3}L${x4},${h}L${x4},${ry4}L${hc},${ry4}${PPTXShapeUtils.shapeArc(hc, ry4 - hR, wd32, hR, 90, 180, false).replace("M", "L")}L${x2},${ly3}L${x1},${ly3}L${x1},${ly4} zM${x3},${y1}L${x3},${ry2}M${x2},${y2}L${x2},${ly3}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "ribbon":
                    case "ribbon2": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 16667 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 25000 * SLIDE_FACTOR;
                        const cnstVal2 = 33333 * SLIDE_FACTOR;
                        const cnstVal3 = 75000 * SLIDE_FACTOR;
                        const cnstVal4 = 100000 * SLIDE_FACTOR;
                        const cnstVal5 = 200000 * SLIDE_FACTOR;
                        const cnstVal6 = 400000 * SLIDE_FACTOR;
                        let hc = w / 2, t = 0, l = 0, b: any = h, r: any = w, wd8 = w / 8, wd32 = w / 32;
                        var a1, a2, x10: any, dx2: any, x2: any, x9: any, x3, x8: any, x5, x6, x4, x7, y1, y2: any, y4, y3, hR, y6;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        x10 = r - wd8;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x9 = hc + dx2;
                        x3 = x2 + wd32;
                        x8 = x9 - wd32;
                        x5 = x2 + wd8;
                        x6 = x9 - wd8;
                        x4 = x5 - wd32;
                        x7 = x6 + wd32;
                        hR = h * a1 / cnstVal6;
                        if (shapType == "ribbon2") {
                            let dy1, dy2, y7;
                            dy1 = h * a1 / cnstVal5;
                            y1 = b - dy1;
                            dy2 = h * a1 / cnstVal4;
                            y2 = b - dy2;
                            y4 = t + dy2;
                            y3 = (y4 + b) / 2;
                            y6 = b - hR;///////////////////
                            y7 = y1 - hR;

                            d_val = `M${l},${b} L${wd8},${y3} L${l},${y4} L${x2},${y4} L${x2},${hR}${PPTXShapeUtils.shapeArcAlt(x3, hR, wd32, hR, 180, 270, false).replace!("M", "L")} L${x8},${t}${PPTXShapeUtils.shapeArcAlt(x8, hR, wd32, hR, 270, 360, false).replace!("M", "L")} L${x9},${y4} L${x9},${y4} L${r},${y4} L${x10},${y3} L${r},${b} L${x7},${b}${PPTXShapeUtils.shapeArc(x7, y6, wd32, hR, 90, 270, false).replace("M", "L")} L${x8},${y1}${PPTXShapeUtils.shapeArc(x8, y7, wd32, hR, 90, -90, false).replace("M", "L")} L${x3},${y2}${PPTXShapeUtils.shapeArc(x3, y7, wd32, hR, 270, 90, false).replace("M", "L")} L${x4},${y1}${PPTXShapeUtils.shapeArc(x4, y6, wd32, hR, 270, 450, false).replace("M", "L")} z M${x5},${y2} L${x5},${y6}M${x6},${y6} L${x6},${y2}M${x2},${y7} L${x2},${y4}M${x9},${y4} L${x9},${y7}`;
                        } else if (shapType == "ribbon") {
                            let y5;
                            y1 = h * a1 / cnstVal5;
                            y2 = h * a1 / cnstVal4;
                            y4 = b - y2;
                            y3 = y4 / 2;
                            y5 = b - hR; ///////////////////////
                            y6 = y2 - hR;
                            d_val = `M${l},${t} L${x4},${t}${PPTXShapeUtils.shapeArcAlt(x4, hR, wd32, hR, 270, 450, false).replace!("M", "L")} L${x3},${y1}${PPTXShapeUtils.shapeArcAlt(x3, y6, wd32, hR, 270, 90, false).replace!("M", "L")} L${x8},${y2}${PPTXShapeUtils.shapeArcAlt(x8, y6, wd32, hR, 90, -90, false).replace!("M", "L")} L${x7},${y1}${PPTXShapeUtils.shapeArcAlt(x7, hR, wd32, hR, 90, 270, false).replace!("M", "L")} L${r},${t} L${x10},${y3} L${r},${y4} L${x9},${y4} L${x9},${y5}${PPTXShapeUtils.shapeArc(x8, y5, wd32, hR, 0, 90, false).replace("M", "L")} L${x3},${b}${PPTXShapeUtils.shapeArc(x3, y5, wd32, hR, 90, 180, false).replace("M", "L")} L${x2},${y4} L${l},${y4} L${wd8},${y3} z M${x5},${hR} L${x5},${y2}M${x6},${y2} L${x6},${hR}M${x2},${y4} L${x2},${y6}M${x9},${y6} L${x9},${y4}`;
                        }
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "doubleWave":
                    case "wave": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = (shapType == "doubleWave") ? 6250 * SLIDE_FACTOR : 12500 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 0;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal2 = -10000 * SLIDE_FACTOR;
                        const cnstVal3 = 50000 * SLIDE_FACTOR;
                        const cnstVal4 = 100000 * SLIDE_FACTOR;
                        let hc = w / 2, t = 0, l = 0, b: any = h, r: any = w, wd8 = w / 8, wd32 = w / 32;
                        if (shapType == "doubleWave") {
                            const cnstVal1 = 12500 * SLIDE_FACTOR;
                            var a1, a2, y1, dy2, y2: any, y3, y4, y5, y6, of2, dx2: any, x2: any, dx8, x8: any, dx3, x3, dx4, x4, x5, x6, x7, x9: any, x15, x10: any, x11: any, x12, x13, x14;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal1) ? cnstVal1 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx8 = (of2 > 0) ? of2 : 0;
                            x8 = r - dx8;
                            dx3 = (dx2 + x8) / 6;
                            x3 = x2 + dx3;
                            dx4 = (dx2 + x8) / 3;
                            x4 = x2 + dx4;
                            x5 = (x2 + x8) / 2;
                            x6 = x5 + dx3;
                            x7 = (x6 + x8) / 2;
                            x9 = l + dx8;
                            x15 = r + dx2;
                            x10 = x9 + dx3;
                            x11 = x9 + dx4;
                            x12 = (x9 + x15) / 2;
                            x13 = x12 + dx3;
                            x14 = (x13 + x15) / 2;

                            d_val = `M${x2},${y1} C${x3},${y2} ${x4},${y3} ${x5},${y1} C${x6},${y2} ${x7},${y3} ${x8},${y1} L${x15},${y4} C${x14},${y6} ${x13},${y5} ${x12},${y4} C${x11},${y6} ${x10},${y5} ${x9},${y4} z`;
                        } else if (shapType == "wave") {
                            const cnstVal5 = 20000 * SLIDE_FACTOR;
                            var a1, a2, y1, dy2, y2: any, y3, y4, y5, y6, of2, dx2: any, x2: any, dx5, x5, dx3, x3, x4, x6, x10: any, x7, x8: any;
                            a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal5) ? cnstVal5 : adj1;
                            a2 = (adj2 < cnstVal2) ? cnstVal2 : (adj2 > cnstVal4) ? cnstVal4 : adj2;
                            y1 = h * a1 / cnstVal4;
                            dy2 = y1 * 10 / 3;
                            y2 = y1 - dy2;
                            y3 = y1 + dy2;
                            y4 = b - y1;
                            y5 = y4 - dy2;
                            y6 = y4 + dy2;
                            of2 = w * a2 / cnstVal3;
                            dx2 = (of2 > 0) ? 0 : of2;
                            x2 = l - dx2;
                            dx5 = (of2 > 0) ? of2 : 0;
                            x5 = r - dx5;
                            dx3 = (dx2 + x5) / 3;
                            x3 = x2 + dx3;
                            x4 = (x3 + x5) / 2;
                            x6 = l + dx5;
                            x10 = r + dx2;
                            x7 = x6 + dx3;
                            x8 = (x7 + x10) / 2;

                            d_val = `M${x2},${y1} C${x3},${y2} ${x4},${y3} ${x5},${y1} L${x10},${y4} C${x8},${y6} ${x7},${y5} ${x6},${y4} z`;
                        }
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "ellipseRibbon":
                    case "ellipseRibbon2": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 12500 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        let d_val;
                        const cnstVal1 = 25000 * SLIDE_FACTOR;
                        const cnstVal3 = 75000 * SLIDE_FACTOR;
                        const cnstVal4 = 100000 * SLIDE_FACTOR;
                        const cnstVal5 = 200000 * SLIDE_FACTOR;
                        let hc = w / 2, t = 0, l = 0, b: any = h, r: any = w, wd8 = w / 8;
                        var a1, a2, q10, q11, q12, minAdj3, a3, dx2: any, x2: any, x3, x4, x5, x6, dy1, f1, q1, q2,
                            cx1, cx2, q1, dy3, q3, q4, q5, rh, q8, cx4, q9, cx5;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal4) ? cnstVal4 : adj1;
                        a2 = (adj2 < cnstVal1) ? cnstVal1 : (adj2 > cnstVal3) ? cnstVal3 : adj2;
                        q10 = cnstVal4 - a1;
                        q11 = q10 / 2;
                        q12 = a1 - q11;
                        minAdj3 = (0 > q12) ? 0 : q12;
                        a3 = (adj3 < minAdj3) ? minAdj3 : (adj3 > a1) ? a1 : adj3;
                        dx2 = w * a2 / cnstVal5;
                        x2 = hc - dx2;
                        x3 = x2 + wd8;
                        x4 = r - x3;
                        x5 = r - x2;
                        x6 = r - wd8;
                        dy1 = h * a3 / cnstVal4;
                        f1 = 4 * dy1 / w;
                        q1 = x3 * x3 / w;
                        q2 = x3 - q1;
                        cx1 = x3 / 2;
                        cx2 = r - cx1;
                        q1 = h * a1 / cnstVal4;
                        dy3 = q1 - dy1;
                        q3 = x2 * x2 / w;
                        q4 = x2 - q3;
                        q5 = f1 * q4;
                        rh = b - q1;
                        q8 = dy1 * 14 / 16;
                        cx4 = x2 / 2;
                        q9 = f1 * cx4;
                        cx5 = r - cx4;
                        if (shapType == "ellipseRibbon") {
                            var y1, cy1, y3, q6, q7, cy3, y2: any, y5, y6,
                                cy4, cy6, y7, cy7, y8;
                            y1 = f1 * q2;
                            cy1 = f1 * cx1;
                            y3 = q5 + dy3;
                            q6 = dy1 + dy3 - y3;
                            q7 = q6 + dy1;
                            cy3 = q7 + dy3;
                            y2 = (q8 + rh) / 2;
                            y5 = q5 + rh;
                            y6 = y3 + rh;
                            cy4 = q9 + rh;
                            cy6 = cy3 + rh;
                            y7 = y1 + dy3;
                            cy7 = q1 + q1 - y7;
                            y8 = b - dy1;
                            //
                            d_val = `M${l},${t} Q${cx1},${cy1} ${x3},${y1} L${x2},${y3} Q${hc},${cy3} ${x5},${y3} L${x4},${y1} Q${cx2},${cy1} ${r},${t} L${x6},${y2} L${r},${rh} Q${cx5},${cy4} ${x5},${y5} L${x5},${y6} Q${hc},${cy6} ${x2},${y6} L${x2},${y5} Q${cx4},${cy4} ${l},${rh} L${wd8},${y2} zM${x2},${y5} L${x2},${y3}M${x5},${y3} L${x5},${y5}M${x3},${y1} L${x3},${y7}M${x4},${y7} L${x4},${y1}`;
                        } else if (shapType == "ellipseRibbon2") {
                            var u1, y1, cu1, cy1, q3, q5, u3, y3, q6, q7, cu3, cy3, rh, q8, u2, y2: any,
                                u5, y5, u6, y6, cu4, cy4, cu6, cy6, u7, y7, cu7, cy7;
                            u1 = f1 * q2;
                            y1 = b - u1;
                            cu1 = f1 * cx1;
                            cy1 = b - cu1;
                            u3 = q5 + dy3;
                            y3 = b - u3;
                            q6 = dy1 + dy3 - u3;
                            q7 = q6 + dy1;
                            cu3 = q7 + dy3;
                            cy3 = b - cu3;
                            u2 = (q8 + rh) / 2;
                            y2 = b - u2;
                            u5 = q5 + rh;
                            y5 = b - u5;
                            u6 = u3 + rh;
                            y6 = b - u6;
                            cu4 = q9 + rh;
                            cy4 = b - cu4;
                            cu6 = cu3 + rh;
                            cy6 = b - cu6;
                            u7 = u1 + dy3;
                            y7 = b - u7;
                            cu7 = q1 + q1 - u7;
                            cy7 = b - cu7;
                            //
                            d_val = `M${l},${b} L${wd8},${y2} L${l},${q1} Q${cx4},${cy4} ${x2},${y5} L${x2},${y6} Q${hc},${cy6} ${x5},${y6} L${x5},${y5} Q${cx5},${cy4} ${r},${q1} L${x6},${y2} L${r},${b} Q${cx2},${cy1} ${x4},${y1} L${x5},${y3} Q${hc},${cy3} ${x2},${y3} L${x3},${y1} Q${cx1},${cy1} ${l},${b} zM${x2},${y3} L${x2},${y5}M${x5},${y5} L${x5},${y3}M${x3},${y7} L${x3},${y1}M${x4},${y1} L${x4},${y7}`;
                        }
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "line":
                    case "straightConnector1":
                    case "bentConnector4":
                    case "bentConnector5": {
                        // 使用drawW和drawH（原始尺寸）而不是w和h（可能被调整的SVG容器尺寸）
                        let lineW = drawW;
                        let lineH = drawH;
                        // 如果drawW或drawH未定义（非连接器情况），回退到w和h
                        if (lineW === undefined) lineW = w;
                        if (lineH === undefined) lineH = h;
                        
                        // 根据flipH和flipV确定线条的起点和终点
                        var x1: any = 0, y1: any = 0, x2: any = lineW, y2: any = lineH;
                        
                        result += `<line x1='${x1}' y1='${y1}' x2='${x2}' y2='${y2}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' `;
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-start='url(#markerTriangle_${shpId})' `;
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-end='url(#markerTriangle_${shpId})' `;
                        }
                        result += "/>";
                        break;
                    }
                    case "curvedConnector2":
                    case "curvedConnector3":
                    case "curvedConnector4":
                    case "curvedConnector5": {
                        // 获取调整值
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let adj1 = 50000; // 默认值
                        if (shapAdjst_ary !== undefined) {
                            if (Array.isArray(shapAdjst_ary)) {
                                for (const i of shapAdjst_ary.keys()){
                                    const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                    if (sAdj_name == "adj1") {
                                        let sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                        adj1 = parseInt(sAdj1.substr(4));
                                        break;
                                    }
                                }
                            } else {
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    let sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary, ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4));
                                }
                            }
                        }

                        // 使用drawW和drawH（原始尺寸）
                        const curveW = (drawW !== undefined) ? drawW : w;
                        const curveH = (drawH !== undefined) ? drawH : h;

                        // 计算曲线控制点
                        let cx1, cy1, cx2, cy2;
                        let pathD;
                        
                        // 路径方向（SVG容器会通过flip变换处理翻转）
                        if (shapType === "curvedConnector2" || shapType === "curvedConnector3") {
                            // 对于 curvedConnector2 和 curvedConnector3，使用简单的二次贝塞尔曲线
                            const controlPointRatio = adj1 / 100000;
                            cx1 = curveW * controlPointRatio;
                            cy1 = 0;
                            cx2 = curveW * (1 - controlPointRatio);
                            cy2 = curveH;
                        } else {
                            // 对于其他弯曲连接器，使用默认控制点
                            cx1 = curveW / 4;
                            cy1 = 0;
                            cx2 = curveW * 3 / 4;
                            cy2 = curveH;
                        }
                        // 正常路径
                        pathD = `M 0,0 Q ${cx1},${cy1} ${curveW/2},${curveH/2} Q ${cx2},${cy2} ${curveW},${curveH}`;

                        // 使用 SVG 路径元素创建曲线
                        result += `<path d='${pathD}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' fill='none' `;
                        
                        if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-start='url(#markerTriangle_${shpId})' `;
                        }
                        if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                            result += `marker-end='url(#markerTriangle_${shpId})' `;
                        }
                        result += "/>";
                        break;
                    }
                    case "rightArrow":
                    case "leftArrow":
                    case "downArrow":
                    case "upArrow":
                    case "leftRightArrow":
                    case "upDownArrow": {
                        // 使用drawW和drawH（原始尺寸）而不是w和h（缩放后尺寸）
                        // SVG会通过transform scale()进行缩放
                        result += renderArrow(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node!);
                        break;
                    }
                    case "quadArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 22500 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 22500 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 22500 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2: any, dx2: any, x3, dx3, x4, x5, x6, y2: any, y3, y4, y5, y6, maxAdj1, maxAdj3;
                        const minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        y2 = vc - dx2;
                        y5 = vc + dx2;
                        y3 = vc - dx3;
                        y4 = vc + dx3;
                        y6 = h - x1;
                        let d_val = `M${0},${vc} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${hc},${0} L${x5},${x1} L${x4},${x1} L${x4},${y3} L${x6},${y3} L${x6},${y2} L${w},${vc} L${x6},${y5} L${x6},${y4} L${x4},${y4} L${x4},${y6} L${x5},${y6} L${hc},${h} L${x2},${y6} L${x3},${y6} L${x3},${y4} L${x1},${y4} L${x1},${y5} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "leftRightUpArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, q1, x1, x2: any, dx2: any, x3, dx3, x4, x5, x6, y2: any, dy2, y3, y4, y5, maxAdj1, maxAdj3;
                        const minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q1 = cnstVal2 - maxAdj1;
                        maxAdj3 = q1 / 2;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal2;
                        x2 = hc - dx2;
                        x5 = hc + dx2;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = hc - dx3;
                        x4 = hc + dx3;
                        x6 = w - x1;
                        dy2 = minWH * a2 / cnstVal1;
                        y2 = h - dy2;
                        y4 = h - dx2;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        let d_val = `M${0},${y4} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${hc},${0} L${x5},${x1} L${x4},${x1} L${x4},${y3} L${x6},${y3} L${x6},${y2} L${w},${y4} L${x6},${h} L${x6},${y5} L${x1},${y5} L${x1},${h} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "leftUpArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, x1, x2: any, dx4, dx3, x3, x4, x5, y2: any, y3, y4, y5, maxAdj1, maxAdj3;
                        const minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        maxAdj3 = cnstVal2 - maxAdj1;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        x1 = minWH * a3 / cnstVal2;
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        y2 = h - dx2;
                        dx4 = minWH * a2 / cnstVal2;
                        x4 = w - dx4;
                        y4 = h - dx4;
                        dx3 = minWH * a1 / cnstVal3;
                        x3 = x4 - dx3;
                        x5 = x4 + dx3;
                        y3 = y4 - dx3;
                        y5 = y4 + dx3;
                        let d_val = `M${0},${y4} L${x1},${y2} L${x1},${y3} L${x3},${y3} L${x3},${x1} L${x2},${x1} L${x4},${0} L${w},${x1} L${x5},${x1} L${x5},${y5} L${x1},${y5} L${x1},${h} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "bentUpArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, a1, a2, a3, dx1, x1, dx2: any, x2: any, dx3, x3, x4, y1, y2: any, dy2;
                        const minWH = Math.min(w, h);
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        if (adj3 < 0) a3 = 0
                        else if (maxAdj3 !== undefined && adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        y1 = minWH * a3! / cnstVal2;
                        dx1 = minWH * a2 / cnstVal1;
                        x1 = w - dx1;
                        dx3 = minWH * a2 / cnstVal2;
                        x3 = w - dx3;
                        dx2 = minWH * a1 / cnstVal3;
                        x2 = x3 - dx2;
                        x4 = x3 + dx2;
                        dy2 = minWH * a1 / cnstVal2;
                        y2 = h - dy2;
                        let d_val = `M${0},${y2} L${x2},${y2} L${x2},${y1} L${x1},${y1} L${x3},${0} L${w},${y1} L${x4},${y1} L${x4},${h} L${0},${h} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "bentArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 43750 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var a1, a2, a3, a4, x3, x4, y3, y4, y5, y6, maxAdj1, maxAdj4;
                        const minWH = Math.min(w, h);
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
                        var th, aw2, th2, dh2, ah, bw, bh, bs, bd, bd3, bd2,
                            th: any = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        ah = minWH * a3 / cnstVal2;
                        bw = w - ah;
                        bh = h - dh2;
                        bs = (bw < bh) ? bw : bh;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x4 = w - ah;
                        y3 = dh2 + th;
                        y4 = y3 + dh2;
                        y5 = dh2 + bd;
                        y6 = y3 + bd2;

                        let d_val = `M${0},${h} L${0},${y5}${PPTXShapeUtils.shapeArc(bd, y5, bd, bd, 180, 270, false).replace("M", "L")} L${x4},${dh2} L${x4},${0} L${w},${aw2} L${x4},${y4} L${x4},${y3} L${x3},${y3}${PPTXShapeUtils.shapeArc(x3, y6, bd2, bd2, 270, 180, false).replace("M", "L")} L${th},${h} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "uturnArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 43750 * SLIDE_FACTOR;
                        let sAdj5, adj5 = 75000 * SLIDE_FACTOR;
                        const cnstVal1 = 25000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var a1, a2, a3, a4, a5, q1, q2, q3, x3, x4, x5, x6, x7, x8: any, x9: any, y4, y5, minAdj5, maxAdj1, maxAdj3, maxAdj4;
                        const minWH = Math.min(w, h);
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > cnstVal1) a2 = cnstVal1
                        else a2 = adj2
                        maxAdj1 = 2 * a2;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > maxAdj1) a1 = maxAdj1
                        else a1 = adj1
                        q2 = a1 * minWH / h;
                        q3 = cnstVal2 - q2;
                        maxAdj3 = q3 * h / minWH;
                        if (adj3 < 0) a3 = 0
                        else if (adj3 > maxAdj3) a3 = maxAdj3
                        else a3 = adj3
                        q1 = a3 + a1;
                        minAdj5 = q1 * minWH / h;
                        if (adj5 < minAdj5) a5 = minAdj5
                        else if (adj5 > cnstVal2) a5 = cnstVal2
                        else a5 = adj5

                        var th, aw2, th2, dh2, ah, bw, bs, bd, bd3, bd2,
                            th: any = minWH * a1 / cnstVal2;
                        aw2 = minWH * a2 / cnstVal2;
                        th2 = th / 2;
                        dh2 = aw2 - th2;
                        y5 = h * a5 / cnstVal2;
                        ah = minWH * a3 / cnstVal2;
                        y4 = y5 - ah;
                        x9 = w - dh2;
                        bw = x9 / 2;
                        bs = (bw < y4) ? bw : y4;
                        maxAdj4 = cnstVal2 * bs / minWH;
                        if (adj4 < 0) a4 = 0
                        else if (adj4 > maxAdj4) a4 = maxAdj4
                        else a4 = adj4
                        bd = minWH * a4 / cnstVal2;
                        bd3 = bd - th;
                        bd2 = (bd3 > 0) ? bd3 : 0;
                        x3 = th + bd2;
                        x8 = w - aw2;
                        x6 = x8 - aw2;
                        x7 = x6 + dh2;
                        x4 = x9 - bd;
                        x5 = x7 - bd2;
                        var cx = (th + x7) / 2
                        var cy = (y4 + th) / 2
                        let d_val = `M${0},${h} L${0},${bd}${shapeArcAlt!(bd, bd, bd, bd, 180, 270, false).replace!("M", "L")} L${x4},${0}${shapeArcAlt!(x4, bd, bd, bd, 270, 360, false).replace!("M", "L")} L${x9},${y4} L${w},${y4} L${x8},${y5} L${x6},${y4} L${x7},${y4} L${x7},${x3}${shapeArcAlt!(x5, x3, bd2, bd2, 0, -90, false).replace!("M", "L")} L${x3},${th}${shapeArcAlt!(x3, x3, bd2, bd2, 270, 180, false).replace!("M", "L")} L${th},${h} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "stripedRightArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 50000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        const cnstVal2 = 200000 * SLIDE_FACTOR;
                        const cnstVal3 = 84375 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var a1, a2, x4, x5, dx5, x6, dx6, y1, dy1, y2: any, maxAdj2, vc = h / 2;
                        const minWH = Math.min(w, h);
                        maxAdj2 = cnstVal3 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        x4 = minWH * 5 / 32;
                        dx5 = minWH * a2 / cnstVal1;
                        x5 = w - dx5;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        //dx6 = dy1*dx5/hd2;
                        //x6 = w-dx6;
                        const ssd8 = minWH / 8,
                            ssd16 = minWH / 16,
                            ssd32 = minWH / 32;
                        let d_val = `M${0},${y1} L${ssd32},${y1} L${ssd32},${y2} L${0},${y2} z M${ssd16},${y1} L${ssd8},${y1} L${ssd8},${y2} L${ssd16},${y2} z M${x4},${y1} L${x5},${y1} L${x5},${0} L${w},${vc} L${x5},${h} L${x5},${y2} L${x4},${y2} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "notchedRightArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 50000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        const cnstVal2 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var a1, a2, x1, x2: any, dx2: any, y1, dy1, y2: any, maxAdj2, vc = h / 2, hd2: any = vc;
                        const minWH = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / minWH;
                        if (adj1 < 0) a1 = 0
                        else if (adj1 > cnstVal1) a1 = cnstVal1
                        else a1 = adj1
                        if (adj2 < 0) a2 = 0
                        else if (adj2 > maxAdj2) a2 = maxAdj2
                        else a2 = adj2
                        dx2 = minWH * a2 / cnstVal1;
                        x2 = w - dx2;
                        dy1 = h * a1 / cnstVal2;
                        y1 = vc - dy1;
                        y2 = vc + dy1;
                        x1 = dy1 * dx2 / hd2;
                        let d_val = `M${0},${y1} L${x2},${y1} L${x2},${0} L${w},${vc} L${x2},${h} L${x2},${y2} L${0},${y2} L${x1},${vc} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "homePlate": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 50000 * SLIDE_FACTOR;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var a, x1, dx1, maxAdj, vc = h / 2;
                        const minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        dx1 = minWH * a / cnstVal1;
                        x1 = w - dx1;
                        let d_val = `M${0},${0} L${x1},${0} L${w},${vc} L${x1},${h} L${0},${h} z`;

                        result += `<path  d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "chevron": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 50000 * SLIDE_FACTOR;
                        const cnstVal1 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        var a, x1, dx1, x2: any, maxAdj, vc = h / 2;
                        const minWH = Math.min(w, h);
                        maxAdj = cnstVal1 * w / minWH;
                        if (adj < 0) a = 0
                        else if (adj > maxAdj) a = maxAdj
                        else a = adj
                        x1 = minWH * a / cnstVal1;
                        x2 = w - x1;
                        let d_val = `M${0},${0} L${x2},${0} L${w},${vc} L${x2},${h} L${0},${h} L${x1},${vc} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;


                        break;
                    }
                    case "rightArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 64977 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2: any, y3, y4, dx3, x3, x2: any, x1;
                        let vc = h / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        dx3 = ss * a3 / cnstVal2;
                        x3 = r - dx3;
                        x2 = w * a4 / cnstVal2;
                        x1 = x2 / 2;
                        let d_val = `M${l},${t} L${x2},${t} L${x2},${y2} L${x3},${y2} L${x3},${y1} L${r},${vc} L${x3},${y4} L${x3},${y3} L${x2},${y3} L${x2},${b} L${l},${b} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "downArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 64977 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2: any, x1, x2: any, x3, x4, dy3, y3, y2: any, y1;
                        let hc = w / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        dy3 = ss * a3 / cnstVal2;
                        y3 = b - dy3;
                        y2 = h * a4 / cnstVal2;
                        y1 = y2 / 2;
                        let d_val = `M${l},${t} L${r},${t} L${r},${y2} L${x3},${y2} L${x3},${y3} L${x4},${y3} L${hc},${b} L${x1},${y3} L${x2},${y3} L${x2},${y2} L${l},${y2} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "leftArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 64977 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2: any, y3, y4, x1, dx2: any, x2: any, x3;
                        let vc = h / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);

                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / w;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        dx2 = w * a4 / cnstVal2;
                        x2 = r - dx2;
                        x3 = (x2 + r) / 2;
                        let d_val = `M${l},${vc} L${x1},${y1} L${x1},${y2} L${x2},${y2} L${x2},${t} L${r},${t} L${r},${b} L${x2},${b} L${x2},${y3} L${x1},${y3} L${x1},${y4} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "upArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 64977 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx1, dx2: any, x1, x2: any, x3, x4, y1, dy2, y2: any, y3;
                        let hc = w / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal2 * h / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / h;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx1 = ss * a2 / cnstVal2;
                        dx2 = ss * a1 / cnstVal3;
                        x1 = hc - dx1;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        x4 = hc + dx1;
                        y1 = ss * a3 / cnstVal2;
                        dy2 = h * a4 / cnstVal2;
                        y2 = b - dy2;
                        y3 = (y2 + b) / 2;

                        let d_val = `M${l},${y2} L${x2},${y2} L${x2},${y1} L${x1},${y1} L${hc},${t} L${x4},${y1} L${x3},${y1} L${x3},${y2} L${r},${y2} L${r},${b} L${l},${b} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "leftRightArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 25000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 48123 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var maxAdj2, a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dy1, dy2, y1, y2: any, y3, y4, x1, x4, dx2: any, x2: any, x3;
                        let vc = h / 2, hc = w / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);
                        maxAdj2 = cnstVal1 * h / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 * w / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * ss / wd2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < 0) ? 0 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dy1 = ss * a2 / cnstVal2;
                        dy2 = ss * a1 / cnstVal3;
                        y1 = vc - dy1;
                        y2 = vc - dy2;
                        y3 = vc + dy2;
                        y4 = vc + dy1;
                        x1 = ss * a3 / cnstVal2;
                        x4 = r - x1;
                        dx2 = w * a4 / cnstVal3;
                        x2 = hc - dx2;
                        x3 = hc + dx2;
                        let d_val = `M${l},${vc} L${x1},${y1} L${x1},${y2} L${x2},${y2} L${x2},${t} L${x3},${t} L${x3},${y2} L${x4},${y2} L${x4},${y1} L${r},${vc} L${x4},${y4} L${x4},${y3} L${x3},${y3} L${x3},${b} L${x2},${b} L${x2},${y3} L${x1},${y3} L${x1},${y4} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "quadArrowCallout": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 18515 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 18515 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 18515 * SLIDE_FACTOR;
                        let sAdj4, adj4 = 48123 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const cnstVal3 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = parseInt(sAdj4.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        let vc = h / 2, hc = w / 2, r: any = w, b: any = h, l = 0, t = 0;
                        const ss = Math.min(w, h);
                        var a2, maxAdj1, a1, maxAdj3, a3, q2, maxAdj4, a4, dx2: any, dx3, ah, dx1, dy1, x8: any, x2: any, x7, x3, x6, x4, x5, y8, y2: any, y7, y3, y6, y4, y5;
                        a2 = (adj2 < 0) ? 0 : (adj2 > cnstVal1) ? cnstVal1 : adj2;
                        maxAdj1 = a2 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        maxAdj3 = cnstVal1 - a2;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        q2 = a3 * 2;
                        maxAdj4 = cnstVal2 - q2;
                        a4 = (adj4 < a1) ? a1 : (adj4 > maxAdj4) ? maxAdj4 : adj4;
                        dx2 = ss * a2 / cnstVal2;
                        dx3 = ss * a1 / cnstVal3;
                        ah = ss * a3 / cnstVal2;
                        dx1 = w * a4 / cnstVal3;
                        dy1 = h * a4 / cnstVal3;
                        x8 = r - ah;
                        x2 = hc - dx1;
                        x7 = hc + dx1;
                        x3 = hc - dx2;
                        x6 = hc + dx2;
                        x4 = hc - dx3;
                        x5 = hc + dx3;
                        y8 = b - ah;
                        y2 = vc - dy1;
                        y7 = vc + dy1;
                        y3 = vc - dx2;
                        y6 = vc + dx2;
                        y4 = vc - dx3;
                        y5 = vc + dx3;
                        let d_val = `M${l},${vc} L${ah},${y3} L${ah},${y4} L${x2},${y4} L${x2},${y2} L${x4},${y2} L${x4},${ah} L${x3},${ah} L${hc},${t} L${x6},${ah} L${x5},${ah} L${x5},${y2} L${x7},${y2} L${x7},${y4} L${x8},${y4} L${x8},${y3} L${r},${vc} L${x8},${y6} L${x8},${y5} L${x7},${y5} L${x7},${y7} L${x5},${y7} L${x5},${y8} L${x6},${y8} L${hc},${b} L${x3},${y8} L${x4},${y8} L${x4},${y7} L${x2},${y7} L${x2},${y5} L${ah},${y5} L${ah},${y6} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "curvedDownArrow": {
                        // 下弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        let cw = (drawW !== undefined) ? drawW : w;
                        let ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, wd2: any = cw / 2, r: any = cw, b: any = ch, l = 0, t = 0, c3d4 = 270, cd2: any = 180, cd4: any = 90;
                        const ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8: any, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, stAng, stAng2, swAng2, swAng3;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num: any) {
                            return parseFloat(num.toFixed(2));
                        }

                        maxAdj2 = cnstVal1 * cw / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * ch / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = ch * ch;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / ch;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = b - ah;
                        swAng = Math.atan(dx / ah);
                        const swAngDeg = swAng * 180 / Math.PI;
                        mswAng = -swAngDeg;
                        iy = b - idy;
                        ix = (wR + x3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        const dang2Deg = dang2 * 180 / Math.PI;
                        stAng = c3d4 + swAngDeg;
                        stAng2 = c3d4 - dang2Deg;
                        swAng2 = dang2Deg - cd4;
                        swAng3 = cd4 + dang2Deg;
                        //var cX = x5 - Math.cos(stAng*Math.PI/180) * wR;
                        //var cY = y1 - Math.sin(stAng*Math.PI/180) * h;

                        // 格式化所有坐标值
                        x6 = fmt(x6);
                        b = fmt(b);
                        x4 = fmt(x4);
                        y1 = fmt(y1);
                        x5 = fmt(x5);
                        x3 = fmt(x3);
                        t = fmt(t);
                        th = fmt(th);
                        x8 = fmt(x8);
                        wR = fmt(wR);
                        ch = fmt(ch);

                        let d_val = `M${x6},${b} L${x4},${y1} L${x5},${y1}${PPTXShapeUtils.shapeArc(wR, ch, wR, ch, stAng, (stAng + mswAng), false).replace("M", "L")} L${x3},${t}${PPTXShapeUtils.shapeArc(x3, ch, wR, ch, c3d4, (c3d4 + swAngDeg), false).replace("M", "L")} L${fmt(x5 + th)},${y1} L${x8},${y1} zM${x3},${t}${PPTXShapeUtils.shapeArc(x3, ch, wR, ch, stAng2, (stAng2 + swAng2), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(wR, ch, wR, ch, cd2, (cd2 + swAng3), false).replace("M", "L")}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "curvedLeftArrow": {
                        // 左弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        let cw = (drawW !== undefined) ? drawW : w;
                        let ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, hd2: any = ch / 2, r: any = cw, b: any = ch, l = 0, t = 0, c3d4 = 270, cd2: any = 180, cd4: any = 90;
                        const ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy, y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num: any) {
                            return parseFloat(num.toFixed(2));
                        }

                        maxAdj2 = cnstVal1 * ch / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * cw / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = cw * cw;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / cw;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = l + ah;
                        swAng = Math.atan(dy / ah);
                        mswAng = -swAng;
                        ix = l + iDx;
                        iy = (hR + y3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng3 = -dang2;
                        var swAngDg, swAng2Dg, swAng3Dg, stAng3dg;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2Dg = swAng2 * 180 / Math.PI;
                        swAng3Dg = swAng3 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;

                        // 格式化所有坐标值
                        r = fmt(r);
                        y3 = fmt(y3);
                        l = fmt(l);
                        hR = fmt(hR);
                        cw = fmt(cw);
                        t = fmt(t);
                        x1 = fmt(x1);
                        y7 = fmt(y7);
                        y8 = fmt(y8);
                        y6 = fmt(y6);
                        y4 = fmt(y4);
                        y5 = fmt(y5);

                        let d_val = `M${r},${y3}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L")} L${l},${t}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L")} L${r},${y3}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, 0, swAngDg, false).replace("M", "L")} L${x1},${y7} L${x1},${y8} L${l},${y6} L${x1},${y4} L${x1},${y5}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, swAngDg, (swAngDg + swAng2Dg), false).replace("M", "L")}${PPTXShapeUtils.shapeArc(l, hR, cw, hR, 0, -cd4, false).replace("M", "L")}${PPTXShapeUtils.shapeArc(l, y3, cw, hR, c3d4, (c3d4 + cd4), false).replace("M", "L")}`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "curvedRightArrow": {
                        /**
                         * curvedRightArrow: 手杖形箭头（弯曲向右的箭头）
                         *
                         * 形状说明：
                         * - 从左侧开始，向上弯曲，最后指向右侧
                         * - 有箭头头部
                         *
                         * 参数说明：
                         * - adj1: 控制箭头尖端的高度
                         * - adj2: 控制箭头宽度
                         * - adj3: 控制弯曲程度（箭头宽度）
                         */
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        let cw = (drawW !== undefined) ? drawW : w;
                        let ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, hd2: any = ch / 2, r: any = cw, b: any = ch, l = 0, t = 0, c3d4 = 270, cd2: any = 180, cd4: any = 90;
                        const ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, hR, q7, q8, q9, q10, q11, iDx, maxAdj3, a3, ah, y3, q2, q3, q4, q5, dy,
                            y5, y7, q6, dh, y4, y8, aw2, y6, x1, swAng, stAng, mswAng, ix, iy, q12, dang2, swAng2, swAng3, stAng3;

                        maxAdj2 = cnstVal1 * ch / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > a2) ? a2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        hR = hd2 - q1;
                        q7 = hR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        iDx = q11 * cw / q7;
                        maxAdj3 = cnstVal2 * iDx / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * a3 / cnstVal2;
                        y3 = hR + th;
                        q2 = cw * cw;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dy = q5 * hR / cw;
                        y5 = hR + dy;
                        y7 = y3 + dy;
                        q6 = aw - th;
                        dh = q6 / 2;
                        y4 = y5 - dh;
                        y8 = y7 + dh;
                        aw2 = aw / 2;
                        y6 = b - aw2;
                        x1 = r - ah;
                        swAng = Math.atan(dy / ah);
                        stAng = Math.PI + 0 - swAng;
                        mswAng = -swAng;
                        ix = r - iDx;
                        iy = (hR + y3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / iDx);
                        swAng2 = dang2 - Math.PI / 2;
                        swAng3 = Math.PI / 2 + dang2;
                        stAng3 = Math.PI - dang2;

                        var stAngDg, mswAngDg, swAngDg, swAng2dg;
                        stAngDg = stAng * 180 / Math.PI;
                        mswAngDg = mswAng * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;

                        /**
                         * 路径绘制顺序（参考 pptxjs.js）：
                         * 1. 从左侧 (l, hR) 开始，画第一个大圆弧
                         * 2. 画箭头下翼（多条线段）
                         * 3. 连接到箭头上翼
                         * 4. 画第二个大圆弧（沿上边缘）
                         * 5. 画第三个圆弧（箭头部分）
                         * 6. 闭合
                         */
                        let d_val = `M${l},${hR}${shapeArcAlt!(cw, hR, cw, hR, cd2, cd2 + mswAngDg, false).replace!("M", "L")} L${x1},${y5} L${x1},${y4} L${r},${y6} L${x1},${y8} L${x1},${y7}${shapeArcAlt!(cw, y3, cw, hR, stAngDg, stAngDg + swAngDg, false).replace!("M", "L")} L${l},${hR}${shapeArcAlt!(cw, hR, cw, hR, cd2, cd2 + cd4, false).replace!("M", "L")} L${r},${th}${shapeArcAlt!(cw, y3, cw, hR, c3d4, c3d4 + swAng2dg, false).replace!("M", "L")} z`;

                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "curvedUpArrow": {
                        // 上弧形箭头使用drawW和drawH（原始尺寸）进行形状计算
                        // 这样在group-abs类型组合中不会被缩放影响
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 25000 * SLIDE_FACTOR;
                        let sAdj2, adj2 = 50000 * SLIDE_FACTOR;
                        let sAdj3, adj3 = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = parseInt(sAdj3.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        // 使用drawW和drawH进行形状计算
                        let cw = (drawW !== undefined) ? drawW : w;
                        let ch = (drawH !== undefined) ? drawH : h;
                        var vc = ch / 2, hc = cw / 2, wd2: any = cw / 2, r: any = cw, b: any = ch, l = 0, t = 0, c3d4 = 270, cd2: any = 180, cd4: any = 90;
                        const ss = Math.min(cw, ch);
                        var maxAdj2, a2, a1, th, aw, q1, wR, q7, q8, q9, q10, q11, idy, maxAdj3, a3, ah, x3, q2, q3, q4, q5, dx, x5, x7, q6, dh, x4, x8: any, aw2, x6, y1, swAng, mswAng, iy, ix, q12, dang2, swAng2, mswAng2, stAng3, swAng3, stAng2;

                        // 辅助函数：格式化数字为2位小数
                        function fmt(num: any) {
                            return parseFloat(num.toFixed(2));
                        }

                        // 辅助函数：格式化弧线路径中的所有坐标
                        function fmtArc(arcStr: any) {
                            return arcStr.replace(/[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?/g, (match: any) => {
    return fmt(parseFloat(match)).toString();
});
                        }

                        maxAdj2 = cnstVal1 * cw / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > cnstVal2) ? cnstVal2 : adj1;
                        th = ss * a1 / cnstVal2;
                        aw = ss * a2 / cnstVal2;
                        q1 = (th + aw) / 4;
                        wR = wd2 - q1;
                        q7 = wR * 2;
                        q8 = q7 * q7;
                        q9 = th * th;
                        q10 = q8 - q9;
                        q11 = Math.sqrt(q10);
                        idy = q11 * ch / q7;
                        maxAdj3 = cnstVal2 * idy / ss;
                        a3 = (adj3 < 0) ? 0 : (adj3 > maxAdj3) ? maxAdj3 : adj3;
                        ah = ss * adj3 / cnstVal2;
                        x3 = wR + th;
                        q2 = ch * ch;
                        q3 = ah * ah;
                        q4 = q2 - q3;
                        q5 = Math.sqrt(q4);
                        dx = q5 * wR / ch;
                        x5 = wR + dx;
                        x7 = x3 + dx;
                        q6 = aw - th;
                        dh = q6 / 2;
                        x4 = x5 - dh;
                        x8 = x7 + dh;
                        aw2 = aw / 2;
                        x6 = r - aw2;
                        y1 = t + ah;
                        swAng = Math.atan(dx / ah);
                        mswAng = -swAng;
                        iy = t + idy;
                        ix = (wR + x3) / 2;
                        q12 = th / 2;
                        dang2 = Math.atan(q12 / idy);
                        swAng2 = dang2 - swAng;
                        mswAng2 = -swAng2;
                        stAng3 = Math.PI / 2 - swAng;
                        swAng3 = swAng + dang2;
                        stAng2 = Math.PI / 2 - dang2;

                        var stAng2dg, swAng2dg, swAngDg, swAng2dg;
                        stAng2dg = stAng2 * 180 / Math.PI;
                        swAng2dg = swAng2 * 180 / Math.PI;
                        stAng3dg = stAng3 * 180 / Math.PI;
                        swAngDg = swAng * 180 / Math.PI;

                        // 格式化所有坐标值
                        wR = fmt(wR);
                        ch = fmt(ch);
                        cw = fmt(cw);
                        x3 = fmt(x3);
                        x5 = fmt(x5);
                        x7 = fmt(x7);
                        x4 = fmt(x4);
                        x8 = fmt(x8);
                        x6 = fmt(x6);
                        y1 = fmt(y1);
                        b = fmt(b);
                        th = fmt(th);
                        t = fmt(t);

                        let d_val = //"M" + ix + "," +iy +
                            `${fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, stAng2dg, stAng2dg + swAng2dg, false))} L${x5},${y1} L${x4},${y1} L${x6},${t} L${x8},${y1} L${x7},${y1}${fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, stAng3dg, stAng3dg + swAngDg, false)).replace("M", "L")} L${wR},${b}${fmtArc(PPTXShapeUtils.shapeArc(wR, 0, wR, ch, cd4, cd2, false)).replace("M", "L")} L${th},${t}${fmtArc(PPTXShapeUtils.shapeArc(x3, 0, wR, ch, cd2, cd4, false)).replace("M", "L")}`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "mathDivide":
                    case "mathEqual":
                    case "mathMinus":
                    case "mathMultiply":
                    case "mathNotEqual":
                    case "mathPlus": {
                        // 使用drawW和drawH（原始尺寸）进行形状计算
                        result += renderMathSymbol(shapType, drawW, drawH, imgFillFlg, grndFillFlg, fillColor, border, shpId, node!);
                        break;
                    }
                    case "cylinder":
                    case "can":
                    case "flowChartMagneticDisk":
                    case "flowChartMagneticDrum": {
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 25000 * SLIDE_FACTOR;
                        const cnstVal1 = 50000 * SLIDE_FACTOR;
                        const cnstVal2 = 200000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        const ss = Math.min(w, h);
                        var maxAdj, a, y1, y2: any, y3, dVal;
                        if (shapType == "flowChartMagneticDisk" || shapType == "flowChartMagneticDrum") {
                            adj = 50000 * SLIDE_FACTOR;
                        }
                        maxAdj = cnstVal1 * h / ss;
                        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
                        y1 = ss * a / cnstVal2;
                        y2 = y1 + y1;
                        y3 = h - y1;
                        var cd2: any = 180, wd2: any = w / 2;

                        let tranglRott = "";
                        if (shapType == "flowChartMagneticDrum") {
                            tranglRott = `transform='rotate(90 ${w / 2},${h / 2})'`;
                        }

                        // 使用 shapeArcAlt，参数是半径而非直径（参考 pptxjs.js）
                        dVal = `${shapeArcAlt(wd2, y1, wd2, y1, 0, cd2, false)}${shapeArcAlt!(wd2, y1, wd2, y1, cd2, cd2 + cd2, false).replace!("M", "L")} L${w},${y3}${shapeArcAlt!(wd2, y3, wd2, y1, 0, cd2, false).replace!("M", "L")} L${0},${y1}`;

                        result += `<path ${tranglRott} d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "swooshArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        const refr = SLIDE_FACTOR;
                        let sAdj1, adj1 = 25000 * refr;
                        let sAdj2, adj2 = 16667 * refr;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * refr;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = parseInt(sAdj2.substr(4)) * refr;
                                }
                            }
                        }
                        const cnstVal1 = 1 * refr;
                        const cnstVal2 = 70000 * refr;
                        const cnstVal3 = 75000 * refr;
                        const cnstVal4 = 100000 * refr;
                        const ss = Math.min(w, h);
                        const ssd8 = ss / 8;
                        const hd6 = h / 6;

                        let a1, maxAdj2, a2, ad1, ad2, xB, yB, alfa, dx0, xC, dx1, yF, xF, xE, yE, dy2, dy22, dy3, yD, dy4, yP1, xP1, dy5, yP2, xP2;

                        a1 = (adj1 < cnstVal1) ? cnstVal1 : (adj1 > cnstVal3) ? cnstVal3 : adj1;
                        maxAdj2 = cnstVal2 * w / ss;
                        a2 = (adj2 < 0) ? 0 : (adj2 > maxAdj2) ? maxAdj2 : adj2;
                        ad1 = h * a1 / cnstVal4;
                        ad2 = ss * a2 / cnstVal4;
                        xB = w - ad2;
                        yB = ssd8;
                        alfa = (Math.PI / 2) / 14;
                        dx0 = ssd8 * Math.tan(alfa);
                        xC = xB - dx0;
                        dx1 = ad1 * Math.tan(alfa);
                        yF = yB + ad1;
                        xF = xB + dx1;
                        xE = xF + dx0;
                        yE = yF + ssd8;
                        dy2 = yE - 0;
                        dy22 = dy2 / 2;
                        dy3 = h / 20;
                        yD = dy22 - dy3;
                        dy4 = hd6;
                        yP1 = hd6 + dy4;
                        xP1 = w / 6;
                        dy5 = hd6 / 2;
                        yP2 = yF + dy5;
                        xP2 = w / 4;

                        let dVal = `M${0},${h} Q${xP1},${yP1} ${xB},${yB} L${xC},${0} L${w},${yD} L${xE},${yE} L${xF},${yF} Q${xP2},${yP2} ${0},${h} z`;

                        result += `<path d='${dVal}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "circularArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 12500 * SLIDE_FACTOR;
                        let sAdj2, adj2 = (1142319 / 60000) * Math.PI / 180;
                        let sAdj3, adj3 = (20457681 / 60000) * Math.PI / 180;
                        let sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        let sAdj5, adj5 = 12500 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r: any = w, b: any = h, l = 0, t = 0, wd2: any = w / 2, hd2: any = h / 2;
                        const ss = Math.min(w, h);
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH,
                            dyH, xH, yH, rI, u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17,
                            u18, u19, u20, u21, maxAng, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE,
                            dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO,
                            q13, q14, dyF1, q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF,
                            x1I, y1I, x2I, y2I, dxI, dyI, dI, v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11,
                            dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16, v17, v18, v19, v20, v21, v22, dxC, dyC,
                            sdxC, sdyC, xC, yC, ist0, ist1, istAng, isw1, isw2, iswAng, p1, p2, p3, p4, p5, xGp, yGp,
                            xBp, yBp, en0, en1, en2, sw0, sw1, swAng;
                        const cnstVal1 = 25000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        const rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        const rdAngVal3 = 2 * Math.PI;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4; //////////////////////////////////////////
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);

                        //dxH = rw3*Math.cos(Math.atan(wtH/htH));
                        //dyH = rh3*Math.sin(Math.atan(wtH/htH));
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));

                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;

                        //u13 = Math.atan(u12/1);
                        u13 = Math.atan2(u12, 1);

                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        maxAng = Math.abs(u21);
                        aAng = (adj2 < 0) ? 0 : (adj2 > maxAng) ? maxAng : adj2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        //dxA = rw3*Math.cos(Math.atan(wtA/htA));
                        //dyA = rh3*Math.sin(Math.atan(wtA/htA));
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));

                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);

                        //dxE = rw1*Math.cos(Math.atan(wtE/htE));
                        //dyE = rh1*Math.sin(Math.atan(wtE/htE));
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));

                        xE = hc + dxE;
                        yE = vc + dyE;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;

                        //ist0 = Math.atan(sdyC/sdxC);
                        ist0 = Math.atan2(sdyC, sdxC);

                        ist1 = ist0 + rdAngVal3;
                        istAng = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng;
                        isw2 = isw1 - rdAngVal3;
                        iswAng = (isw1 > 0) ? isw2 : isw1;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;

                        //en0 = Math.atan(sdyF/sdxF);
                        en0 = Math.atan2(sdyF, sdxF);

                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 + rdAngVal3;
                        swAng = (sw0 > 0) ? sw0 : sw1;

                        const strtAng = stAng * 180 / Math.PI
                        const endAng = strtAng + (swAng * 180 / Math.PI);
                        const stiAng = istAng * 180 / Math.PI;
                        const swiAng = iswAng * 180 / Math.PI;
                        const ediAng = stiAng + swiAng;

                        let d_val = `${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false)} L${xGp},${yGp} L${xA},${yA} L${xBp},${yBp} L${xC},${yC}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L")} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "leftCircularArrow": {
                        const shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
                        let sAdj1, adj1 = 12500 * SLIDE_FACTOR;
                        let sAdj2, adj2 = (-1142319 / 60000) * Math.PI / 180;
                        let sAdj3, adj3 = (1142319 / 60000) * Math.PI / 180;
                        let sAdj4, adj4 = (10800000 / 60000) * Math.PI / 180;
                        let sAdj5, adj5 = 12500 * SLIDE_FACTOR;
                        if (shapAdjst_ary !== undefined) {
                            for (const i of shapAdjst_ary.keys()){
                                const sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                                if (sAdj_name == "adj1") {
                                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj1 = parseInt(sAdj1.substr(4)) * SLIDE_FACTOR;
                                } else if (sAdj_name == "adj2") {
                                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj2 = (parseInt(sAdj2.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj3") {
                                    sAdj3 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj3 = (parseInt(sAdj3.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj4") {
                                    sAdj4 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj4 = (parseInt(sAdj4.substr(4)) / 60000) * Math.PI / 180;
                                } else if (sAdj_name == "adj5") {
                                    sAdj5 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                                    adj5 = parseInt(sAdj5.substr(4)) * SLIDE_FACTOR;
                                }
                            }
                        }
                        var vc = h / 2, hc = w / 2, r: any = w, b: any = h, l = 0, t = 0, wd2: any = w / 2, hd2: any = h / 2;
                        const ss = Math.min(w, h);
                        const cnstVal1 = 25000 * SLIDE_FACTOR;
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        const rdAngVal1 = (1 / 60000) * Math.PI / 180;
                        const rdAngVal2 = (21599999 / 60000) * Math.PI / 180;
                        const rdAngVal3 = 2 * Math.PI;
                        var a5, maxAdj1, a1, enAng, stAng, th, thh, th2, rw1, rh1, rw2, rh2, rw3, rh3, wtH, htH, dxH, dyH, xH, yH, rI,
                            u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12, u13, u14, u15, u16, u17, u18, u19, u20, u21, u22,
                            minAng, u23, a2, aAng, ptAng, wtA, htA, dxA, dyA, xA, yA, wtE, htE, dxE, dyE, xE, yE, wtD, htD, dxD, dyD,
                            xD, yD, dxG, dyG, xG, yG, dxB, dyB, xB, yB, sx1, sy1, sx2, sy2, rO, x1O, y1O, x2O, y2O, dxO, dyO, dO,
                            q1, q2, DO, q3, q4, q5, q6, q7, q8, sdelO, ndyO, sdyO, q9, q10, q11, dxF1, q12, dxF2, adyO, q13, q14, dyF1,
                            q15, dyF2, q16, q17, q18, q19, q20, q21, q22, dxF, dyF, sdxF, sdyF, xF, yF, x1I, y1I, x2I, y2I, dxI, dyI, dI,
                            v1, v2, DI, v3, v4, v5, v6, v7, v8, sdelI, v9, v10, v11, dxC1, v12, dxC2, adyI, v13, v14, dyC1, v15, dyC2, v16,
                            v17, v18, v19, v20, v21, v22, dxC, dyC, sdxC, sdyC, xC, yC, ist0, ist1, istAng0, isw1, isw2, iswAng0, istAng,
                            iswAng, p1, p2, p3, p4, p5, xGp, yGp, xBp, yBp, en0, en1, en2, sw0, sw1, swAng, stAng0;

                        a5 = (adj5 < 0) ? 0 : (adj5 > cnstVal1) ? cnstVal1 : adj5;
                        maxAdj1 = a5 * 2;
                        a1 = (adj1 < 0) ? 0 : (adj1 > maxAdj1) ? maxAdj1 : adj1;
                        enAng = (adj3 < rdAngVal1) ? rdAngVal1 : (adj3 > rdAngVal2) ? rdAngVal2 : adj3;
                        stAng = (adj4 < 0) ? 0 : (adj4 > rdAngVal2) ? rdAngVal2 : adj4;
                        th = ss * a1 / cnstVal2;
                        thh = ss * a5 / cnstVal2;
                        th2 = th / 2;
                        rw1 = wd2 + th2 - thh;
                        rh1 = hd2 + th2 - thh;
                        rw2 = rw1 - th;
                        rh2 = rh1 - th;
                        rw3 = rw2 + th2;
                        rh3 = rh2 + th2;
                        wtH = rw3 * Math.sin(enAng);
                        htH = rh3 * Math.cos(enAng);
                        dxH = rw3 * Math.cos(Math.atan2(wtH, htH));
                        dyH = rh3 * Math.sin(Math.atan2(wtH, htH));
                        xH = hc + dxH;
                        yH = vc + dyH;
                        rI = (rw2 < rh2) ? rw2 : rh2;
                        u1 = dxH * dxH;
                        u2 = dyH * dyH;
                        u3 = rI * rI;
                        u4 = u1 - u3;
                        u5 = u2 - u3;
                        u6 = u4 * u5 / u1;
                        u7 = u6 / u2;
                        u8 = 1 - u7;
                        u9 = Math.sqrt(u8);
                        u10 = u4 / dxH;
                        u11 = u10 / dyH;
                        u12 = (1 + u9) / u11;
                        u13 = Math.atan2(u12, 1);
                        u14 = u13 + rdAngVal3;
                        u15 = (u13 > 0) ? u13 : u14;
                        u16 = u15 - enAng;
                        u17 = u16 + rdAngVal3;
                        u18 = (u16 > 0) ? u16 : u17;
                        u19 = u18 - cd2;
                        u20 = u18 - rdAngVal3;
                        u21 = (u19 > 0) ? u20 : u18;
                        u22 = Math.abs(u21);
                        minAng = u22 * -1;
                        u23 = Math.abs(adj2);
                        a2 = u23 * -1;
                        aAng = (a2 < minAng) ? minAng : (a2 > 0) ? 0 : a2;
                        ptAng = enAng + aAng;
                        wtA = rw3 * Math.sin(ptAng);
                        htA = rh3 * Math.cos(ptAng);
                        dxA = rw3 * Math.cos(Math.atan2(wtA, htA));
                        dyA = rh3 * Math.sin(Math.atan2(wtA, htA));
                        xA = hc + dxA;
                        yA = vc + dyA;
                        wtE = rw1 * Math.sin(stAng);
                        htE = rh1 * Math.cos(stAng);
                        dxE = rw1 * Math.cos(Math.atan2(wtE, htE));
                        dyE = rh1 * Math.sin(Math.atan2(wtE, htE));
                        xE = hc + dxE;
                        yE = vc + dyE;
                        wtD = rw2 * Math.sin(stAng);
                        htD = rh2 * Math.cos(stAng);
                        dxD = rw2 * Math.cos(Math.atan2(wtD, htD));
                        dyD = rh2 * Math.sin(Math.atan2(wtD, htD));
                        xD = hc + dxD;
                        yD = vc + dyD;
                        dxG = thh * Math.cos(ptAng);
                        dyG = thh * Math.sin(ptAng);
                        xG = xH + dxG;
                        yG = yH + dyG;
                        dxB = thh * Math.cos(ptAng);
                        dyB = thh * Math.sin(ptAng);
                        xB = xH - dxB;
                        yB = yH - dyB;
                        sx1 = xB - hc;
                        sy1 = yB - vc;
                        sx2 = xG - hc;
                        sy2 = yG - vc;
                        rO = (rw1 < rh1) ? rw1 : rh1;
                        x1O = sx1 * rO / rw1;
                        y1O = sy1 * rO / rh1;
                        x2O = sx2 * rO / rw1;
                        y2O = sy2 * rO / rh1;
                        dxO = x2O - x1O;
                        dyO = y2O - y1O;
                        dO = Math.sqrt(dxO * dxO + dyO * dyO);
                        q1 = x1O * y2O;
                        q2 = x2O * y1O;
                        DO = q1 - q2;
                        q3 = rO * rO;
                        q4 = dO * dO;
                        q5 = q3 * q4;
                        q6 = DO * DO;
                        q7 = q5 - q6;
                        q8 = (q7 > 0) ? q7 : 0;
                        sdelO = Math.sqrt(q8);
                        ndyO = dyO * -1;
                        sdyO = (ndyO > 0) ? -1 : 1;
                        q9 = sdyO * dxO;
                        q10 = q9 * sdelO;
                        q11 = DO * dyO;
                        dxF1 = (q11 + q10) / q4;
                        q12 = q11 - q10;
                        dxF2 = q12 / q4;
                        adyO = Math.abs(dyO);
                        q13 = adyO * sdelO;
                        q14 = DO * dxO / -1;
                        dyF1 = (q14 + q13) / q4;
                        q15 = q14 - q13;
                        dyF2 = q15 / q4;
                        q16 = x2O - dxF1;
                        q17 = x2O - dxF2;
                        q18 = y2O - dyF1;
                        q19 = y2O - dyF2;
                        q20 = Math.sqrt(q16 * q16 + q18 * q18);
                        q21 = Math.sqrt(q17 * q17 + q19 * q19);
                        q22 = q21 - q20;
                        dxF = (q22 > 0) ? dxF1 : dxF2;
                        dyF = (q22 > 0) ? dyF1 : dyF2;
                        sdxF = dxF * rw1 / rO;
                        sdyF = dyF * rh1 / rO;
                        xF = hc + sdxF;
                        yF = vc + sdyF;
                        x1I = sx1 * rI / rw2;
                        y1I = sy1 * rI / rh2;
                        x2I = sx2 * rI / rw2;
                        y2I = sy2 * rI / rh2;
                        dxI = x2I - x1I;
                        dyI = y2I - y1I;
                        dI = Math.sqrt(dxI * dxI + dyI * dyI);
                        v1 = x1I * y2I;
                        v2 = x2I * y1I;
                        DI = v1 - v2;
                        v3 = rI * rI;
                        v4 = dI * dI;
                        v5 = v3 * v4;
                        v6 = DI * DI;
                        v7 = v5 - v6;
                        v8 = (v7 > 0) ? v7 : 0;
                        sdelI = Math.sqrt(v8);
                        v9 = sdyO * dxI;
                        v10 = v9 * sdelI;
                        v11 = DI * dyI;
                        dxC1 = (v11 + v10) / v4;
                        v12 = v11 - v10;
                        dxC2 = v12 / v4;
                        adyI = Math.abs(dyI);
                        v13 = adyI * sdelI;
                        v14 = DI * dxI / -1;
                        dyC1 = (v14 + v13) / v4;
                        v15 = v14 - v13;
                        dyC2 = v15 / v4;
                        v16 = x1I - dxC1;
                        v17 = x1I - dxC2;
                        v18 = y1I - dyC1;
                        v19 = y1I - dyC2;
                        v20 = Math.sqrt(v16 * v16 + v18 * v18);
                        v21 = Math.sqrt(v17 * v17 + v19 * v19);
                        v22 = v21 - v20;
                        dxC = (v22 > 0) ? dxC1 : dxC2;
                        dyC = (v22 > 0) ? dyC1 : dyC2;
                        sdxC = dxC * rw2 / rI;
                        sdyC = dyC * rh2 / rI;
                        xC = hc + sdxC;
                        yC = vc + sdyC;
                        ist0 = Math.atan2(sdyC, sdxC);
                        ist1 = ist0 + rdAngVal3;
                        istAng0 = (ist0 > 0) ? ist0 : ist1;
                        isw1 = stAng - istAng0;
                        isw2 = isw1 + rdAngVal3;
                        iswAng0 = (isw1 > 0) ? isw1 : isw2;
                        istAng = istAng0 + iswAng0;
                        iswAng = -iswAng0;
                        p1 = xF - xC;
                        p2 = yF - yC;
                        p3 = Math.sqrt(p1 * p1 + p2 * p2);
                        p4 = p3 / 2;
                        p5 = p4 - thh;
                        xGp = (p5 > 0) ? xF : xG;
                        yGp = (p5 > 0) ? yF : yG;
                        xBp = (p5 > 0) ? xC : xB;
                        yBp = (p5 > 0) ? yC : yB;
                        en0 = Math.atan2(sdyF, sdxF);
                        en1 = en0 + rdAngVal3;
                        en2 = (en0 > 0) ? en0 : en1;
                        sw0 = en2 - stAng;
                        sw1 = sw0 - rdAngVal3;
                        swAng = (sw0 > 0) ? sw1 : sw0;
                        stAng0 = stAng + swAng;

                        const strtAng = stAng0 * 180 / Math.PI;
                        const endAng = stAng * 180 / Math.PI;
                        const stiAng = istAng * 180 / Math.PI;
                        const swiAng = iswAng * 180 / Math.PI;
                        const ediAng = stiAng + swiAng;

                        let d_val = `M${xE},${yE} L${xD},${yD}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw2, rh2, stiAng, ediAng, false).replace("M", "L")} L${xBp},${yBp} L${xA},${yA} L${xGp},${yGp} L${xF},${yF}${PPTXShapeUtils.shapeArc(w / 2, h / 2, rw1, rh1, strtAng, endAng, false).replace("M", "L")} z`;
                        result += `<path d='${d_val}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;

                        break;
                    }
                    case "funnel": {
                        /**
                         * funnel: 漏斗形
                         * 
                         * 形状说明：
                         * - 上宽下窄的漏斗形状
                         * - 常用于数据分析和流程图
                         */
                        const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd", "attrs", "fmla"]);
                        let adj = 40000 * SLIDE_FACTOR;
                        if (shapAdjst !== undefined) {
                            adj = parseInt(shapAdjst.substr(4)) * SLIDE_FACTOR;
                        }
                        const cnstVal2 = 100000 * SLIDE_FACTOR;
                        let a = (adj < 0) ? 0 : (adj > cnstVal2) ? cnstVal2 : adj;
                        
                        // 漏斗底部宽度
                        const bottomW = w * a / cnstVal2;
                        
                        var d: any = `M0,0 L${w},0 L${((w + bottomW) / 2)},${h} L${((w - bottomW) / 2)},${h} z`;
                        
                        result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "leftRightCircularArrow": {
                        /**
                         * leftRightCircularArrow: 双向圆形箭头
                         * 
                         * 形状说明：
                         * - 圆形路径，两端有向左和向右的箭头
                         */
                        var wd2: any = w / 2;
                        let hd2: any = h / 2;
                        let r: any = Math.min(wd2, hd2);
                        
                        var d: any = `M${(wd2 - r)},${hd2}${PPTXShapeUtils.shapeArc(wd2, hd2, r, r, 180, 360, false).replace("M", "L")} M${(wd2 - r - r * 0.3)},${(hd2 - r * 0.2)} L${(wd2 - r)},${hd2} L${(wd2 - r - r * 0.3)},${(hd2 + r * 0.2)} M${(wd2 + r + r * 0.3)},${(hd2 - r * 0.2)} L${(wd2 + r)},${hd2} L${(wd2 + r + r * 0.3)},${(hd2 + r * 0.2)}`;
                        
                        result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "flowChartOfflineStorage": {
                        /**
                         * flowChartOfflineStorage: 流程图 - 离线存储
                         * 
                         * 形状说明：
                         * - 底部有三个向下的尖角（代表存储）
                         */
                        var d: any = `M0,0 L${w},0 L${w},${(h * 0.7)} L${(w * 0.66)},${h} L${(w * 0.34)},${h} L0,${(h * 0.7)} z`;
                        
                        result += `<path d='${d}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
                        break;
                    }
                    case "chartPlus":
                    case "chartStar":
                    case "chartX":
                    case "cornerTabs":
                    case "folderCorner":
                    case "lineInv":
                    case "nonIsoscelesTrapezoid":
                    case "plaqueTabs":
                    case "squareTabs":
                    case "upDownArrowCallout": {
                        // 其他占位符形状暂未实现
                        break;
                    }
                    case undefined:
                    default:
                }

                result += "</svg>";

                // 生成 data- 属性
                const dataAttrs1 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData = extractAnimationData(node, warpObj);
                let animationAttrs = "";
                if (animationData) {
                    animationAttrs = ` data-animation='${JSON.stringify(animationData)}'`;
                }

                // 简单形状（rect/roundRect 等）在 SVG 分支里没有生成 <path> 等图形元素，
                // 仅靠 SVG 无法显示填充/描边；此时把填充补到文本容器 div 上，
                // 与编辑器（div + background）保持一致。已用 SVG 路径绘制填充的形状不受影响。
                const svgHasShape = /<(path|rect|circle|ellipse|polygon|polyline|line)\b/.test(result);
                const divFill = svgHasShape ? "" : (await PPTXStyleUtils.getShapeFill(node, pNode, false, warpObj, source));

                result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)}${divFill}${transform3dStyle} z-index: ${order};'${dataAttrs1}${animationAttrs}>`;

                // TextBody
                if (node!["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    result += await PPTXTextUtils.genTextBody(node!["p:txBody"], node!, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj); //type='shape'
                }
                result += "</div>";
            } else if (custShapType !== undefined) {
                // 对于 group-abs 情况，使用缩放后的尺寸进行形状计算
                // 否则使用原始尺寸
                const renderW = (sType === 'group-abs') ? w : drawW;
                const renderH = (sType === 'group-abs') ? h : drawH;
                result += renderCustomShape(custShapType, renderW, renderH, imgFillFlg, grndFillFlg, fillColor, border, shpId, shapeArc);
                //console.log(result);

                result += "</svg>";

                // 生成 data- 属性
                const dataAttrs2 = genShapeDataAttributes(node, workingXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData2 = extractAnimationData(node, warpObj);
                let animationAttrs2 = "";
                if (animationData2) {
                    animationAttrs2 = ` data-animation='${JSON.stringify(animationData2)}'`;
                }

                result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(workingXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(workingXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)} z-index: ${order};'${dataAttrs2}${animationAttrs2}>`;

                // TextBody
                if (node!["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    if (type != "diagram" && type != "textBox") {
                        type = "shape";
                    }
                    // 对于 group-abs 情况，使用 workingXfrmNode 替换原始 xfrmNode
                    // 这样 genTextBody 就能获取到正确的缩放后尺寸
                    let textNode = node;
                    if (sType === 'group-abs' && workingXfrmNode !== slideXfrmNode) {
                        textNode = JSON.parse(JSON.stringify(node));
                        if (textNode!["p:spPr"] && textNode!["p:spPr"]["a:xfrm"]) {
                            textNode!["p:spPr"]["a:xfrm"] = workingXfrmNode;
                        }
                    }
                    result += await PPTXTextUtils.genTextBody(textNode!["p:txBody"], textNode!, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj); //type=shape
                }
                result += "</div>";

                // result = "";
            } else {

                // 生成 data- 属性
                const dataAttrs3 = genShapeDataAttributes(node, slideXfrmNode, id, name, idx, type, rotate, sType);
                
                // 检测动画信息并添加到data属性
                const animationData3 = extractAnimationData(node, warpObj);
                let animationAttrs3 = "";
                if (animationData3) {
                    animationAttrs3 = ` data-animation='${JSON.stringify(animationData3)}'`;
                }

                result += `<div class='block ${PPTXStyleUtils.getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)} ${PPTXStyleUtils.getContentDir(node, type, warpObj)}' _id='${id}' _idx='${idx}' _type='${type}' _name='${name}' style='${PPTXXmlUtils.getPosition(slideXfrmNode, pNode, slideLayoutXfrmNode, slideMasterXfrmNode, sType)}${PPTXXmlUtils.getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)}${PPTXStyleUtils.getBorder(node, pNode, false, "shape", warpObj)}${await PPTXStyleUtils.getShapeFill(node, pNode, false, warpObj, source)} z-index: ${order};'${dataAttrs3}>`;

                // TextBody
                if (node!["p:txBody"] !== undefined && (isUserDrawnBg === undefined || isUserDrawnBg === true)) {
                    result += await PPTXTextUtils.genTextBody(node!["p:txBody"], node!, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj);
                }
                result += "</div>";

            }
            return result;
        }

    return {
        // 重新导出路径生成函数
        shapeArc: shapeArc,
        shapeArcAlt: shapeArcAlt,
        shapePie: shapePie,
        shapeGear: shapeGear,
        shapeSnipRoundRect: shapeSnipRoundRect,
        shapeSnipRoundRectAlt: shapeSnipRoundRectAlt,
        polarToCartesian: polarToCartesian,
        // 核心形状生成函数
        genShape,
    };

/**
 * 提取形状的动画数据
 * @param {Object} node - 形状节点
 * @param {Object} warpObj - 包装对象
 * @returns {Object|null} 动画数据或null
 */
function extractAnimationData(node: any, warpObj: any) {
    // 检查形状是否有动画引用
    const nvSpPr = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr"]);
    if (!nvSpPr) return null;
    
    const nvPr = PPTXXmlUtils.getTextByPathList(nvSpPr, ["p:nvPr"]);
    if (!nvPr) return null;
    
    // 检查是否有动画效果
    const animLst = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animLst"]);
    if (animLst) {
        // 解析动画列表
        return parseAnimationList(animLst);
    }
    
    // 检查是否有动画引用
    const animRef = PPTXXmlUtils.getTextByPathList(nvPr, ["p:animRef"]);
    if (animRef && animRef.attrs) {
        const rId = animRef.attrs["r:embed"];
        if (rId && warpObj.slideAnims && warpObj.slideAnims[rId]) {
            return warpObj.slideAnims[rId];
        }
    }
    
    return null;
}

/**
 * 解析动画列表
 * @param {Object} animLst - 动画列表
 * @returns {Object} 解析后的动画数据
 */
function parseAnimationList(animLst: any) {
    // 处理单个动画或动画数组
    const animArray = Array.isArray(animLst["p:par"]) ? animLst["p:par"] : 
                     (animLst["p:par"] ? [animLst["p:par"]] : []);
    
    if (animArray.length === 0) return null;
    
    const par = animArray[0];
    if (par["p:cTn"]) {
        const cTn = par["p:cTn"];
        const animType = getAnimationType(cTn);
        const duration = cTn.attrs["dur"] || "1000";
        const delay = cTn.attrs["st"] || "0";
        
        return {
            type: animType,
            duration: parseInt(duration),
            delay: parseInt(delay)
        };
    }
    
    return null;
}

/**
 * 获取动画类型
 * @param {Object} cTn - 动画时间节点
 * @returns {string} 动画类型
 */
function getAnimationType(cTn: any) {
    // 检查子动画类型
    if (cTn["p:childTnLst"]) {
        const childTnLst = cTn["p:childTnLst"];
        
        // 检查是否有set动画（属性设置）
        if (childTnLst["p:set"]) {
            const set = childTnLst["p:set"];
            if (set["p:to"]) {
                const to = set["p:to"];
                if (to["p:strVal"] && to["p:strVal"].attrs["val"] === "visible") {
                    return "fade-in";
                }
            }
        }
        
        // 检查是否有motion动画（移动）
        if (childTnLst["p:cmd"]) {
            return "custom";
        }
    }
    
    // 默认动画类型
    return "appear";
}

/**
 * 处理3D效果
 * @param {Object} scene3d - 场景3D数据
 * @param {Object} sp3d - 形状3D数据
 * @returns {string} CSS 3D变换样式
 */
function process3DEffects(scene3d: any, sp3d: any) {
    let transform = "";
    
    // 处理相机设置
    if (scene3d && scene3d["a:camera"]) {
        const camera = scene3d["a:camera"];
        const prst = camera.attrs?.["prst"];
        
        // 根据预设相机类型应用不同的视角
        switch (prst) {
            case "orthographicFront":
                // 正交前视图 - 无3D效果
                break;
            case "orthographicTop":
                transform += " rotateX(-90deg)";
                break;
            case "orthographicBottom":
                transform += " rotateX(90deg)";
                break;
            case "orthographicLeft":
                transform += " rotateY(90deg)";
                break;
            case "orthographicRight":
                transform += " rotateY(-90deg)";
                break;
            case "perspectiveFront":
                // 透视前视图 - 轻微的3D效果
                transform += " perspective(1000px)";
                break;
            case "perspectiveTop":
                transform += " perspective(1000px) rotateX(-60deg)";
                break;
            case "perspectiveBottom":
                transform += " perspective(1000px) rotateX(60deg)";
                break;
            case "perspectiveLeft":
                transform += " perspective(1000px) rotateY(60deg)";
                break;
            case "perspectiveRight":
                transform += " perspective(1000px) rotateY(-60deg)";
                break;
            default:
                // 默认轻微透视
                transform += " perspective(800px)";
        }
    }
    
    // 处理形状3D效果（挤出、斜角等）
    if (sp3d) {
        // 挤出深度
        if (sp3d["a:extrusionH"]) {
            const extrusionH = parseInt(sp3d["a:extrusionH"].attrs?.["val"] || "0");
            if (extrusionH > 0) {
                // 转换为像素值
                const depth = Math.round(extrusionH * SLIDE_FACTOR);
                if (depth > 0) {
                    transform += ` translateZ(${depth}px)`;
                }
            }
        }
        
        // 顶部斜角
        if (sp3d["a:bevelT"]) {
            const bevelT = sp3d["a:bevelT"];
            const w = parseInt(bevelT.attrs?.["w"] || "0");
            const h = parseInt(bevelT.attrs?.["h"] || "0");
            if (w > 0 || h > 0) {
                // 应用轻微的3D旋转来模拟斜角效果
                transform += " rotateX(5deg) rotateY(5deg)";
            }
        }
        
        // 底部斜角
        if (sp3d["a:bevelB"]) {
            const bevelB = sp3d["a:bevelB"];
            const w = parseInt(bevelB.attrs?.["w"] || "0");
            const h = parseInt(bevelB.attrs?.["h"] || "0");
            if (w > 0 || h > 0) {
                // 应用轻微的3D旋转来模拟底部斜角
                transform += " rotateX(-3deg) rotateY(-3deg)";
            }
        }
    }
    
    if (transform) {
        return ` transform:${transform}; transform-style: preserve-3d;`;
    }
    
    return "";
}
})();