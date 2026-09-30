/**
 * 文本处理模块
 * 
 * 处理 PPTX 中的文本内容，包括：
 * - 文本样式解析（字体、大小、颜色、对齐等）
 * - 段落和文本运行处理
 * - 项目符号和编号
 * - 超链接处理
 * - 文本宽度计算
 * - RTL（从右到左）语言支持
 * 
 * @module utils/text
 */
import type { XmlNode, WarpObject } from '../core/types';

import { PPTXXmlUtils } from './xml';
import { PPTXStyleUtils } from './style';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR, RTL_LANGS_ARRAY, DINGBAT_UNICODE } from '../core/constants';
import { genChart } from './chart';
import TinyColor from 'tinycolor2';

// 创建 tinycolor 工厂函数以保持向后兼容
const tinycolor = (color: string, opts?: object) => new TinyColor(color, opts);
let is_first_br = false;



function getTextWidth(html: string) {
        let div = document.createElement('div');
        div.style.position = 'absolute';
        div.style.float = 'left';
        div.style.whiteSpace = 'nowrap';
        div.style.visibility = 'hidden';
        div.innerHTML = html;
        document.body.appendChild(div);
        let width = div.offsetWidth;
        document.body.removeChild(div);
        return width;
    }

    async function genTextBody(textBodyNode: XmlNode | undefined, spNode: XmlNode, slideLayoutSpNode: XmlNode | undefined, slideMasterSpNode: XmlNode | undefined, type: string, idx: number | undefined, warpObj: WarpObject, tbl_col_width?: number) {
            let text = "";
            let slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            if (textBodyNode === undefined) {
                return text;
            }
            //rtl : <p:txBody>
            //          <a:bodyPr wrap="square" rtlCol="1">

            // 获取anchor属性（垂直对齐方式）
            let anchor = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    if (anchor === undefined) {
                        anchor = "t";
                    }
                }
            }

            // 对于圆形/椭圆类形状，强制使用居中对齐以确保文本在形状内正确居中显示
            if (type === "shape") {
                let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
                const circularShapes = [
                    "ellipse", "ovalCallout", "wedgeEllipseCallout",
                    "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
                ];
                if (circularShapes.includes(shapeType) && anchor === "t") {
                    anchor = "ctr";
                }
            }

            // 获取bodyPr的内边距设置
            let bodyPrPadding = getBodyPrPadding(textBodyNode, type, anchor);
            text += bodyPrPadding;

            let pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
            let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "wrap"]);
            let spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["a:spAutoFit"]);
            let rtlColAttr = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "rtlCol"]);
            // 只有在没有 wrap 属性时，rtlCol="1" 才表示每个单词单独一行
            // 如果有 wrap="square" 等属性，则表示正常自动换行
            let isRTLCol = (rtlColAttr === "1" && wrapAttr === undefined);
            let isNoWrap = (wrapAttr === "none");
            let isAutoFit = (spAutoFitNode !== undefined);

            
            let apNode = textBodyNode["a:p"];
            if (apNode.constructor !== Array) {
                apNode = [apNode];
            }

            for (const i of apNode.keys()){
                let pNode = apNode[i];
                let rNode = pNode["a:r"];
                let fldNode = pNode["a:fld"];
                let brNode = pNode["a:br"];
                if (rNode !== undefined) {
                    rNode = (rNode.constructor === Array) ? rNode : [rNode];
                }
                if (rNode !== undefined && fldNode !== undefined) {
                    fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                    rNode = rNode.concat(fldNode)
                }
                if (rNode !== undefined && brNode !== undefined) {
                    is_first_br = true;
                    brNode = (brNode.constructor === Array) ? brNode : [brNode];
                    brNode.forEach((item: XmlNode, indx: number) => {
                        item.type = "br";
                    });
                    if (brNode.length > 1) {
                        brNode.shift();
                    }
                    rNode = rNode.concat(brNode)
                    rNode.sort((a: XmlNode, b: XmlNode) => {
                        return (a.attrs!.order ?? 0) - (b.attrs!.order ?? 0);
                    });
                }
                //rtlStr = "";//`dir='${isRTL}'`;
                let styleText = "";
                let marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj, apNode.length, i, anchor);
                if (marginsVer != "") {
                    styleText = marginsVer;
                }
                // 移除 font-size: 0px 设置，避免影响文本显示
                // if (type == "body" || type == "obj" || type == "shape") {
                //     styleText += "font-size: 0px;";
                //     //styleText += "line-height: 0;";
                //     styleText += "font-weight: 100;";
                //     styleText += "font-style: normal;";
                // }
                let cssName = "";

                if (styleText in warpObj.styleTable) {
                    cssName = warpObj.styleTable[styleText]["name"];
                } else {
                    cssName = `_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
                    warpObj.styleTable[styleText] = {
                        "name": cssName,
                        "text": styleText
                    };
                }

                let prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                // 占位符可能没有自己的xfrm（<p:spPr/>为空），尺寸继承自版式/母版中的同名占位符，需回退获取
                if (prg_width_node === undefined || prg_width_node === null) {
                    prg_width_node = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                }
                if (prg_width_node === undefined || prg_width_node === null) {
                    prg_width_node = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                }
                let prg_height_node;// = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
                
                // 获取bodyPr的内边距属性，用于计算可用宽度
                let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
                let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
                            
                // 计算内边距像素值
                let lInsPx, rInsPx;
                if (type === "table") {
                    // 对于表格，如果没有明确设置内边距，则使用0
                    lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR) : 0;
                    rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR) : 0;
                } else {
                    lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
                    rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR) : (type === "diagram" ? 0.04 * 96 : 0.1 * 96);
                }
                            
                // 如果明确设置了lIns="0"或rIns="0"，则不应用默认值
                if (lIns === "0") lInsPx = 0;
                if (rIns === "0") rInsPx = 0;
                            
                // 检查是否为圆形/椭圆类形状，如果是则应用额外的安全边距
                let shapeType = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
                const circularShapes = [
                    "ellipse", "ovalCallout", "wedgeEllipseCallout",
                    "pie", "pieWedge", "chord", "sector", "arc", "blockArc"
                ];
                const isCircularShape = circularShapes.includes(shapeType);
                
                // 处理组合缩放 - 如果形状在group-abs组合中,需要应用缩放到段落宽度
                let sld_prg_width_val = null;
                if (prg_width_node !== undefined && prg_width_node !== null) {
                    let parsedWidth = parseInt(prg_width_node);
                    if (!isNaN(parsedWidth) && parsedWidth > 0) {
                        sld_prg_width_val = Math.round(parsedWidth * SLIDE_FACTOR * 100) / 100;
                    }
                }
                if (sld_prg_width_val !== null && warpObj.currentGroupScale) {
                    const { scaleX, scaleY } = warpObj.currentGroupScale;
                    sld_prg_width_val = Math.round(sld_prg_width_val * scaleX * 100) / 100;
                }
                
                let sld_prg_width = "";
                if (sld_prg_width_val !== null && !isNoWrap) {
                    // 减去内边距宽度，得到实际可用宽度
                    let availableWidth = sld_prg_width_val - lInsPx - rInsPx;
                                
                    // 对于圆形/椭圆类形状，应用额外的安全边距（减少5%）以确保正确换行
                    if (isCircularShape) {
                        availableWidth = availableWidth * 0.95;
                    }
                                
                    sld_prg_width = `width:${Math.max(0, Math.round(availableWidth * 100) / 100)}px;`;
                } else if (sld_prg_width_val === null) {
                    sld_prg_width = "width:inherit;";
                }
                let sld_prg_height = ""; // 移除高度设置，避免段落叠加
                let prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
                let isRTL = (prg_dir == "pregraph-rtl");
                let directionStyle = isRTL ? "direction: rtl;" : "direction: ltr;";
                let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj, spNode);
                // 在外层div上也设置justify-content，确保对齐正确
                // 注意：表格单元格的对齐由td元素的text-align控制，这里不设置justify-content
                let outerFlexStyle = "";
                if (type !== "table") {
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        // 如果是 RTL，则 flex-start 是右对齐；如果是 LTR，则 flex-end 是右对齐
                        outerFlexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
                    } else if (horizontalAlign === "h-mid") {
                        outerFlexStyle = "justify-content: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        // RTL 模式下，左对齐使用 flex-end
                        outerFlexStyle = "justify-content: flex-end;";
                    } else {
                        outerFlexStyle = "justify-content: flex-start;";
                    }
                }
                text += `<div style='display: flex;${sld_prg_width}${sld_prg_height}${outerFlexStyle}${directionStyle}' class='slide-prgrph ${horizontalAlign}` + ` ${prg_dir} ` + cssName + "' >";
                let buText_ary = await genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj, apNode.length, anchor);
                let isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
                let bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? (Number(buText_ary[1]) + Number(buText_ary[2])) : 0;

                // 在 RTL 模式下，项目符号在右边，所以先添加文本，再添加项目符号
                if (isRTL && isBullate) {
                    // 暂时不添加项目符号，等文本添加后再添加
                } else {
                    text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
                }
                //get text margin 
                // 获取段落的字体大小，用于计算项目符号边距
                let fontSize = undefined;
                if (rNode !== undefined && rNode.length > 0) {
                    // 使用第一个文本运行的字体大小作为参考
                    fontSize = PPTXStyleUtils.getFontSize(rNode[0], textBodyNode, pFontStyle, 1, type, warpObj);
                    if (fontSize && fontSize.endsWith('px')) {
                        fontSize = parseFloat(fontSize);
                    }
                }
                let margin_ary = PPTXStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize);
                let margin = margin_ary[0];
                let mrgin_val = margin_ary[1];
                if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
                    //sorce : table text
                    prg_width_node = tbl_col_width;
                }

                let prgrph_text = "";
                //let prgr_txt_art = [];
                let total_text_len = 0;
                

                if (rNode === undefined && pNode !== undefined) {
                    // without r
                    let prgr_text = await genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
                    if (isBullate) {
                        total_text_len += getTextWidth(prgr_text);
                    }
                    prgrph_text += prgr_text;
                } else if (rNode !== undefined) {
                    // with multi r
                    let previousStyle: Record<string, unknown> = {};
                    for (const j of rNode.keys()){
                        // 如果当前元素没有sz属性，使用前面元素的样式
                        if (rNode[j]["a:rPr"] && !rNode[j]["a:rPr"]["attrs"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"] = { "sz": previousStyle["sz"] };
                        } else if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && !rNode[j]["a:rPr"]["attrs"]["sz"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"]["sz"] = previousStyle["sz"];
                        }
                        
                        let prgr_text = await genSpanElement(rNode[j], j, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
                        if (isBullate) {
                            total_text_len += getTextWidth(prgr_text);
                        }
                        

                        prgrph_text += prgr_text;
                        
                        // 保存当前元素的样式，供后面元素继承
                        if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && rNode[j]["a:rPr"]["attrs"]["sz"]) {
                            previousStyle["sz"] = rNode[j]["a:rPr"]["attrs"]["sz"];
                        }
                    }
                }

                prg_width_node = parseInt(prg_width_node) * SLIDE_FACTOR - bu_width - Number(mrgin_val);
                prg_width_node = Math.round(prg_width_node * 100) / 100;
                // 不根据文本测量宽度收缩段落容器：PPT 段落应以文本框可用宽度为准，
                // 否则在中文/混合字体场景下会因测量误差导致提前换行。
                // 如果没有明确设置wrap="none"或spAutoFit，默认不设置内层div的宽度，让文本自然流动
                let prg_width = "";
                let textContainerWidth = ""; // 默认不设置宽度，让文本容器自适应
                // 如果是 noAutofit（文本框宽度固定），需要设置文本容器宽度以限制换行
                // 但对于表格，不设置内层容器宽度，让其自适应
                // 只有拿到了有效的文本框宽度才设置内层容器宽度，避免 width:NaNpx/0px 导致逐字换行（视觉上变竖排）
                if (!isAutoFit && !isNoWrap && sld_prg_width_val !== null && !isNaN(sld_prg_width_val) && type !== "table") {
                    // 使用与外层段落相同的可用宽度
                    let availableWidthForTextContainer = sld_prg_width_val - lInsPx - rInsPx;
                    // 对于圆形/椭圆类形状，应用额外的安全边距
                    if (isCircularShape) {
                        availableWidthForTextContainer = availableWidthForTextContainer * 0.95;
                    }
                    textContainerWidth = `width:${Math.max(0, Math.round(availableWidthForTextContainer * 100) / 100)}px;`;
                }
                if (isRTL && isBullate) {
                    // RTL 模式下有项目符号时，文本容器不设宽度，让内容自适应
                    textContainerWidth = "";
                }
                if (prg_width_node !== undefined && prg_width_node !== null && !isNoWrap) {
                    // 只有明确不需要换行时才设置宽度
                    prg_width = `width:${(Math.round(prg_width_node * 100) / 100)}px;`;
                }
                // 对于圆形/椭圆类形状，使用normal white-space和break-word以确保正确换行
                let whiteSpaceStyle;
                if (isCircularShape) {
                    whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: normal; overflow-wrap: break-word;";
                } else {
                    whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: pre-wrap;";
                }


                let textAlignStyle = "";
                if (horizontalAlign === "h-mid") {
                    textAlignStyle = "text-align: center;";
                } else if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                    textAlignStyle = "text-align: right;";
                } else if (horizontalAlign === "h-left-rtl") {
                    textAlignStyle = "text-align: left;";
                } else {
                    textAlignStyle = "text-align: left;";
                }
                
                // 添加强制换行样式以确保中文文本正确换行
                textAlignStyle += " word-break: break-all;";
                // 为了确保右对齐生效，添加flex布局的justify-content属性
                // 注意：表格单元格的对齐由td元素的text-align控制，这里不设置justify-content
                let flexStyle = "";
                if (type !== "table") {
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        // 如果是 RTL，则 flex-start 是右对齐；如果是 LTR，则 flex-end 是右对齐
                        flexStyle = isRTL ? "justify-content: flex-start;" : "justify-content: flex-end;";
                    } else if (horizontalAlign === "h-mid") {
                        flexStyle = "justify-content: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        // RTL 模式下，左对齐使用 flex-end
                        flexStyle = "justify-content: flex-end;";
                    } else {
                        flexStyle = "justify-content: flex-start;";
                    }
                }
                text += `<div style='display: flex;${flexStyle}${textContainerWidth}${directionStyle}'>`;
                // 在 RTL 模式下，项目符号应该和文本在同一个容器中
                if (isRTL && isBullate && buText_ary[0] !== undefined) {
                    // 先添加项目符号，再添加文本（在 RTL 容器中，第一个子元素显示在最右边）
                    text += buText_ary[0];
                }
                text += `<div style='${styleText}${directionStyle}${whiteSpaceStyle}${margin}${textAlignStyle}'>`;
                text += prgrph_text;
                text += "</div>";
                text += "</div>";
                text += "</div>";
            }

            // 关闭bodyPr内边距div（如果存在）
            // 与 getBodyPrPadding 中的条件保持一致：type !== "table"
            if (type !== "table") {
                text += "</div>";
            }

            return text;
    }

    /**
     * 获取bodyPr的内边距设置
     * @param {Object} textBodyNode - 文本体节点
     * @param {string} type - 形状类型
     * @param {string} anchor - 垂直对齐方式（t=顶部, ctr=居中, b=底部）
     * @returns {string} CSS padding字符串
     */
    function getBodyPrPadding(textBodyNode: XmlNode, type: string, anchor: string) {
        let paddingStyle = "";

        // 获取bodyPr的各个内边距属性
        let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
        let tIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "tIns"]);
        let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
        let bIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "bIns"]);

        // 文本框、形状、diagram 或普通对象时都应用内边距
        // 注意：type 可能为 "textBox", "shape", "diagram", "obj" 或 undefined
        // 只要不是 table 类型，就应该应用内边距
        if (type !== "table") {
            // 根据PPTX规范，bodyPr的ins属性单位是EMU（English Metric Units）
            // 1 inch = 914400 EMU, 1 inch = 96px, 所以 1 EMU = 96/914400 px ≈ 0.000105 px
            // 如果没有设置内边距，使用默认值：
            // - diagram 类型使用更小的默认值，使其更接近原始 PPT 效果
            // - 其他类型：tIns 和 bIns 默认值为 0.05 inch = 45720 EMU ≈ 4.8px
            // - 其他类型：lIns 和 rIns 默认值为 0.1 inch = 91440 EMU ≈ 9.6px
            let defaultLIns, defaultTIns, defaultRIns, defaultBIns;
            if (type === "diagram") {
                // diagram 类型使用更小的默认内边距
                defaultLIns = 0.04 * 96;  // 0.04 inch ≈ 3.84px
                defaultTIns = 0.02 * 96;  // 0.02 inch ≈ 1.92px
                defaultRIns = 0.04 * 96;  // 0.04 inch ≈ 3.84px
                defaultBIns = 0.02 * 96;  // 0.02 inch ≈ 1.92px
            } else {
                defaultLIns = 0.1 * 96;   // 0.1 inch ≈ 9.6px
                defaultTIns = 0.05 * 96;  // 0.05 inch ≈ 4.8px
                defaultRIns = 0.1 * 96;   // 0.1 inch ≈ 9.6px
                defaultBIns = 0.05 * 96;  // 0.05 inch ≈ 4.8px
            }

            let lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR).toFixed(2) : defaultLIns.toFixed(2);
            let tInsPx = tIns ? (parseInt(tIns) * SLIDE_FACTOR).toFixed(2) : defaultTIns.toFixed(2);
            let rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR).toFixed(2) : defaultRIns.toFixed(2);
            let bInsPx = bIns ? (parseInt(bIns) * SLIDE_FACTOR).toFixed(2) : defaultBIns.toFixed(2);

            // 如果明确设置了lIns="0"或rIns="0"，则不应用默认值
            if (lIns === "0") lInsPx = "0";
            if (rIns === "0") rInsPx = "0";

            // 根据anchor决定padding div的高度设置
            // 如果是垂直居中（anchor="ctr"），不设置height: 100%，让内容自然撑开
            // 这样外层的v-mid类的justify-content: center才能生效
            let heightStyle = "";
            if (anchor !== "ctr") {
                heightStyle = "height: 100%;";
            }

            paddingStyle = `<div style="padding: ${tInsPx}px ${rInsPx}px ${bInsPx}px ${lInsPx}px; box-sizing: border-box; ${heightStyle}">`;
        }

        return paddingStyle;
    }
        
        async function genBuChar(node: XmlNode, i: number, spNode: XmlNode, textBodyNode: XmlNode, pFontStyle: Record<string, unknown>, idx: number | undefined, type: string, warpObj: WarpObject, totalParagraphs?: number, anchor?: string) {

            ///////////////////////////////////////Amir///////////////////////////////
            let sldMstrTxtStyles = warpObj["slideMasterTextStyles"];
            let lstStyle = textBodyNode["a:lstStyle"];

            let rNode = PPTXXmlUtils.getTextByPathList(node, ["a:r"]);
            if (rNode !== undefined && rNode.constructor === Array) {
                rNode = rNode[0]; //bullet only to first "a:r"
            }

            let lvl = parseInt (PPTXXmlUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
            if (isNaN(lvl)) {
                lvl = 1;
            }
            let lvlStr = `a:lvl${lvl}pPr`;
            let dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;

            if (rNode !== undefined) {
                dfltBultColor = await PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
                color_tye = dfltBultColor[2];
                dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);

            } else {
                return "";
            }


            let bullet = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
            /////////////////////////////////////////////////////////////////


            let pPrNode = node["a:pPr"];
            let BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }

            let buType = "TYPE_NONE";

            let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;

            let buChar = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
            let buNum = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
            let buPic = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }

            let buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    let prcnt = parseInt(buFontSize) / 100000;
                    //dfltBultSize = XXpt
                    //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                    bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;// + "pt";
                }
            } else {
                bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
            }

            //get definde bullet COLOR
            let buClrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buClr"]);


            if (buChar === undefined && buNum === undefined && buPic === undefined) {

                if (lstStyle !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                    if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                        pPrNode = lstStyle[lvlStr];
                    }
                }
            }
            if (buChar === undefined && buNum === undefined && buPic === undefined) {
                //check in slidelayout and masterlayout - TODO
                if (pPrNodeLaout !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                }
                if (buChar === undefined && buNum === undefined && buPic === undefined) {
                    //masterlayout

                    if (pPrNodeMaster !== undefined) {
                        BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                        if (BullNONE !== undefined) {
                            return "";
                        }
                        buType = "TYPE_NONE";
                        buChar = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                        buNum = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                        buPic = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                        if (buChar !== undefined) {
                            buType = "TYPE_BULLET";
                        }
                        if (buNum !== undefined) {
                            buType = "TYPE_NUMERIC";
                        }
                        if (buPic !== undefined) {
                            buType = "TYPE_BULPIC";
                        }
                    }

                }

            }
            //rtl
            let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            let isRTL = false;
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
            }
            //align
            let alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
            if (alignNode === undefined) {
                alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                if (alignNode === undefined) {
                    alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }
            //indent?
            let indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
                if (indentNode === undefined) {
                    indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
                }
            }
            let indent = 0;
            if (indentNode !== undefined) {
                indent = parseInt(indentNode) * SLIDE_FACTOR;
            }
            //marL
            let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }

            if (marLNode !== undefined) {
                let marginLeft = parseInt(marLNode) * SLIDE_FACTOR;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
                marLStr += `${margin_val}px;`;
            }
            
            //marR?
            let marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
            if (marRNode === undefined && marLNode === undefined) {
                //need to check if this posble - TODO
                marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
                if (marRNode === undefined) {
                    marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
                }
            }
            if (marRNode !== undefined) {
                let marginRight = parseInt(marRNode) * SLIDE_FACTOR;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                marRStr += `${((marginRight + indent < 0) ? 0 : (marginRight + indent))}px;`;
            }

            if (buType != "TYPE_NONE") {
                //let buFontAttrs = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
            }

            //get definde bullet COLOR
            if (buClrNode === undefined){
                //lstStyle
                buClrNode = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
            }
            if (buClrNode === undefined) {
                buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buClr"]);
                if (buClrNode === undefined) {
                    buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buClr"]);
                }
            }
            let defBultColor;
            if (buClrNode !== undefined) {
                defBultColor = PPTXStyleUtils.getSolidFill(buClrNode, undefined, undefined, warpObj);
            }
            if (defBultColor === undefined || defBultColor == "NONE") {
                bultColor = dfltBultColor;
            } else {
                bultColor = [defBultColor, "", "solid"];
                color_tye = "solid";
            }


            //get definde bullet SIZE
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        let prcnt = parseInt(buFontSize) / 100000;
                        //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                        bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;// + "pt";
                    }
                }else{
                    bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
                }
            }
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        let prcnt = parseInt(buFontSize) / 100000;
                        //dfltBultSize = XXpt
                        //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                        bultSize = `${prcnt * (parseInt(String(dfltBultSizeNoPt)))}px`;// + "pt";
                    }
                } else {
                    bultSize = `${(parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR}px`;
                }
            }
            if (buFontSize === undefined) {
                bultSize = dfltBultSize;
            }
            font_val = parseInt(bultSize ?? "", 10);

            // 项目符号要垂直居中于「段落首行」，而不是段落的顶边或整段：
            // 宿主页面的样式（如 .slide div.h-left{align-items:flex-start}）会把兄弟元素顶到行首，
            // 所以这里显式声明自身对齐方式，并套用与首行一致的行高与段前距。
            const bulletVAlignStyle = (() => {
                const fontSizePx = parseFloat(dfltBultSize ?? "") || 0;
                const vMargins = PPTXStyleUtils.getVerticalMargins(
                    node, textBodyNode, type, idx, warpObj,
                    totalParagraphs !== undefined ? totalParagraphs : 1, i, anchor);
                // 行高是比例值，需要乘正文字号换算成首行盒高（符号字号可能与正文不同）
                const lhMatch = /line-height:\s*([\d.]+)/.exec(vMargins);
                const mtMatch = /margin-top:\s*(-?[\d.]+)px/.exec(vMargins);
                let style = "align-self: flex-start;";
                if (mtMatch) {
                    style += `margin-top:${mtMatch[1]}px;`;
                }
                if (lhMatch && fontSizePx > 0) {
                    style += `height:${(parseFloat(lhMatch[1]) * fontSizePx).toFixed(2)}px;`;
                }
                return style;
            })();
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                let typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
                let typeface = "";
                let isWingdingsFont = false;
                if (typefaceNode !== undefined) {
                    isWingdingsFont = (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings");
                    typeface = `font-family: ${typefaceNode}`;
                }
                // let marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(marLNode)) * SLIDE_FACTOR;
                // let marginRight = parseInt (PPTXXmlUtils.getTextByPathList(marRNode)) * SLIDE_FACTOR;
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * SLIDE_FACTOR;
                // }
                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }

                bullet = `<div style='${typeface};` +
                    marLStr + marRStr +
                    `font-size:${bultSize};` ;
                
                //bullet += "display: table-cell;";
                //"line-height: 0px;";
                if (color_tye == "solid") {
                    if (bultColor[0] !== undefined && bultColor[0] != "") {
                        let bulletColorValue = bultColor[0];
                        if (bulletColorValue.length === 8) {
                            let colorObj = tinycolor(bulletColorValue);
                            bulletColorValue = colorObj.toRgbString();
                        } else {
                            bulletColorValue = `#${bulletColorValue}`;
                        }
                        bullet += `color:${bulletColorValue}; `;
                    }
                    if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                        bullet += `text-shadow:${bultColor[1]};`;
                    }
                    //no highlight/background-color to bullet
                    // if (bultColor[3] !== undefined && bultColor[3] != "") {
                    //     styleText += "background-color: #" + bultColor[3] + ";";
                    // }
                } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
                    if (color_tye == "pattern") {
                        bullet += `background:${bultColor[0][0]};`;
                        if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                            bullet += `background-size:${bultColor[0][1]};`;//" 2px 2px;" +
                        }
                        if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                            bullet += `background-position:${bultColor[0][2]};`;//" 2px 2px;" +
                        }
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                        //     "filter: " + bultColor[1].effcts + ";";
                    } else if (color_tye == "pic") {
                        bullet += `${bultColor[0]};`;
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";

                    } else if (color_tye == "gradient") {

                        let colorAry = bultColor[0].color;
                        // 归一化：非数组时（单个值/undefined）避免 .keys() 报错
                        if (!Array.isArray(colorAry)) {
                            colorAry = colorAry ? [colorAry] : [];
                        }
                        let rot = bultColor[0].rot;

                        bullet += `background: linear-gradient(${rot}deg,`;
                        for (const i of colorAry.keys()){
                            if (i == colorAry.length - 1) {
                                bullet += `#${colorAry[i]});`;
                            } else {
                                bullet += `#${colorAry[i]}, `;
                            }
                        }
                        // bullet += "color: transparent;" +
                        //     "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    bullet += `-webkit-background-clip: text;background-clip: text;color: transparent;`;
                    if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                        bullet += `-webkit-text-stroke: ${bultColor[1].border};`;
                    }
                    if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                        bullet += `filter: ${bultColor[1].effcts};`;
                    }
                }

                if (isRTL) {
                    //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
                    bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
                }
                let isIE11 = !!(window as { MSInputMethodContext?: unknown }).MSInputMethodContext && !!(document as { documentMode?: unknown }).documentMode;
                let htmlBu = buChar;
                let useUnicodeFont = false;

                // 只有在非 Wingdings 字体时才进行 Unicode 转换
                if (!isIE11 && !isWingdingsFont) {
                    //ie11 does not support unicode ?
                    htmlBu = getHtmlBullet(typefaceNode, buChar);
                    useUnicodeFont = (htmlBu !== buChar);
                }
                
                // 如果使用了 Unicode 转换且是 Wingdings 字体，则使用标准字体
                if (useUnicodeFont && isWingdingsFont && typefaceNode !== undefined) {
                    // 使用正则表达式替换所有可能的 Wingdings 字体变体
                    bullet = bullet.replace(/font-family:\s*(Wingdings|Wingdings\s*2|Wingdings\s*3|Webdings)\s*/gi, "font-family: Arial, sans-serif");
                }
                
                bullet += `${bulletVAlignStyle}display: flex; align-items: center;'><div>${htmlBu}</div></div>`;
                //} 
                // else {
                //     marginLeft = 328600 * SLIDE_FACTOR * lvl;

                //     bullet = `<div style='${marLStr}'>` + buChar + "</div>";
                // }
            } else if (buType == "TYPE_NUMERIC") {
                // 初始化项目符号计数器
                if (!warpObj.bulletCounter) {
                    warpObj.bulletCounter = {};
                }
                
                // 生成项目符号的唯一键
                const bulletKey = `${buNum}_${lvl}`;
                
                // 初始化或获取当前计数器
                if (!warpObj.bulletCounter[bulletKey]) {
                    warpObj.bulletCounter[bulletKey] = {
                        index: 0,
                        type: buNum,
                        level: lvl
                    };
                }
                
                // 增加计数器
                warpObj.bulletCounter[bulletKey].index++;
                
                // 生成数字编号
                const bulletIndex = warpObj.bulletCounter[bulletKey].index;
                const bulletText = getNumTypeNum(buNum, bulletIndex);

                bullet = `<div style='${marLStr}${marRStr}`;
                if (bultColor && bultColor[0] !== undefined && bultColor[0] != "") {
                    let bulletNumColorValue = bultColor[0];
                    if (bulletNumColorValue.length === 8) {
                        let colorObj = tinycolor(bulletNumColorValue);
                        bulletNumColorValue = colorObj.toRgbString();
                    } else {
                        bulletNumColorValue = `#${bulletNumColorValue}`;
                    }
                    bullet += `color:${bulletNumColorValue};`;
                }
                bullet += `font-size:${bultSize};`;
                if (isRTL) {
                    bullet += "white-space: nowrap ;direction:rtl;";
                } else {
                    bullet += "white-space: nowrap ;direction:ltr;";
                }
                bullet += `${bulletVAlignStyle}display: flex; align-items: center;'><div>${bulletText}</div></div>`;

            } else if (buType == "TYPE_BULPIC") { //PIC BULLET
                // let marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * SLIDE_FACTOR;
                // let marginRight = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * SLIDE_FACTOR;

                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }
                // 
                // //buPic
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * SLIDE_FACTOR;
                // } else {
                //     marginLeft = 0;
                // }
                //let buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
                let buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                let svgPicPath = "";
                let buImg;
                if (buPicId !== undefined) {
                    //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                    //buImg = warpObj["zip"].file(svgPicPath).asText();
                    //}else{
                    //buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                    let imgPath = (warpObj["slideResObj"][buPicId] !== undefined) ? warpObj["slideResObj"][buPicId]["target"] : undefined;

                    if (imgPath === undefined) {
                        buImg = "";
                    } else {
                        let imgFile = warpObj["zip"].file(imgPath);
                        if (imgFile === null) {
                            buImg = "";
                        } else {
                            let imgArrayBuffer = await imgFile.async("arraybuffer");
                            let imgExt = imgPath.split(".").pop();
                            let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                            buImg = `<img src='data:${imgMimeType};base64,` + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>"// height: 100%
        
                        }
                    }
                }
                if (buPicId === undefined) {
                    buImg = "&#8227;";
                }
                bullet = `<div style='${marLStr}${marRStr}` +
                    `width:${bultSize};${bulletVAlignStyle}display: flex; align-items: center;`;// +
                //"line-height: 0px;";
                if (isRTL) {
                    bullet += "white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
                }
                bullet += `'>${buImg}  </div>`;
                //////////////////////////////////////////////////////////////////////////////////////
            }
            // else {
            //     bullet = "<div style='margin-left: " + 328600 * SLIDE_FACTOR * lvl + "px" +
            //         `; margin-right: ${0}px;'></div>`;
            // }

            return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
        }
        function getHtmlBullet(typefaceNode: string, buChar: string) {
            //http://www.alanwood.net/demos/wingdings.html
            //not work for IE11
            //console.log("genBuChar typefaceNode:", typefaceNode, " buChar:", buChar, "charCodeAt:", buChar.charCodeAt(0))
            switch (buChar) {
                case "§":
                    return "&#9632;";//"■"; //9632 | U+25A0 | Black square
                    break;
                case "q":
                    return "&#10065;";//"❑"; // 10065 | U+2751 | Lower right shadowed white square
                    break;
                case "v":
                    return "&#10070;";//"❖"; //10070 | U+2756 | Black diamond minus white X
                    break;
                case "Ø":
                    return "&#11162;";//"⮚"; //11162 | U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
                    break;
                case "ü":
                    return "&#10004;";//"✔";  //10004 | U+2714 | Heavy check mark
                    break;
                case "o":
                    return "&#9679;";//"●"; //9679 | U+25CF | Black circle
                    break;
                case "O":
                    return "&#9675;";//"○"; //9675 | U+25CB | White circle
                    break;
                case "a":
                    return "&#9650;";//"▲"; //9650 | U+25B2 | Black up-pointing triangle
                    break;
                case "A":
                    return "&#9651;";//"△"; //9651 | U+25B3 | White up-pointing triangle
                    break;
                case "b":
                    return "&#9660;";//"▼"; //9660 | U+25BC | Black down-pointing triangle
                    break;
                case "B":
                    return "&#9661;";//"▽"; //9661 | U+25BD | White down-pointing triangle
                    break;
                case "c":
                    return "&#9654;";//"▶"; //9654 | U+25B6 | Black right-pointing triangle
                    break;
                case "C":
                    return "&#9655;";//"▷"; //9655 | U+25B7 | White right-pointing triangle
                    break;
                case "d":
                    return "&#9664;";//"◀"; //9664 | U+25C0 | Black left-pointing triangle
                    break;
                case "D":
                    return "&#9665;";//"◁"; //9665 | U+25C1 | White left-pointing triangle
                    break;
                case "e":
                    return "&#9670;";//"◆"; //9670 | U+25C6 | Black diamond
                    break;
                case "E":
                    return "&#9671;";//"◇"; //9671 | U+25C7 | White diamond
                    break;
                case "f":
                    return "&#10003;";//"✓"; //10003 | U+2713 | Check mark
                    break;
                case "F":
                    return "&#10007;";//"✗"; //10007 | U+2717 | Ballot X
                    break;
                case "g":
                    return "&#10002;";//"✔"; //10002 | U+2714 | Heavy check mark
                    break;
                case "G":
                    return "&#10008;";//"✘"; //10008 | U+2718 | Heavy ballot X
                    break;
                case "h":
                    return "&#9899;";//"★"; //9899 | U+2605 | Black star
                    break;
                case "H":
                    return "&#9734;";//"☆"; //9734 | U+2606 | White star
                    break;
                case "i":
                    return "&#10052;";//"✤"; //10052 | U+2724 | Heavy four-pointed star
                    break;
                case "I":
                    return "&#10053;";//"✥"; //10053 | U+2725 | Four-pointed star
                    break;
                case "j":
                    return "&#10022;";//"✶"; //10022 | U+2736 | Six-pointed star
                    break;
                case "J":
                    return "&#10023;";//"✷"; //10023 | U+2737 | Eight-pointed star
                    break;
                case "k":
                    return "&#10016;";//"✈"; //10016 | U+2708 | Airplane
                    break;
                case "K":
                    return "&#10024;";//"✈"; //10024 | U+2708 | Airplane
                    break;
                case "l":
                    return "&#10038;";//"✦"; //10038 | U+2726 | Black four-pointed star
                    break;
                case "L":
                    return "&#10039;";//"✧"; //10039 | U+2727 | White four-pointed star
                    break;
                case "m":
                    return "&#10017;";//"✉"; //10017 | U+2709 | Envelope
                    break;
                case "M":
                    return "&#9993;";//"✉"; //9993 | U+2709 | Envelope
                    break;
                case "n":
                    return "&#10084;";//"❤"; //10084 | U+2764 | Heavy black heart
                    break;
                case "N":
                    return "&#9829;";//"♥"; //9829 | U+2665 | Black heart suit
                    break;
                case "p":
                    return "&#9830;";//"♦"; //9830 | U+2666 | Black diamond suit
                    break;
                case "P":
                    return "&#9826;";//"♢"; //9826 | U+2662 | White diamond suit
                    break;
                case "r":
                    return "&#9827;";//"♣"; //9827 | U+2663 | Black club suit
                    break;
                case "R":
                    return "&#9827;";//"♣"; //9827 | U+2663 | Black club suit
                    break;
                case "s":
                    return "&#9824;";//"♠"; //9824 | U+2660 | Black spade suit
                    break;
                case "S":
                    return "&#9824;";//"♠"; //9824 | U+2660 | Black spade suit
                    break;
                case "t":
                    return "&#9828;";//"♣"; //9828 | U+2664 | White club suit
                    break;
                case "T":
                    return "&#9825;";//"♥"; //9825 | U+2661 | White heart suit
                    break;
                case "u":
                    return "&#9829;";//"♥"; //9829 | U+2665 | Black heart suit
                    break;
                case "U":
                    return "&#9825;";//"♥"; //9825 | U+2661 | White heart suit
                    break;
                case "w":
                    return "&#10071;";//"❗"; //10071 | U+2757 | Heavy exclamation mark symbol
                    break;
                case "W":
                    return "&#10071;";//"❗"; //10071 | U+2757 | Heavy exclamation mark symbol
                    break;
                case "x":
                    return "&#10062;";//"❞"; //10062 | U+275E | Heavy right-pointing angle quotation mark ornament
                    break;
                case "X":
                    return "&#10063;";//"❟"; //10063 | U+275F | Heavy low single comma quotation mark ornament
                    break;
                case "y":
                    return "&#10064;";//"❠"; //10064 | U+2760 | Heavy low double comma quotation mark ornament
                    break;
                case "Y":
                    return "&#10064;";//"❠"; //10064 | U+2760 | Heavy low double comma quotation mark ornament
                    break;
                case "z":
                    return "&#10061;";//"❝"; //10061 | U+275D | Heavy double turned comma quotation mark ornament
                    break;
                case "Z":
                    return "&#10061;";//"❝"; //10061 | U+275D | Heavy double turned comma quotation mark ornament
                    break;
                default:
                    if (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings"){
                        let wingCharCode =  getDingbatToUnicode(typefaceNode, buChar);
                        if (wingCharCode !== null){
                            return `&#${wingCharCode};`;
                        }
                    }
                    return `&#${(buChar.charCodeAt(0))};`;
            }
        }
        function getDingbatToUnicode(typefaceNode: string, buChar: string){
            // 原为 dingbatUnicode（未定义，运行时 ReferenceError），修正为已导入的 DINGBAT_UNICODE
            if (DINGBAT_UNICODE){
                let dingbat_code = (buChar.codePointAt(0) ?? 0) & 0xFFF;
                let char_unicode = null;
                let len = DINGBAT_UNICODE.length;
                let i = 0;
                while (len--) {
                    // blah blah
                    let item = DINGBAT_UNICODE[i];
                    // code 为字符串、dingbat_code 为数值，此处依赖 == 的类型转换语义
                    if (item.f == typefaceNode && Number(item.code) == dingbat_code) {
                        char_unicode = item.unicode;
                        break;
                    }
                    i++;
                }
                return char_unicode
        }
    }

    /**
     * alphaNumeric - 将数字转换为字母数字格式
     * @param {number} num - 数字
     * @param {string} upperLower - 大小写选项（upperCase或lowerCase）
     * @returns {string} 字母数字格式的字符串
     */
    function alphaNumeric(num: number, upperLower: string) {
        num = Number(num) - 1;
        let aNum = "";
        if (upperLower == "upperCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
        } else if (upperLower == "lowerCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
        }
        return aNum;
    }

    /**
     * hebrewAlphaNumeric - 将数字转换为希伯来字母编号格式
     * @param {number} num - 数字
     * @returns {string} 希伯来字母编号字符串
     */
    function hebrewAlphaNumeric(num: number) {
        num = Number(num) - 1;
        // 希伯来字母表（22个字母）
        const hebrewLetters = [
            'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט',
            'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ',
            'ק', 'ר', 'ש', 'ת'
        ];
        const hebrewLength = hebrewLetters.length;

        if (num < hebrewLength) {
            // 单字母（1-22）
            return hebrewLetters[num];
        } else if (num < hebrewLength * (hebrewLength + 1)) {
            // 双字母（23-506）
            const first = Math.floor(num / hebrewLength);
            const second = num % hebrewLength;
            return hebrewLetters[first] + hebrewLetters[second];
        } else {
            // 三字母（507+）
            const third = num % hebrewLength;
            const remaining = Math.floor(num / hebrewLength);
            const second = remaining % hebrewLength;
            const first = Math.floor(remaining / hebrewLength);
            return hebrewLetters[first] + hebrewLetters[second] + hebrewLetters[third];
        }
    }

    /**
     * archaicNumbers - 处理古数字格式
     * @param {Array} arr - 数字映射数组
     * @returns {Object} 包含format方法的对象
     */
    function archaicNumbers(arr: Array<[string | number | RegExp, string]>) {
        let arrParse = arr.slice().sort((a, b) => { return b[1].length - a[1].length });
        return {
            format: (n: number) => {
                let ret = '';
                for (const item of arr){
                    let num = item[0];
                    if (parseInt(String(num)) > 0) {
                        const numVal = parseInt(String(num));
                        for (; n >= numVal; n -= numVal) ret += item[1];
                    } else {
                        ret = ret.replace(num as string | RegExp, item[1]);
                    }
                }
                return ret;
            }
        }
    }

    /**
     * romanize - 将数字转换为罗马数字
     * @param {number} num - 数字
     * @returns {string} 罗马数字字符串
     */
    function romanize(num: number) {
        if (!+num)
            return false;
        let digits = String(+num).split(""),
            key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
                "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
                "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
            roman = "",
            i = 3;
        while (i--)
            roman = (key[+(digits.pop() ?? 0) + i * 10] || "") + roman;
        return Array(+digits.join("") + 1).join("M") + roman;
    }
    let hebrew2Minus = archaicNumbers([
            [1000, ''],
            [400, 'ת'],
            [300, 'ש'],
            [200, 'ר'],
            [100, 'ק'],
            [90, 'צ'],
            [80, 'פ'],
            [70, 'ע'],
            [60, 'ס'],
            [50, 'נ'],
            [40, 'מ'],
            [30, 'ל'],
            [20, 'כ'],
            [10, 'י'],
            [9, 'ט'],
            [8, 'ח'],
            [7, 'ז'],
            [6, 'ו'],
            [5, 'ה'],
            [4, 'ד'],
            [3, 'ג'],
            [2, 'ב'],
            [1, 'א'],
            [/יה/, 'ט״ו'],
            [/יו/, 'ט״ז'],
            [/([א-ת])([א-ת])$/, '$1״$2'],
            [/^([א-ת])$/, "$1׳"]
        ]);
    /**
     * 中文数字（一、二、三…；financial=true 时用壹、贰、叁…）
     * @param num 1-9999 的整数
     */
    function chineseNumeric(num: number, financial = false) {
        const digits = financial
            ? ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
            : ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
        const units = financial ? ['', '拾', '佰', '仟'] : ['', '十', '百', '千'];
        if (!isFinite(num) || num < 1 || num > 9999 || Math.floor(num) !== num) {
            return String(num);
        }
        if (num < 10) {
            return digits[num];
        }
        // 中文「十」习惯：10-19 写作 十/十一（财务写法为 壹拾/壹拾壹）
        if (num < 20) {
            const rest = num % 10 ? digits[num % 10] : '';
            return (financial ? digits[1] + units[1] : units[1]) + rest;
        }
        const str = String(num);
        const len = str.length;
        let out = '';
        for (let i = 0; i < len; i++) {
            const d = Number(str[i]);
            const unitIdx = len - 1 - i;
            if (d === 0) {
                // 中间的零只保留一个，末尾零省略
                if (out !== '' && i < len - 1 && !out.endsWith(digits[0])) {
                    out += digits[0];
                }
            } else {
                out += digits[d] + units[unitIdx];
            }
        }
        return out;
    }
    /**
     * getNumTypeNum - 根据数字类型获取格式化的数字
     * @param {string} numTyp - 数字类型
     * @param {number} num - 数字
     * @returns {string} 格式化的数字字符串
     */
    function getNumTypeNum(numTyp: string, num: number) {
        let rtrnNum = "";
        switch (numTyp) {
            case "arabicPeriod":
                rtrnNum = `${num}. `;
                break;
            case "arabicPlain":
                rtrnNum = `${num} `;
                break;
            case "arabicParenBoth":
                rtrnNum = `(${num}) `;
                break;
            case "arabicParenR":
                rtrnNum = `${num}) `;
                break;
            // East Asian 编号（中文环境下 PowerPoint/WPS 常用）
            case "chineseCounting":
            case "chineseCountingThousand":
            case "ideographDigital":
                rtrnNum = `${chineseNumeric(num)}`;
                break;
            case "chineseLegalSimplified":
                rtrnNum = `${chineseNumeric(num, true)}、`;
                break;
            case "ea1ChsPeriod":
            case "ea1ChtPeriod":
            case "ea1JpnChsDbPeriod":
            case "ea1JpnKorPeriod":
                rtrnNum = `${chineseNumeric(num)}、`;
                break;
            case "ea1ChsPlain":
            case "ea1ChtPlain":
            case "ea1JpnKorPlain":
                rtrnNum = `${chineseNumeric(num)} `;
                break;
            case "circleNumDbPlain":
                // ① ② ③ …
                rtrnNum = (num >= 1 && num <= 20) ? `${String.fromCodePoint(0x2460 + num - 1)} ` : `${num} `;
                break;
            case "alphaLcParenR":
                rtrnNum = `${alphaNumeric(num, "lowerCase")}) `;
                break;
            case "alphaLcPeriod":
                rtrnNum = `${alphaNumeric(num, "lowerCase")}. `;
                break;

            case "alphaUcParenR":
                rtrnNum = `${alphaNumeric(num, "upperCase")}) `;
                break;
            case "alphaUcPeriod":
                rtrnNum = `${alphaNumeric(num, "upperCase")}. `;
                break;

            case "romanUcPeriod":
                rtrnNum = `${romanize(num)}. `;
                break;
            case "romanLcParenR":
                rtrnNum = `${romanize(num)}) `;
                break;
            case "hebrew2Minus":
                // 希伯来字母编号：使用现代希伯来字母（א, ב, ג, ד, ...）类似英文字母编号
                rtrnNum = `${hebrewAlphaNumeric(num)}-`;
                break;
            default:
                rtrnNum = String(num);
        }
        return rtrnNum;
    }

    async function genSpanElement(node: XmlNode, rIndex: number | undefined, pNode: XmlNode, textBodyNode: XmlNode, pFontStyle: Record<string, unknown>, slideLayoutSpNode: XmlNode | undefined, idx: number | undefined, type: string, rNodeLength: number, warpObj: WarpObject, isBullate: boolean) {
            //https://codepen.io/imdunn/pen/GRgwaye ?
            let text_style = "";
            let lstStyle = textBodyNode["a:lstStyle"];
            let slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            let text = node["a:t"];

            // 检查是否启用 rtlCol 模式（每个单词单独一行）
            // 只有在没有 wrap 属性时，rtlCol="1" 才表示每个单词单独一行
            let rtlColAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rtlCol"]);
            let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "wrap"]);
            let isRTLCol = (rtlColAttr === "1" && wrapAttr === undefined);
            //let text_count = text.length;

            let openElemnt = "<span";//"<bdi";
            let closeElemnt = "</span>";// "</bdi>";
            let styleText = "";
            if (text === undefined && node["type"] !== undefined) {
                if (is_first_br) {
                    //openElemnt = "<br";
                    //closeElemnt = "";
                    //return "<br style='font-size: initial'>"
                    is_first_br = false;
                    return "<span class='line-break-br' ></span>";
                } else {
                    // styleText += "display: block;";
                    // openElemnt = "<span";
                    // closeElemnt = "</span>";
                }

                styleText += "display: block;";
                //openElemnt = "<span";
                //closeElemnt = "</span>";
            } else {

                is_first_br = true;
            }
            if (typeof text !== 'string') {
                text = PPTXXmlUtils.getTextByPathList(node, ["a:fld", "a:t"]);
                if (typeof text !== 'string') {
                    text = "&nbsp;";
                    //return "<span class='text-block '>&nbsp;</span>";
                }
                // if (text === undefined) {
                //     return "";
                // }
            }

            let pPrNode = pNode["a:pPr"];
            //lvl
            let lvl = 1;
            let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, warpObj);
            let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
            let { nodeLaout: pPrNodeLaout, nodeMaster: pPrNodeMaster } = layoutMasterNode;

            //Language
            let lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
            let isRtlLan = (lang !== undefined && RTL_LANGS_ARRAY.indexOf(lang) !== -1)?true:false;
            //rtl
            let getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            let isRTL = false;
            let dirStr = "ltr";
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
                dirStr = "rtl";
            }

            let linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
            let linkTooltip = "";
            let defLinkClr;
            if (linkID !== undefined) {
                const tip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
                if (tip !== undefined) {
                    linkTooltip = `title='${tip}'`;
                }
                defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
            } else {
                // Fallback to hover hyperlink (a:hlinkHover)
                linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "r:id"]);
                if (linkID !== undefined) {
                    const tip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkHover", "attrs", "tooltip"]);
                    if (tip !== undefined) {
                        linkTooltip = `title='${tip}'`;
                    }
                    defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);
                }
            }

            if (linkID !== undefined) {
                let linkClrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                let rPrlinkClr = PPTXStyleUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);


                //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
                // 对于超链接，优先使用主题中的超链接颜色，而不是文本运行中的颜色定义
                // 注释掉下面的覆盖逻辑，让超链接始终使用主题颜色
                // if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                //     defLinkClr = rPrlinkClr;
                // }
            }
            /////////////////////////////////////////////////////////////////////////////////////
            //getFontColor
            let fontClrPr = await PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
            let fontClrType = fontClrPr[2];
            //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
            if (fontClrType == "solid") {
                if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                    let colorValue = fontClrPr[0];
                    if (colorValue.length === 8) {
                        let colorObj = tinycolor(colorValue);
                        colorValue = colorObj.toRgbString();
                    } else {
                        colorValue = `#${colorValue}`;
                    }
                    styleText += `color: ${colorValue};`;
                }
                else if (linkID !== undefined && defLinkClr !== undefined) {
                    styleText += `color: #${defLinkClr};`;
                }

                if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                    styleText += `text-shadow:${fontClrPr[1]};`;
                }
                if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                    let highlightColorValue = fontClrPr[3];
                    if (highlightColorValue.length === 8) {
                        let colorObj = tinycolor(highlightColorValue);
                        highlightColorValue = colorObj.toRgbString();
                    } else {
                        highlightColorValue = `#${highlightColorValue}`;
                    }
                    styleText += `background-color: ${highlightColorValue};`;
                }
            } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
                if (fontClrType == "pattern") {
                    styleText += `background:${fontClrPr[0][0]};`;
                    if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                        styleText += `background-size:${fontClrPr[0][1]};`;//" 2px 2px;" +
                    }
                    if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                        styleText += `background-position:${fontClrPr[0][2]};`;//" 2px 2px;" +
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
                    //     "filter: " + fontClrPr[1].effcts + ";";
                } else if (fontClrType == "pic") {
                    styleText += `${fontClrPr[0]};`;
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                } else if (fontClrType == "gradient") {

                    let colorAry = fontClrPr[0].color;
                    // 归一化：非数组时（单个值/undefined）避免 .keys() 报错
                    if (!Array.isArray(colorAry)) {
                        colorAry = colorAry ? [colorAry] : [];
                    }
                    let rot = fontClrPr[0].rot;

                    styleText += `background: linear-gradient(${rot}deg,`;
                    for (const i of colorAry.keys()){
                        if (i == colorAry.length - 1) {
                            styleText += `#${colorAry[i]});`;
                        } else {
                            styleText += `#${colorAry[i]}, `;
                        }
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

                }
                styleText += `-webkit-background-clip: text;background-clip: text;color: transparent;`;
                if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
                    styleText += `-webkit-text-stroke: ${fontClrPr[1].border};`;
                }
                if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
                    styleText += `filter: ${fontClrPr[1].effcts};`;
                }
            }
            let font_size = PPTXStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
            //text_style += `font-size:${font_size};`
            
            text_style += `font-size:${font_size};` +
                // marLStr +
                "font-family:" + PPTXStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
                "font-weight:" + PPTXStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
                "font-style:" + PPTXStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
                "text-decoration:" + PPTXStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
                "text-align:" + PPTXStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
                "vertical-align:" + PPTXStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
            
            // Merge styleText into text_style
            text_style += styleText;
            //rNodeLength
            //console.log("genSpanElement node:", node, "lang:", lang, "isRtlLan:", isRtlLan, "span parent dir:", dirStr)
            if (isRtlLan) { //|| rIndex === undefined
                styleText += "direction:rtl;";
            }else{ //|| rIndex === undefined
                styleText += "direction:ltr;";
            }
            // } else if (dirStr == "rtl" && isRtlLan ) {
            //     styleText += "direction:rtl;";

            // } else if (dirStr == "ltr" && !isRtlLan ) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan){
            //     styleText += "direction:ltr;";
            // }else{
            //     styleText += "direction:inherit;";
            // }

            // if (dirStr == "rtl" && !isRtlLan) { //|| rIndex === undefined
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "rtl" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else if (dirStr == "ltr" && !isRtlLan) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else {
            //     styleText += "direction:inherit;";
            // }

            //     //`direction:${dirStr};`;
            //if (rNodeLength == 1 || rIndex == 0 ){
            //styleText += "display: table-cell;white-space: nowrap;";
            //}
            let highlight = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
            if (highlight !== undefined) {
                let highlightColor = PPTXStyleUtils.getSolidFill(highlight, undefined, undefined, warpObj);
                if (highlightColor !== undefined && highlightColor != "") {
                    if (highlightColor.length === 8) {
                        let colorObj = tinycolor(highlightColor);
                        highlightColor = colorObj.toRgbString();
                    } else {
                        highlightColor = `#${highlightColor}`;
                    }
                    styleText += `background-color:${highlightColor};`;
                }
                //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
            }

            //letter-spacing:
            let spcNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
                if (spcNode === undefined) {
                    spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
                }
            }
            if (spcNode !== undefined) {
                let ltrSpc = parseInt(spcNode) / 100; //pt
                styleText += `letter-spacing: ${ltrSpc}px;`;// + "pt;";
            }

            //Text Cap Types
            let capNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
                if (capNode === undefined) {
                    capNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
                }
            }
            if (capNode == "small" || capNode == "all") {
                styleText += "text-transform: uppercase";
            }
            //styleText += "word-break: break-word;";
            //console.log("genSpanElement node: ", node, ", capNode: ", capNode, ",pPrNodeLaout: ", pPrNodeLaout, ", pPrNodeMaster: ", pPrNodeMaster, "warpObj:", warpObj);

            let cssName = "";

            if (styleText in warpObj.styleTable) {
                cssName = warpObj.styleTable[styleText]["name"];
            } else {
                cssName = `_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
                warpObj.styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }
            let linkColorSyle = "";
            if (fontClrType == "solid" && linkID !== undefined) {
                // 对于超链接，始终使用主题中的超链接颜色，而不是文本运行中的颜色
                if (defLinkClr !== undefined) {
                    linkColorSyle = `style='color: #${defLinkClr};'`;
                }
            }

            if (linkID !== undefined && linkID != "") {
                const linkRes = warpObj["slideResObj"][linkID];
                let linkURL = linkRes && linkRes.target ? linkRes.target : "";
                const linkType = linkRes && linkRes.type ? linkRes.type : "";

                // 内部幻灯片跳转（如 PPT 中“跳到第 N 页”）：关系类型为 slide，
                // target 指向 slideX.xml。渲染为同页锚点跳转到对应幻灯片（id="slide-N"），
                // 而非打开不存在的原始 xml 文件；因此不设置 target='_blank'。
                let linkTargetAttr = " target='_blank'";
                if (linkType === "slide") {
                    const m = linkURL.match(/slide(\d+)\.xml$/i);
                    if (m) {
                        linkURL = `#slide-${m[1]}`;
                        linkTargetAttr = "";
                    }
                }

                linkURL = PPTXXmlUtils.escapeHtml(linkURL);
                // 处理文本：制表符、换行符、多个连续空格
                let processedText = text
                    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')  // 制表符转4个空格
                    .replace(/\n/g, "<br>")                      // 换行符转<br>
                    .replace(/  +/g, (spaces: string) => '&nbsp;'.repeat(spaces.length));  // 多个空格转&nbsp;

                // 在 rtlCol 模式下，每个单词单独一行
                if (isRTLCol) {
                    processedText = processedText.split(/\s+/).filter((word: string) => word.length > 0).join("<br>");
                }

                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + `'><a href='${linkURL}' ` + linkColorSyle + `  ${linkTooltip}${linkTargetAttr}>` +
                        processedText + "</a>" + closeElemnt;
            } else {
                // 处理文本：制表符、换行符、多个连续空格
                let processedText = text
                    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')  // 制表符转4个空格
                    .replace(/\n/g, "<br>")                      // 换行符转<br>
                    .replace(/  +/g, (spaces: string) => '&nbsp;'.repeat(spaces.length));  // 多个空格转&nbsp;

                // 在 rtlCol 模式下，每个单词单独一行
                if (isRTLCol) {
                    processedText = processedText.split(/\s+/).filter((word: string) => word.length > 0).join("<br>");
                }

                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + "'>" + processedText + closeElemnt;//"</bdi>";
            }

        }


        async function genTable(node: XmlNode, warpObj: WarpObject, shapeType: string) {
            let order = node["attrs"]!["order"];
            let tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);

            // 处理组合缩放 - 当table在group-abs类型组合中时需要应用缩放
            let workingXfrmNode = xfrmNode;
            if (shapeType === 'group-abs' && warpObj.currentGroupScale && xfrmNode) {
                const { scaleX, scaleY, childX, childY } = warpObj.currentGroupScale;

                // 创建缩放后的xfrmNode
                workingXfrmNode = JSON.parse(JSON.stringify(xfrmNode));

                // 缩放尺寸
                if (xfrmNode['a:ext'] && xfrmNode['a:ext'].attrs) {
                    const originalCx = parseInt(xfrmNode['a:ext'].attrs.cx);
                    const originalCy = parseInt(xfrmNode['a:ext'].attrs.cy);
                    workingXfrmNode['a:ext'].attrs.cx = Math.round(originalCx * scaleX);
                    workingXfrmNode['a:ext'].attrs.cy = Math.round(originalCy * scaleY);
                }

                // 调整位置(相对于childX/childY)
                if (xfrmNode['a:off'] && xfrmNode['a:off'].attrs) {
                    const originalOffX = parseInt(xfrmNode['a:off'].attrs.x);
                    const originalOffY = parseInt(xfrmNode['a:off'].attrs.y);

                    // 计算相对于childOff的偏移
                    const relativeX = originalOffX - (childX / SLIDE_FACTOR);
                    const relativeY = originalOffY - (childY / SLIDE_FACTOR);

                    // 应用缩放
                    workingXfrmNode['a:off'].attrs.x = Math.round(childX / SLIDE_FACTOR + relativeX * scaleX);
                    workingXfrmNode['a:off'].attrs.y = Math.round(childY / SLIDE_FACTOR + relativeY * scaleY);
                }
            }
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            let getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
            let getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
            // PPTX中的rtl属性通常表示文本方向，而不是表格布局方向
            // 不在table标签上设置dir属性，避免列顺序反转
            // 单元格内部的文本会根据自身的RTL设置正确显示
            let tblDir = "";
            // if (getTblPr !== undefined) {
            //     let isRTL = getTblPr["attrs"]["rtl"];
            //     tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
            // }
            let firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
            let firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
            let lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
            let lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
            let bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
            let bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
            //console.log("getTblPr: ", getTblPr);
            let tblStylAttrObj = {
                isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
                isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
                isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
                isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
                isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
                isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
            }

            let thisTblStyle;
            const tblStylesRoot = warpObj.tableStyles || {};
            const tbleStyleId = getTblPr["a:tableStyleId"];
            const tbleStylList = PPTXXmlUtils.getTextByPathList(tblStylesRoot, ["a:tblStyleLst", "a:tblStyle"]);
            const pickTblStyle = (styleId: string | undefined): any => {
                if (styleId === undefined || tbleStylList === undefined) {
                    return undefined;
                }
                if (tbleStylList.constructor === Array) {
                    for (const item of tbleStylList) {
                        if (item["attrs"] && item["attrs"]["styleId"] == styleId) {
                            return item;
                        }
                    }
                    return undefined;
                }
                return (tbleStylList["attrs"] && tbleStylList["attrs"]["styleId"] == styleId) ? tbleStylList : undefined;
            };
            thisTblStyle = pickTblStyle(tbleStyleId);
            // 未声明或未匹配到指定样式时，按 OOXML 规则回退到 a:tblStyleLst/@def 的默认样式
            if (thisTblStyle === undefined) {
                thisTblStyle = pickTblStyle(PPTXXmlUtils.getTextByPathList(tblStylesRoot, ["a:tblStyleLst", "attrs", "def"]));
            }
            if (thisTblStyle !== undefined) {
                thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
                warpObj["thisTbiStyle"] = thisTblStyle;
            }
            // 表格外边框不再画在 <table> 上：它由各单元格自身的边框决定（见 getTableCellParams），
            // 否则会盖住显式「无边框」的单元格，并与相邻单元格的边框叠成双线。
            const rightBorderGrid: any[][] = [];
            const bottomBorderGrid: any[][] = [];
            let tbl_bgcolor = "";
            let tbl_opacity = 1;
            let tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
            //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
            if (tbl_bgFillschemeClr !== undefined) {
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgFillschemeClr === undefined) {
                tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            // 表格无背景填充时 getSolidFill 返回 undefined，需归一为空串，否则会写进 style 变成字面量 "undefined"
            if (typeof tbl_bgcolor !== 'string') {
                tbl_bgcolor = "";
            }
            if (tbl_bgcolor !== "" && typeof tbl_bgcolor === 'string') {
                if (tbl_bgcolor.length === 8) {
                    let colorObj = tinycolor(tbl_bgcolor);
                    tbl_bgcolor = colorObj.toRgbString();
                } else {
                    tbl_bgcolor = `#${tbl_bgcolor}`;
                }
                tbl_bgcolor = `background-color: ${tbl_bgcolor};`;
            }
            ////////////////////////////////////////////////////////////////////////////////////////////
            let tableHtml = `<table ${tblDir} style='border-collapse: collapse;` +
                PPTXXmlUtils.getPosition(workingXfrmNode, node, undefined, undefined, shapeType) +
                PPTXXmlUtils.getSize(workingXfrmNode, undefined, undefined) +
                ` z-index: ${order};` +
                `${tbl_bgcolor}'>`;

            let trNodes = tableNode["a:tr"];
            if (trNodes.constructor !== Array) {
                trNodes = [trNodes];
            }
            // 组装单元格边框解析所需的上下文：
            // 命中区域（用于样式回退）+ 内部共享边（取左侧邻格的右边框 / 上方邻格的下边框）
            const buildBorderCtx = (i: number, j: number, cellCount: number, cellSource: string | undefined) => {
                const regions: string[] = [];
                if (cellSource !== undefined) regions.push(cellSource);
                if (i === 0 && tblStylAttrObj["isFrstRowAttr"] == 1) regions.push("a:firstRow");
                else if (i === (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1) regions.push("a:lastRow");
                else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1) regions.push((i % 2) === 0 ? "a:band2H" : "a:band1H");
                if (j === 0 && tblStylAttrObj["isFrstColAttr"] == 1) regions.push("a:firstCol");
                if (j === (cellCount - 1) && tblStylAttrObj["isLstColAttr"] == 1) regions.push("a:lastCol");
                regions.push("a:wholeTbl");
                return {
                    regions,
                    lastRow: i === (trNodes.length - 1),
                    lastCol: j === (cellCount - 1),
                    leftBorder: (j > 0 && rightBorderGrid[i]) ? rightBorderGrid[i][j - 1] : undefined,
                    topBorder: (i > 0 && bottomBorderGrid[i - 1]) ? bottomBorderGrid[i - 1][j] : undefined
                };
            };
            //if (trNodes.constructor === Array) {
                //multi rows
                let totalrowSpan = 0;
                let rowSpanAry: number[] = [];
                for (const i of trNodes.keys()){
                    //////////////rows Style ////////////Amir
                    let rowHeightParam = trNodes[i]["attrs"]["h"];
                    let rowHeight = 0;
                    let rowsStyl = "";
                    if (rowHeightParam !== undefined) {
                        rowHeight = parseInt(rowHeightParam) * SLIDE_FACTOR;
                        rowHeight = Math.round(rowHeight * 100) / 100;
                        rowsStyl += `height:${rowHeight}px;`;
                    }
                    let fillColor = "";
                    let row_borders: string | undefined = "";
                    let fontClrPr = "";
                    let fontWeight = "";
                    let band_1H_fillColor;
                    let band_2H_fillColor;

                    if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontColor = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontColor !== undefined) {
                                fontClrPr = local_fontColor;
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight
                            }
                        }
                    }

                    if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }

                    } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                        fillColor = "";
                        row_borders = undefined;
                        if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                            // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                            //check if there is a row bg
                            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== "") {
                                    fillColor = local_fillColor;
                                    band_2H_fillColor = local_fillColor;
                                }
                            }


                            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                        if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                            let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== undefined) {
                                    fillColor = local_fillColor;
                                    band_1H_fillColor = local_fillColor;
                                }
                            }
                            let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                                let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                                if (local_fontWeight != "") {
                                    fontWeight = local_fontWeight;
                                }
                            }
                        }

                    }
                    //last row
                    if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                        let bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                            // let local_colorOpacity = getColorOpacity(bgFillschemeClr);
                            // if(local_colorOpacity !== undefined){
                            //     colorOpacity = local_colorOpacity;
                            // }
                        }
                        let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            let local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        let rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }

                            let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }
                    // 行级边框不再写到 <tr> 上（浏览器在 border-collapse 下无法与单元格边框正确合并）：
                    // 首行/末行/斑马行命中的样式区域已通过 borderCtx.regions 参与单元格边框解析。
                    if (fontClrPr !== undefined && typeof fontClrPr === 'string') {
                        let tableColorValue = fontClrPr;
                        if (tableColorValue.length === 8) {
                            let colorObj = tinycolor(tableColorValue);
                            tableColorValue = colorObj.toRgbString();
                        } else {
                            tableColorValue = `#${tableColorValue}`;
                        }
                        rowsStyl += ` color: ${tableColorValue};`;
                    }
                    rowsStyl += ((fontWeight != "") ? ` font-weight:${fontWeight};` : "");
                    if (fillColor !== undefined && fillColor != "" && typeof fillColor === 'string') {
                        if (fillColor.length === 8) {
                            let colorObj = tinycolor(fillColor);
                            fillColor = colorObj.toRgbString();
                        } else {
                            fillColor = `#${fillColor}`;
                        }
                        //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + `,${colorOpacity});`;
                        rowsStyl += `background-color: ${fillColor};`;
                    }
                    tableHtml += `<tr style='${rowsStyl}'>`;
                    ////////////////////////////////////////////////

                    let tcNodes = trNodes[i]["a:tc"];
                    if (tcNodes !== undefined) {
                        if (tcNodes.constructor === Array) {
                            //multi columns
                            let j = 0;
                            if (rowSpanAry.length == 0) {
                                rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(() => { return 0 });
                            }
                            let totalColSpan = 0;
                            while (j < tcNodes.length) {
                                if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                                    let a_sorce;
                                    //j=0 : first col
                                    if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                        a_sorce = "a:firstCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                            a_sorce = "a:seCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                            a_sorce = "a:neCell";
                                        }
                                    } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                        !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                        !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                        j != (tcNodes.length - 1)) {

                                        if ((j % 2) != 0) {

                                            let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                            if (aBandNode === undefined) {
                                                aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                                if (aBandNode !== undefined) {
                                                    a_sorce = "a:band2V";
                                                }
                                            } else {
                                                a_sorce = "a:band2V";
                                            }

                                        }
                                    }

                                    if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                        a_sorce = "a:lastCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                            a_sorce = "a:swCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                            a_sorce = "a:nwCell";
                                        }
                                    }

                                    let cellParmAry = await getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj,
                                        buildBorderCtx(i, j, tcNodes.length, a_sorce))
                                    let text = cellParmAry[0];
                                    let colStyl = cellParmAry[1];
                                    let cssName = cellParmAry[2];
                                    let rowSpan = cellParmAry[3];
                                    let colSpan = cellParmAry[4];
                                    if (!rightBorderGrid[i]) rightBorderGrid[i] = [];
                                    if (!bottomBorderGrid[i]) bottomBorderGrid[i] = [];
                                    rightBorderGrid[i][j] = cellParmAry[5];
                                    bottomBorderGrid[i][j] = cellParmAry[6];



                                    if (rowSpan !== undefined) {
                                        totalrowSpan++;
                                        rowSpanAry[j] = parseInt(rowSpan) - 1;
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' rowspan ='` +
                                            parseInt(rowSpan) + `' style='${colStyl}'>` + text + "</td>";
                                    } else if (colSpan !== undefined) {
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' colspan = '` +
                                            parseInt(colSpan) + `' style='${colStyl}'>` + text + "</td>";
                                        totalColSpan = parseInt(colSpan) - 1;
                                    } else {
                                        tableHtml += `<td class='${cssName}' data-row='` + i + `,${j}' style = '` + colStyl + `'>${text}</td>`;
                                    }

                                } else {
                                    if (rowSpanAry[j] != 0) {
                                        rowSpanAry[j] -= 1;
                                    }
                                    if (totalColSpan != 0) {
                                        totalColSpan--;
                                    }
                                }
                                j++;
                            }
                        } else {
                            //single column 

                            let a_sorce;
                            if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:firstCol";

                            } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                                let aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                if (aBandNode === undefined) {
                                    aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                    if (aBandNode !== undefined) {
                                        a_sorce = "a:band2V";
                                    }
                                } else {
                                    a_sorce = "a:band2V";
                                }
                            }

                            if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:lastCol";
                            }


                            let cellParmAry = await getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj,
                                buildBorderCtx(i, 0, 1, a_sorce))
                            let text = cellParmAry[0];
                            let colStyl = cellParmAry[1];
                            let cssName = cellParmAry[2];
                            let rowSpan = cellParmAry[3];
                            if (!rightBorderGrid[i]) rightBorderGrid[i] = [];
                            if (!bottomBorderGrid[i]) bottomBorderGrid[i] = [];
                            rightBorderGrid[i][0] = cellParmAry[5];
                            bottomBorderGrid[i][0] = cellParmAry[6];

                            if (rowSpan !== undefined) {
                                tableHtml += `<td  class='${cssName}' rowspan='` + parseInt(rowSpan) + `' style = '${colStyl}'>` + text + "</td>";
                            } else {
                                tableHtml += `<td class='${cssName}' style='` + colStyl + `'>${text}</td>`;
                            }
                        }
                    }
                    tableHtml += "</tr>";
                }
                //////////////////////////////////////////////////////////////////////////////////
            

            return tableHtml;
        }
        
        async function getTableCellParams(tcNodes: XmlNode, getColsGrid: XmlNode[], row_idx: number | undefined, col_idx: number | undefined, thisTblStyle: XmlNode, cellSource: string | undefined, warpObj: WarpObject, borderCtx?: {
            /** 该单元格命中的样式区域（按优先级，末尾应含 a:wholeTbl） */
            regions: string[];
            /** 是否为最后一行 / 最后一列（决定回退用 bottom/right 还是 insideH/insideV） */
            lastRow: boolean;
            lastCol: boolean;
            /** 左侧邻格的右边框 / 上方邻格的下边框（内部共享边，由邻格决定） */
            leftBorder?: any;
            topBorder?: any;
        }) {
            //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
            //text, cell-width, cell-borders, 
            //let text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
            let rowSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
            let colSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
            let vMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
            let hMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
            let colStyl = "word-wrap: break-word;";
            let colWidth;
            let celFillColor = "";
            let col_borders = "";
            let colFontClrPr = "";
            let colFontWeight = "";
            // 四边边框节点：XmlNode = 实线边框，'none' = 显式无边框，null = 无边，undefined = 未定
            let lin_bottm: any,
                lin_top: any,
                lin_left: any,
                lin_right: any,
                lin_bottom_left_to_top_right: XmlNode | undefined,
                lin_top_left_to_bottom_right: XmlNode | undefined;
            
            let colSapnInt = parseInt(colSpan);
            let total_col_width = 0;
            if (!isNaN(colSapnInt) && colSapnInt > 1){
                for (let k = 0; k < colSapnInt ; k++) {
                    total_col_width += parseInt (PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx! + k], ["attrs", "w"]));
                }
            }else{
                total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
            }
            

            let text = await PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, "table", undefined, warpObj, total_col_width);//tableStyles

            if (total_col_width != 0 /*&& row_idx == 0*/) {
                colWidth = parseInt(String(total_col_width)) * SLIDE_FACTOR;
                colWidth = Math.round(colWidth * 100) / 100;
                colStyl += `width:${colWidth}px;`;
            }

            //cell horizontal alignment
            let cellAlign = "";
            // 获取表格单元格文本的RTL状态，以确定正确的对齐方式
            let textBodyNode = tcNodes["a:txBody"];
            let prg_dir = "";
            if (textBodyNode !== undefined) {
                // 获取段落的RTL状态
                let pNodes = textBodyNode["a:p"];
                if (pNodes !== undefined) {
                    if (Array.isArray(pNodes) && pNodes.length > 0) {
                        prg_dir = PPTXStyleUtils.getPregraphDir(pNodes[0], textBodyNode, 0, "table", warpObj);
                    } else {
                        prg_dir = PPTXStyleUtils.getPregraphDir(pNodes, textBodyNode, 0, "table", warpObj);
                    }
                }
            }
            let isRTL = (prg_dir == "pregraph-rtl");

            // 获取水平对齐属性
            if (textBodyNode !== undefined) {
                let pNodes = textBodyNode["a:p"];
                if (pNodes !== undefined) {
                    let firstP = Array.isArray(pNodes) ? pNodes[0] : pNodes;
                    let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(firstP, textBodyNode, 0, "table", prg_dir, warpObj);
                    if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                        cellAlign = isRTL ? "text-align: left;" : "text-align: right;";
                    } else if (horizontalAlign === "h-mid") {
                        cellAlign = "text-align: center;";
                    } else if (horizontalAlign === "h-left-rtl") {
                        cellAlign = "text-align: right;";
                    } else {
                        cellAlign = isRTL ? "text-align: right;" : "text-align: left;";
                    }
                }
            }
            if (cellAlign !== "") {
                colStyl += cellAlign;
            }

            // 单元格垂直对齐：a:tcPr/@anchor（t=顶端, ctr=居中, b=底端；just/dist 按居中近似）。
            // 不设置的话 <td> 默认为垂直居中，与 PowerPoint/WPS 的「顶端」表现不一致。
            // 省略该属性时按 ECMA-376 取默认值 t（顶端）。
            let anchorAttr = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "attrs", "anchor"]);
            if (anchorAttr === undefined) {
                anchorAttr = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "anchor"]);
            }
            // 兼容旧版 anchorCtr（垂直居中）
            const anchorCtr = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "attrs", "anchorCtr"]);
            const isAnchorCenter = anchorCtr === "1" || anchorCtr === "true";
            if (isAnchorCenter) {
                colStyl += "vertical-align:middle;";
            } else {
                const vAlign = (anchorAttr === "b") ? "bottom"
                    : ((anchorAttr === "ctr" || anchorAttr === "just" || anchorAttr === "dist") ? "middle" : "top");
                colStyl += `vertical-align:${vAlign};`;
            }

            //cell bords
            // 按 OOXML / PowerPoint 的规则解析单元格四条边：
            //  · 单元格自身的 lnR/lnB 决定右/下边，未指定则回退表格样式（内部边用 insideV/insideH，
            //    最后一行/列用 bottom/right）；
            //  · 表格内部共享边取「左侧单元格的 lnR / 上方单元格的 lnB」，单元格自身的 lnL/lnT
            //    只对最左列/最顶行（外边框）生效——否则同一位置会画出两条不同颜色的线；
            //  · 显式 <a:noFill/> 表示该边无边框，不能被样式网格补上。
            lin_bottom_left_to_top_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
            lin_top_left_to_bottom_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnTlToBr"]);

            const isNoFillLn = (n: any) =>
                n !== undefined && n !== null && typeof n === 'object' && !Array.isArray(n) && n["a:noFill"] !== undefined;
            // 'none' 表示显式无边框，null 表示确实没有该边，undefined 表示「由调用方决定」
            const normLn = (n: any): any => (n === undefined ? undefined : (isNoFillLn(n) ? 'none' : n));
            const regions = (borderCtx && borderCtx.regions) ? borderCtx.regions : ["a:wholeTbl"];
            const styleLn = (side: string): any => {
                if (thisTblStyle === undefined) return null;
                for (const region of regions) {
                    const ln = PPTXXmlUtils.getTextByPathList(thisTblStyle, [region, "a:tcStyle", "a:tcBdr", side, "a:ln"]);
                    if (ln !== undefined) return normLn(ln);
                }
                return null;
            };
            const ownLn = (side: string) => normLn(PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:ln" + side]));
            const isLastRow = (borderCtx && borderCtx.lastRow) ? true : false;
            const isLastCol = (borderCtx && borderCtx.lastCol) ? true : false;

            const ownRight = ownLn("R");
            const ownBottom = ownLn("B");
            lin_right = ownRight !== undefined ? ownRight : styleLn(isLastCol ? "a:right" : "a:insideV");
            lin_bottm = ownBottom !== undefined ? ownBottom : styleLn(isLastRow ? "a:bottom" : "a:insideH");
            const ownLeft = ownLn("L");
            const ownTop = ownLn("T");
            lin_left = (borderCtx && borderCtx.leftBorder !== undefined)
                ? borderCtx.leftBorder
                : (ownLeft !== undefined ? ownLeft : styleLn("a:left"));
            lin_top = (borderCtx && borderCtx.topBorder !== undefined)
                ? borderCtx.topBorder
                : (ownTop !== undefined ? ownTop : styleLn("a:top"));

            const emitBorder = (side: string, resolved: any) => {
                if (resolved === undefined) return;
                if (resolved === null || resolved === 'none') {
                    colStyl += `border-${side}:none;`;
                    return;
                }
                const css = PPTXStyleUtils.getBorder(resolved, undefined, false, "", warpObj);
                // getBorder 在缺少颜色时会产出非法 CSS（如 "1.33px solid "），这类边按无边框处理
                if (typeof css === 'string' && css !== "" && /#|rgb/.test(css)) {
                    colStyl += `border-${side}:${css};`;
                } else {
                    colStyl += `border-${side}:none;`;
                }
            };
            emitBorder("left", lin_left);
            emitBorder("top", lin_top);
            emitBorder("right", lin_right);
            emitBorder("bottom", lin_bottm);

            //cell fill color custom
            let getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
            if (getCelFill !== undefined && getCelFill != "") {
                let cellObj = {
                    "p:spPr": getCelFill
                };
                celFillColor = await PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide")
            }

            //cell fill color theme
            if (celFillColor == "" || celFillColor == "background-color: inherit;") {
                let bgFillschemeClr;
                if (cellSource !== undefined)
                    bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    let local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        celFillColor = ` background-color: #${local_fillColor};`;
                    }
                }
            }
            let cssName = "";
            if (celFillColor !== undefined && celFillColor != "") {
                if (celFillColor in warpObj.styleTable) {
                    cssName = warpObj.styleTable[celFillColor]["name"];
                } else {
                    cssName = `_tbl_cell_css_${(Object.keys(warpObj.styleTable).length + 1)}`;
                    warpObj.styleTable[celFillColor] = {
                        "name": cssName,
                        "text": celFillColor
                    };
                }

            }

            //border
            // let borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
            // if (borderStyl !== undefined) {
            //     let local_col_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
            //     if (local_col_borders != "") {
            //         col_borders = local_col_borders;
            //     }
            // }
            // if (col_borders != "") {
            //     colStyl += col_borders;
            // }

            //Text style
            let rowTxtStyl;
            if (cellSource !== undefined) {
                rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
            }
            // if (rowTxtStyl === undefined) {
            //     rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            // }
            if (rowTxtStyl !== undefined) {
                let local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    colFontClrPr = local_fontClrPr;
                }
                let local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    colFontWeight = local_fontWeight;
                }
            }
            colStyl += ((colFontClrPr !== "" && typeof colFontClrPr === 'string') ?
                ((colFontClrPr.length === 8) ?
                    (() => {
                        let colorObj = tinycolor(colFontClrPr);
                        return `color: ${colorObj.toRgbString()};`;
                    })() :
                    `color: #${colFontClrPr};`) : "");
            colStyl += ((colFontWeight != "") ? ` font-weight:${colFontWeight};` : "");

            // 末尾两项返回解析好的右/下边框，供右侧与下方单元格作为共享边复用
            return [text, colStyl, cssName, rowSpan, colSpan, lin_right, lin_bottm];
        }
const PPTXTextUtils = {
        genTextBody,
        genBuChar,
        getHtmlBullet,
        getDingbatToUnicode,
        genSpanElement,
        genTable,
        getTableCellParams,
        alphaNumeric,
        archaicNumbers,
        romanize,
        getNumTypeNum,
    };

export { PPTXTextUtils };
export default PPTXTextUtils;
