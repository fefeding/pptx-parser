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

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR, RTL_LANGS_ARRAY, DINGBAT_UNICODE } from '../core/constants.js';
import { genChart } from './chart.js';
import tinycolor from '../core/tinycolor.js';
let is_first_br = false;



function getTextWidth(html) {
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

    function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
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

            // 获取bodyPr的内边距设置
            let bodyPrPadding = getBodyPrPadding(textBodyNode, type, anchor);
            text += bodyPrPadding;

            let pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
            let wrapAttr = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["attrs", "wrap"]);
            let spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode["a:bodyPr"], ["a:spAutoFit"]);
            let isNoWrap = (wrapAttr === "none");
            let isAutoFit = (spAutoFitNode !== undefined);
            //console.log("genTextBody spNode: ", PPTXXmlUtils.getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));
            
            let apNode = textBodyNode["a:p"];
            if (apNode.constructor !== Array) {
                apNode = [apNode];
            }

            for (let i = 0; i < apNode.length; i++) {
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
                    brNode.forEach((item, indx) => {
                        item.type = "br";
                    });
                    if (brNode.length > 1) {
                        brNode.shift();
                    }
                    rNode = rNode.concat(brNode)
                    //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
                    rNode.sort((a, b) => {
                        return a.attrs.order - b.attrs.order;
                    });
                    //console.log("sorted rNode:",rNode)
                }
                //rtlStr = "";//`dir='${isRTL}'`;
                let styleText = "";
                let marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
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
                    cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                    warpObj.styleTable[styleText] = {
                        "name": cssName,
                        "text": styleText
                    };
                }
                //console.log("textBodyNode: ", textBodyNode["a:lstStyle"])
                let prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                let prg_height_node;// = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
                let sld_prg_width_val = (prg_width_node !== undefined && prg_width_node !== null) ? Math.round(parseInt(prg_width_node) * SLIDE_FACTOR * 100) / 100 : null;
                let sld_prg_width = "";
                if (sld_prg_width_val !== null && !isNoWrap) {
                    sld_prg_width = "width:" + sld_prg_width_val + "px;";
                } else if (sld_prg_width_val === null) {
                    sld_prg_width = "width:inherit;";
                }
                let sld_prg_height = ""; // 移除高度设置，避免段落叠加
                let prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
                let isRTL = (prg_dir == "pregraph-rtl");
                text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + ` ${prg_dir} ` + cssName + "' >";
                let buText_ary = genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
                let isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
                let bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;

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
                    let prgr_text = genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
                    if (isBullate) {
                        total_text_len += getTextWidth(prgr_text);
                    }
                    prgrph_text += prgr_text;
                } else if (rNode !== undefined) {
                    // with multi r
                    let previousStyle = {};
                    for (let j = 0; j < rNode.length; j++) {
                        // 如果当前元素没有sz属性，使用前面元素的样式
                        if (rNode[j]["a:rPr"] && !rNode[j]["a:rPr"]["attrs"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"] = { "sz": previousStyle["sz"] };
                        } else if (rNode[j]["a:rPr"] && rNode[j]["a:rPr"]["attrs"] && !rNode[j]["a:rPr"]["attrs"]["sz"] && previousStyle["sz"]) {
                            rNode[j]["a:rPr"]["attrs"]["sz"] = previousStyle["sz"];
                        }
                        
                        let prgr_text = genSpanElement(rNode[j], j, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
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

                prg_width_node = parseInt(prg_width_node) * SLIDE_FACTOR - bu_width - mrgin_val;
                prg_width_node = Math.round(prg_width_node * 100) / 100;
                if (isBullate) {
                    //get prg_width_node if there is a bulltes
                    //console.log("total_text_len: ", total_text_len, "prg_width_node:", prg_width_node)

                    if (total_text_len < prg_width_node ){
                        prg_width_node = total_text_len + bu_width;
                    }
                }
                // 如果没有明确设置wrap="none"或spAutoFit，默认不设置内层div的宽度，让文本自然流动
                let prg_width = "";
                let textContainerWidth = "width: 100%;"; // 默认宽度
                if (isRTL && isBullate) {
                    // RTL 模式下有项目符号时，文本容器不设 100%，让内容自适应
                    textContainerWidth = "";
                }
                if (prg_width_node !== undefined && prg_width_node !== null && !isNoWrap) {
                    // 只有明确不需要换行时才设置宽度
                    prg_width = "width:" + (Math.round(prg_width_node * 100) / 100) + "px;";
                }
                let whiteSpaceStyle = isNoWrap ? "white-space: nowrap;" : "white-space: pre-wrap;";
                let horizontalAlign = PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj, spNode);


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
                // 为了确保右对齐生效，添加flex布局的justify-content属性
                let flexStyle = "";
                if (horizontalAlign === "h-right" || horizontalAlign === "h-right-rtl") {
                    flexStyle = "justify-content: flex-end;";
                } else if (horizontalAlign === "h-mid") {
                    flexStyle = "justify-content: center;";
                } else {
                    flexStyle = "justify-content: flex-start;";
                }
                text += "<div style='display: flex;" + flexStyle + textContainerWidth + "'>";
                text += "<div style='direction: initial;" + whiteSpaceStyle + margin + textAlignStyle + "'>";
                text += prgrph_text;
                text += "</div>";
                text += "</div>";
                // 在 RTL 模式下，项目符号放在最后（右边）
                if (isRTL && isBullate && buText_ary[0] !== undefined) {
                    text += buText_ary[0];
                }
                text += "</div>";
            }

            // 关闭bodyPr内边距div（如果存在）
            if (type === "textBox" || type === "shape") {
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
    function getBodyPrPadding(textBodyNode, type, anchor) {
        let paddingStyle = "";
        
        // 获取bodyPr的各个内边距属性
        let lIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "lIns"]);
        let tIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "tIns"]);
        let rIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "rIns"]);
        let bIns = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "attrs", "bIns"]);

        // 只有在文本框或形状类型时才应用内边距
        if (type === "textBox" || type === "shape") {
            // 根据PPTX规范，bodyPr的ins属性单位是EMU（English Metric Units）
            // 1 inch = 914400 EMU, 1 inch = 96px, 所以 1 EMU = 96/914400 px ≈ 0.000105 px
            // 如果没有设置内边距，使用默认值：
            // - tIns 和 bIns 默认值为 0.05 inch = 45720 EMU ≈ 4.8px
            // - lIns 和 rIns 默认值为 0.1 inch = 91440 EMU ≈ 9.6px
            let lInsPx = lIns ? (parseInt(lIns) * SLIDE_FACTOR).toFixed(2) : (0.1 * 96).toFixed(2);  // 默认0.1 inch
            let tInsPx = tIns ? (parseInt(tIns) * SLIDE_FACTOR).toFixed(2) : (0.05 * 96).toFixed(2); // 默认0.05 inch
            let rInsPx = rIns ? (parseInt(rIns) * SLIDE_FACTOR).toFixed(2) : (0.1 * 96).toFixed(2);  // 默认0.1 inch
            let bInsPx = bIns ? (parseInt(bIns) * SLIDE_FACTOR).toFixed(2) : (0.05 * 96).toFixed(2); // 默认0.05 inch

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
        
        function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
            //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
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
                dfltBultColor = PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
                color_tye = dfltBultColor[2];
                dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
            } else {
                return "";
            }
            //console.log("Bullet Size: " + bultSize);

            let bullet = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
            /////////////////////////////////////////////////////////////////


            let pPrNode = node["a:pPr"];
            let BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }

            let buType = "TYPE_NONE";

            let layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

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
                    let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                    bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
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
            //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
            if (marLNode !== undefined) {
                let marginLeft = parseInt(marLNode) * SLIDE_FACTOR;
                // 在 RTL 模式下，项目符号在文本右边，需要左边距靠近文本
                // 在 LTR 模式下，项目符号在文本左边，需要右边距靠近文本
                if (isRTL) {
                    marLStr = "padding-left: 5px;";  // 项目符号左边的小间隔（靠近文本）
                } else {
                    marLStr = "padding-left:";
                    marLStr += ((marginLeft + indent < 0) ? 0 : (marginLeft + indent)) + "px;";
                }
                // margin_val 始终返回实际值，用于文本容器宽度计算
                margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
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
                marRStr = "padding-right:";
                marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
            }

            if (buType != "TYPE_NONE") {
                //let buFontAttrs = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
            }
            //console.log("Bullet Type: " + buType);
            //console.log("NumericTypr: " + buNum);
            //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
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
            } else {
                if (pFontStyle !== undefined) {
                    //console.log("genBuChar pFontStyle: ", pFontStyle)
                    defBultColor = PPTXStyleUtils.getSolidFill(pFontStyle, undefined, undefined, warpObj);
                }
            }
            if (defBultColor === undefined || defBultColor == "NONE") {
                bultColor = dfltBultColor;
            } else {
                bultColor = [defBultColor, "", "solid"];
                color_tye = "solid";
            }
            //console.log("genBuChar node:", node, "pPrNode", pPrNode, " buClrNode: ", buClrNode, "defBultColor:", defBultColor,"dfltBultColor:" , dfltBultColor , "bultColor:", bultColor)

            //console.log("genBuChar: buClrNode: ", buClrNode, "bultColor", bultColor)
            //get definde bullet SIZE
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        let prcnt = parseInt(buFontSize) / 100000;
                        //let dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                }else{
                    bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
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
                        let dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                } else {
                    bultSize = (parseInt(buFontSize) / 100) * FONT_SIZE_FACTOR + "px";
                }
            }
            if (buFontSize === undefined) {
                bultSize = dfltBultSize;
            }
            font_val = parseInt(bultSize, "px");
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                let typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
                let typeface = "";
                let isWingdingsFont = false;
                if (typefaceNode !== undefined) {
                    isWingdingsFont = (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3" || typefaceNode == "Webdings");
                    typeface = "font-family: " + typefaceNode;
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
                            bulletColorValue = "#" + bulletColorValue;
                        }
                        bullet += "color:" + bulletColorValue + "; ";
                    }
                    if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                        bullet += "text-shadow:" + bultColor[1] + ";";
                    }
                    //no highlight/background-color to bullet
                    // if (bultColor[3] !== undefined && bultColor[3] != "") {
                    //     styleText += "background-color: #" + bultColor[3] + ";";
                    // }
                } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
                    if (color_tye == "pattern") {
                        bullet += "background:" + bultColor[0][0] + ";";
                        if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                            bullet += "background-size:" + bultColor[0][1] + ";";//" 2px 2px;" +
                        }
                        if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                            bullet += "background-position:" + bultColor[0][2] + ";";//" 2px 2px;" +
                        }
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                        //     "filter: " + bultColor[1].effcts + ";";
                    } else if (color_tye == "pic") {
                        bullet += bultColor[0] + ";";
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";

                    } else if (color_tye == "gradient") {

                        let colorAry = bultColor[0].color;
                        let rot = bultColor[0].rot;

                        bullet += `background: linear-gradient(${rot}deg,`;
                        for (let i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bullet += "#" + colorAry[i] + ");";
                            } else {
                                bullet += "#" + colorAry[i] + ", ";
                            }
                        }
                        // bullet += "color: transparent;" +
                        //     "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    bullet += "-webkit-background-clip: text;" +
                        "background-clip: text;" +
                        "color: transparent;";
                    if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                        bullet += "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                        bullet += "filter: " + bultColor[1].effcts + ";";
                    }
                }

                if (isRTL) {
                    //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
                    bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
                }
                let isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
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
                
                bullet += "display: flex; align-items: center;'><div>" + htmlBu + "</div></div>";
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

                bullet = "<div style='" + marLStr + marRStr;
                if (bultColor[0] !== undefined && bultColor[0] != "") {
                    let bulletNumColorValue = bultColor[0];
                    if (bulletNumColorValue.length === 8) {
                        let colorObj = tinycolor(bulletNumColorValue);
                        bulletNumColorValue = colorObj.toRgbaString();
                    } else {
                        bulletNumColorValue = "#" + bulletNumColorValue;
                    }
                    bullet += "color:" + bulletNumColorValue + ";";
                }
                bullet += `font-size:${bultSize};`;
                if (isRTL) {
                    bullet += "white-space: nowrap ;direction:rtl;";
                } else {
                    bullet += "white-space: nowrap ;direction:ltr;";
                }
                bullet += `display: flex; align-items: center;'><div>${bulletText}</div></div>`;

            } else if (buType == "TYPE_BULPIC") { //PIC BULLET
                // let marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * SLIDE_FACTOR;
                // let marginRight = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * SLIDE_FACTOR;

                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }
                // //console.log("marginRight: "+marginRight)
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
                    //console.log("imgPath: ", imgPath);
                    if (imgPath === undefined) {
                        console.warn("Bullet image reference not found for buPicId:", buPicId);
                        buImg = "";
                    } else {
                        let imgFile = warpObj["zip"].file(imgPath);
                        if (imgFile === null) {
                            console.warn("Bullet image file not found:", imgPath);
                            buImg = "";
                        } else {
                            let imgArrayBuffer = imgFile.asArrayBuffer();
                            let imgExt = imgPath.split(".").pop();
                            let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                            buImg = `<img src='data:${imgMimeType};base64,` + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>"// height: 100%
                            //console.log(`imgPath: ${imgPath}\nimgMimeType: `+imgMimeType)
                        }
                    }
                }
                if (buPicId === undefined) {
                    buImg = "&#8227;";
                }
                bullet = "<div style='" + marLStr + marRStr +
                    `width:${bultSize};display: flex; align-items: center;`;// +
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
            //console.log("genBuChar: width: ", $(bullet).outerWidth())
            return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
        }
        function getHtmlBullet(typefaceNode, buChar) {
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
                    return "&#" + (buChar.charCodeAt(0)) + ";";
            }
        }
        function getDingbatToUnicode(typefaceNode, buChar){
            if (dingbatUnicode){
                let dingbat_code = buChar.codePointAt(0) & 0xFFF;
                let char_unicode = null;
                let len = dingbatUnicode.length;
                let i = 0;
                while (len--) {
                    // blah blah
                    let item = dingbatUnicode[i];
                    if (item.f == typefaceNode && item.code == dingbat_code) {
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
    function alphaNumeric(num, upperLower) {
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
     * archaicNumbers - 处理古数字格式
     * @param {Array} arr - 数字映射数组
     * @returns {Object} 包含format方法的对象
     */
    function archaicNumbers(arr) {
        let arrParse = arr.slice().sort((a, b) => { return b[1].length - a[1].length });
        return {
            format: (n) => {
                let ret = '';
                for (let i = 0; i < arr.length; i++) {
                    let num = arr[i][0];
                    if (parseInt(num) > 0) {
                        for (; n >= num; n -= num) ret += arr[i][1];
                    } else {
                        ret = ret.replace(num, arr[i][1]);
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
    function romanize(num) {
        if (!+num)
            return false;
        let digits = String(+num).split(""),
            key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
                "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
                "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
            roman = "",
            i = 3;
        while (i--)
            roman = (key[+digits.pop() + (i * 10)] || "") + roman;
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
     * getNumTypeNum - 根据数字类型获取格式化的数字
     * @param {string} numTyp - 数字类型
     * @param {number} num - 数字
     * @returns {string} 格式化的数字字符串
     */
    function getNumTypeNum(numTyp, num) {
        let rtrnNum = "";
        switch (numTyp) {
            case "arabicPeriod":
                rtrnNum = num + ". ";
                break;
            case "arabicParenR":
                rtrnNum = num + ") ";
                break;
            case "alphaLcParenR":
                rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
                break;
            case "alphaLcPeriod":
                rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
                break;

            case "alphaUcParenR":
                rtrnNum = alphaNumeric(num, "upperCase") + ") ";
                break;
            case "alphaUcPeriod":
                rtrnNum = alphaNumeric(num, "upperCase") + ". ";
                break;

            case "romanUcPeriod":
                rtrnNum = romanize(num) + ". ";
                break;
            case "romanLcParenR":
                rtrnNum = romanize(num) + ") ";
                break;
            case "hebrew2Minus":
                rtrnNum = hebrew2Minus.format(num) + "-";
                break;
            default:
                rtrnNum = num;
        }
        return rtrnNum;
    }

    function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
            //https://codepen.io/imdunn/pen/GRgwaye ?
            let text_style = "";
            let lstStyle = textBodyNode["a:lstStyle"];
            let slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            let text = node["a:t"];
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
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

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
                linkTooltip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
                if (linkTooltip !== undefined) {
                    linkTooltip = `title='${linkTooltip}'`;
                }
                defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

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
            let fontClrPr = PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
            let fontClrType = fontClrPr[2];
            //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
            if (fontClrType == "solid") {
                if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                    let colorValue = fontClrPr[0];
                    if (colorValue.length === 8) {
                        let colorObj = tinycolor(colorValue);
                        colorValue = colorObj.toRgbString();
                    } else {
                        colorValue = "#" + colorValue;
                    }
                    styleText += "color: " + colorValue + ";";
                }
                else if (linkID !== undefined && defLinkClr !== undefined) {
                    styleText += `color: #${defLinkClr};`;
                }

                if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                    styleText += "text-shadow:" + fontClrPr[1] + ";";
                }
                if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                    let highlightColorValue = fontClrPr[3];
                    if (highlightColorValue.length === 8) {
                        let colorObj = tinycolor(highlightColorValue);
                        highlightColorValue = colorObj.toRgbString();
                    } else {
                        highlightColorValue = "#" + highlightColorValue;
                    }
                    styleText += "background-color: " + highlightColorValue + ";";
                }
            } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
                if (fontClrType == "pattern") {
                    styleText += "background:" + fontClrPr[0][0] + ";";
                    if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                        styleText += "background-size:" + fontClrPr[0][1] + ";";//" 2px 2px;" +
                    }
                    if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                        styleText += "background-position:" + fontClrPr[0][2] + ";";//" 2px 2px;" +
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
                    //     "filter: " + fontClrPr[1].effcts + ";";
                } else if (fontClrType == "pic") {
                    styleText += fontClrPr[0] + ";";
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                } else if (fontClrType == "gradient") {

                    let colorAry = fontClrPr[0].color;
                    let rot = fontClrPr[0].rot;

                    styleText += `background: linear-gradient(${rot}deg,`;
                    for (let i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            styleText += "#" + colorAry[i] + ");";
                        } else {
                            styleText += "#" + colorAry[i] + ", ";
                        }
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

                }
                styleText += "-webkit-background-clip: text;" +
                    "background-clip: text;" +
                    "color: transparent;";
                if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
                    styleText += "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                }
                if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
                    styleText += "filter: " + fontClrPr[1].effcts + ";";
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
                        highlightColor = "#" + highlightColor;
                    }
                    styleText += "background-color:" + highlightColor + ";";
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
                cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
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
                let linkURL = warpObj["slideResObj"][linkID]["target"];
                linkURL = PPTXXmlUtils.escapeHtml(linkURL);
                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + `'><a href='${linkURL}' ` + linkColorSyle + `  ${linkTooltip} target='_blank'>` +
                        text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a>" + closeElemnt;
            } else {
                return openElemnt + ` class='text-block ${cssName}' style='` + text_style + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + closeElemnt;//"</bdi>";
            }

        }


        function genTable(node, warpObj) {
            let order = node["attrs"]["order"];
            let tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            let xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            let getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
            let getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
            let tblDir = "";
            if (getTblPr !== undefined) {
                let isRTL = getTblPr["attrs"]["rtl"];
                tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
            }
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
            let tbleStyleId = getTblPr["a:tableStyleId"];
            if (tbleStyleId !== undefined) {
                let tbleStylList = warpObj.tableStyles["a:tblStyleLst"]["a:tblStyle"];
                if (tbleStylList !== undefined) {
                    if (tbleStylList.constructor === Array) {
                        for (let k = 0; k < tbleStylList.length; k++) {
                            if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                                thisTblStyle = tbleStylList[k];
                            }
                        }
                    } else {
                        if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList;
                        }
                    }
                }
            }
            if (thisTblStyle !== undefined) {
                thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
                warpObj["thisTbiStyle"] = thisTblStyle;
            }
            let tblStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
            let tblBorderStyl = PPTXXmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
            let tbl_borders = "";
            if (tblBorderStyl !== undefined) {
                tbl_borders = PPTXStyleUtils.getTableBorders(tblBorderStyl, warpObj);
            }
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
            if (tbl_bgcolor !== "" && typeof tbl_bgcolor === 'string') {
                if (tbl_bgcolor.length === 8) {
                    let colorObj = tinycolor(tbl_bgcolor);
                    tbl_bgcolor = colorObj.toRgbString();
                } else {
                    tbl_bgcolor = "#" + tbl_bgcolor;
                }
                tbl_bgcolor = `background-color: ${tbl_bgcolor};`;
            }
            ////////////////////////////////////////////////////////////////////////////////////////////
            let tableHtml = `<table ${tblDir} style='border-collapse: collapse;` +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) +
                PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                ` z-index: ${order};` +
                tbl_borders + `;${tbl_bgcolor}'>`;

            let trNodes = tableNode["a:tr"];
            if (trNodes.constructor !== Array) {
                trNodes = [trNodes];
            }
            //if (trNodes.constructor === Array) {
                //multi rows
                let totalrowSpan = 0;
                let rowSpanAry = [];
                for (let i = 0; i < trNodes.length; i++) {
                    //////////////rows Style ////////////Amir
                    let rowHeightParam = trNodes[i]["attrs"]["h"];
                    let rowHeight = 0;
                    let rowsStyl = "";
                    if (rowHeightParam !== undefined) {
                        rowHeight = parseInt(rowHeightParam) * SLIDE_FACTOR;
                        rowsStyl += `height:${rowHeight}px;`;
                    }
                    let fillColor = "";
                    let row_borders = "";
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
                    rowsStyl += ((row_borders !== undefined) ? row_borders : "");
                    if (fontClrPr !== undefined && typeof fontClrPr === 'string') {
                        let tableColorValue = fontClrPr;
                        if (tableColorValue.length === 8) {
                            let colorObj = tinycolor(tableColorValue);
                            tableColorValue = colorObj.toRgbaString();
                        } else {
                            tableColorValue = "#" + tableColorValue;
                        }
                        rowsStyl += ` color: ${tableColorValue};`;
                    }
                    rowsStyl += ((fontWeight != "") ? ` font-weight:${fontWeight};` : "");
                    if (fillColor !== undefined && fillColor != "" && typeof fillColor === 'string') {
                        if (fillColor.length === 8) {
                            let colorObj = tinycolor(fillColor);
                            fillColor = colorObj.toRgbString();
                        } else {
                            fillColor = "#" + fillColor;
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

                                    let cellParmAry = getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj)
                                    let text = cellParmAry[0];
                                    let colStyl = cellParmAry[1];
                                    let cssName = cellParmAry[2];
                                    let rowSpan = cellParmAry[3];
                                    let colSpan = cellParmAry[4];



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


                            let cellParmAry = getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj)
                            let text = cellParmAry[0];
                            let colStyl = cellParmAry[1];
                            let cssName = cellParmAry[2];
                            let rowSpan = cellParmAry[3];

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
        
        function getTableCellParams(tcNodes, getColsGrid , row_idx , col_idx , thisTblStyle, cellSource, warpObj) {
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
            let lin_bottm = "",
                lin_top = "",
                lin_left = "",
                lin_right = "",
                lin_bottom_left_to_top_right = "",
                lin_top_left_to_bottom_right = "";
            
            let colSapnInt = parseInt(colSpan);
            let total_col_width = 0;
            if (!isNaN(colSapnInt) && colSapnInt > 1){
                for (let k = 0; k < colSapnInt ; k++) {
                    total_col_width += parseInt (PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
                }
            }else{
                total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
            }
            

            let text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width);//tableStyles

            if (total_col_width != 0 /*&& row_idx == 0*/) {
                colWidth = parseInt(total_col_width) * SLIDE_FACTOR;
                colStyl += `width:${colWidth}px;`;
            }

            //cell bords
            lin_bottm = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
            if (lin_bottm === undefined && cellSource !== undefined) {
                if (cellSource !== undefined)
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                if (lin_bottm === undefined) {
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                }
            }
            lin_top = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
            if (lin_top === undefined) {
                if (cellSource !== undefined)
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                if (lin_top === undefined) {
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                }
            }
            lin_left = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
            if (lin_left === undefined) {
                if (cellSource !== undefined)
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                if (lin_left === undefined) {
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                }
            }
            lin_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
            if (lin_right === undefined) {
                if (cellSource !== undefined)
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                if (lin_right === undefined) {
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                }
            }
            lin_bottom_left_to_top_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
            lin_top_left_to_bottom_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

            if (lin_bottm !== undefined && lin_bottm != "") {
                let bottom_line_border = PPTXStyleUtils.getBorder(lin_bottm, undefined, false, "", warpObj)
                if (bottom_line_border != "") {
                    colStyl += `border-bottom:${bottom_line_border};`;
                }
            }
            if (lin_top !== undefined && lin_top != "") {
                let top_line_border = PPTXStyleUtils.getBorder(lin_top, undefined, false, "", warpObj);
                if (top_line_border != "") {
                    colStyl += `border-top: ${top_line_border};`;
                }
            }
            if (lin_left !== undefined && lin_left != "") {
                let left_line_border = PPTXStyleUtils.getBorder(lin_left, undefined, false, "", warpObj)
                if (left_line_border != "") {
                    colStyl += `border-left: ${left_line_border};`;
                }
            }
            if (lin_right !== undefined && lin_right != "") {
                let right_line_border = PPTXStyleUtils.getBorder(lin_right, undefined, false, "", warpObj)
                if (right_line_border != "") {
                    colStyl += `border-right:${right_line_border};`;
                }
            }

            //cell fill color custom
            let getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
            if (getCelFill !== undefined && getCelFill != "") {
                let cellObj = {
                    "p:spPr": getCelFill
                };
                celFillColor = PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide")
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
                    cssName = "_tbl_cell_css_" + (Object.keys(warpObj.styleTable).length + 1);
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

            return [text, colStyl, cssName, rowSpan, colSpan];
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
