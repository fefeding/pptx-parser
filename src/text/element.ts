import { PPTXUtils } from '../core/utils.js';
import { PPTXColorUtils } from '../core/color.js';
import { PPTXLayoutUtils } from '../core/layout.js';
import { PPTXBulletUtils } from './bullet.js';
import { PPTXConstants } from '../core/constants.js';
import { PPTXTextStyleUtils } from './style.js';

class PPTXTextElementUtils {
    /**
     * 生成文本 span 元素
     * @param {Object} node - 文本节点
     * @param {Number} rIndex - 运行索引
     * @param {Object} pNode - 段落节点
     * @param {Object} textBodyNode - 文本主体节点
     * @param {Object} pFontStyle - 段落字体样式
     * @param {Object} slideLayoutSpNode - 幻灯片布局形状节点
     * @param {Number} idx - 索引
     * @param {String} type - 类型
     * @param {Number} rNodeLength - 运行节点长度
     * @param {Object} warpObj - 包装对象
     * @param {Boolean} isBullate - 是否有项目符号
     * @param {Object} styleTable - 样式表对象
     * @returns {String} HTML span 元素
     */
    static genSpanElement(node: any, rIndex: number, pNode: any, textBodyNode: any, pFontStyle: any, slideLayoutSpNode: any, idx: number, type: string, rNodeLength: number, warpObj: any, isBullate: boolean, styleTable: any): string {
        // 需要的依赖变量: rtl_langs_array, styleTable, is_first_br
        // 这些变量需要通过参数传递或从模块中获取
        let text_style: string = "";
        let lstStyle: any = textBodyNode["a:lstStyle"];
        let slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];

        let text: any = node["a:t"];

        // 检查是否包含数学公式内容
        if (node["a:fld"] && node["a:fld"]["m:f"] !== undefined) {
            // 这可能是数学公式字段
            const mathFormula = PPTXTextElementUtils.processMathFormula(node["a:fld"]);
            if (mathFormula) {
                return mathFormula;
            }
        }
        
        // 检查是否直接包含OMath内容
        if (node["m:oMath"] !== undefined) {
            const mathContent = PPTXTextElementUtils.processOMath(node["m:oMath"]);
            if (mathContent) {
                return mathContent;
            }
        }
        
        let openElemnt: string = "<span";
        let closeElemnt: string = "</span>";
        let styleText: string = "";
        if (text === undefined && node["type"] !== undefined) {
            if (PPTXTextElementUtils.isFirstBreak()) {
                PPTXTextElementUtils.setFirstBreak(false);
                return "<span class='line-break-br' ></span>";
            }

            styleText += "display: block;";
        } else {
            PPTXTextElementUtils.setFirstBreak(true);
        }

        if (typeof text !== 'string') {
            text = PPTXUtils.getTextByPathList(node, ["a:fld", "a:t"]);
            if (typeof text !== 'string') {
                text = "&nbsp;";
            }
        }

        let pPrNode: any = pNode["a:pPr"];
        let lvl: number = 1;
        let lvlNode: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }

        let layoutMasterNode: any = PPTXLayoutUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
        let pPrNodeLaout: any = layoutMasterNode.nodeLaout;
        let pPrNodeMaster: any = layoutMasterNode.nodeMaster;

        // Language check
        let lang: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
        let rtlLangs: any = PPTXConstants.RTL_LANGS;
        let isRtlLan: boolean = (lang !== undefined && rtlLangs.indexOf(lang) !== -1) ? true : false;

        // RTL
        let getRtlVal: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        let isRTL: boolean = false;
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
        }

        let linkID: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
        let linkTooltip: string = "";
        let defLinkClr: any;
        if (linkID !== undefined) {
            linkTooltip = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
            if (linkTooltip !== undefined) {
                linkTooltip = "title='" + linkTooltip + "'";
            }
            defLinkClr = PPTXColorUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

            let linkClrNode: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
            let rPrlinkClr: any = PPTXColorUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);

            if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                defLinkClr = rPrlinkClr;
            }
        }

        // Get font color
        let fontClrPr: any = PPTXTextStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
        let fontClrType: string = fontClrPr[2];

        if (fontClrType == "solid") {
            if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                styleText += "color: #" + fontClrPr[0] + ";";
            } else if (linkID !== undefined && defLinkClr !== undefined) {
                styleText += "color: #" + defLinkClr + ";";
            }

            if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                styleText += "text-shadow:" + fontClrPr[1] + ";";
            }
            if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                styleText += "background-color: #" + fontClrPr[3] + ";";
            }
        } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
            if (fontClrType == "pattern") {
                styleText += "background:" + fontClrPr[0][0] + ";";
                if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                    styleText += "background-size:" + fontClrPr[0][1] + ";";
                }
                if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                    styleText += "background-position:" + fontClrPr[0][2] + ";";
                }
            } else if (fontClrType == "pic") {
                styleText += fontClrPr[0] + ";";
            } else if (fontClrType == "gradient") {
                let colorAry: any[] = fontClrPr[0].color;
                let rot: number = fontClrPr[0].rot;

                styleText += "background: linear-gradient(" + rot + "deg,";
                for (let i: number = 0; i < colorAry.length; i++) {
                    if (i == colorAry.length - 1) {
                        styleText += "#" + colorAry[i] + ");";
                    } else {
                        styleText += "#" + colorAry[i] + ", ";
                    }
                }
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

        let font_size: string = PPTXTextStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
        
        text_style += "font-size:" + font_size + ";" +
            "font-family:" + PPTXTextStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
            "font-weight:" + PPTXTextStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
            "font-style:" + PPTXTextStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
            "text-decoration:" + PPTXTextStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
            "text-align:" + PPTXTextStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
            "vertical-align:" + PPTXTextStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";

        // RTL language direction
        if (isRtlLan) {
            styleText += "direction:rtl;";
        } else {
            styleText += "direction:ltr;";
        }

        // Highlight
        let highlight: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
        if (highlight !== undefined) {
            styleText += "background-color:#" + PPTXColorUtils.getSolidFill(highlight, undefined, undefined, warpObj) + ";";
        }

        // Letter spacing
        let spcNode: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
            }
        }
        if (spcNode !== undefined) {
            let ltrSpc: number = parseInt(spcNode) / 100;
            styleText += "letter-spacing: " + ltrSpc + "px;";
        }

        // Text Cap Types
        let capNode: any = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
            }
        }
        if (capNode == "small" || capNode == "all") {
            styleText += "text-transform: uppercase";
        }

        let cssName: string = "";
        
        if (styleText in styleTable) {
            cssName = styleTable[styleText]["name"];
        } else {
            cssName = "_css_" + (Object.keys(styleTable).length + 1);
            styleTable[styleText] = {
                "name": cssName,
                "text": styleText
            };
        }

        let linkColorSyle: string = "";
        if (fontClrType == "solid" && linkID !== undefined) {
            linkColorSyle = "style='color: inherit;'";
        }

        if (linkID !== undefined && linkID != "") {
            let linkURL: string = warpObj["slideResObj"][linkID]["target"];
            linkURL = PPTXUtils.escapeHtml(linkURL);
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + " target='_blank'>" +
                    text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a>" + closeElemnt;
        } else {
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + closeElemnt;
        }
    };

    /**
     * 生成文本主体HTML
     * @param {Object} textBodyNode - 文本主体节点
     * @param {Object} spNode - 形状节点
     * @param {Object} slideLayoutSpNode - 幻灯片布局形状节点
     * @param {Object} slideMasterSpNode - 幻灯片母版形状节点
     * @param {String} type - 类型
     * @param {Number} idx - 索引
     * @param {Object} warpObj - 包装对象
     * @param {Number} tbl_col_width - 表格列宽度
     * @param {Object} styleTable - 样式表对象
     * @returns {String} HTML文本
     */
    static genTextBody(textBodyNode: any, spNode: any, slideLayoutSpNode: any, slideMasterSpNode: any, type: string, idx: number, warpObj: any, tbl_col_width: number, styleTable: any): string {
        let text: string = "";
        const slideMasterTextStyles: any = warpObj["slideMasterTextStyles"];

        if (textBodyNode === undefined) {
            return text;
        }
        //rtl : <p:txBody>
        //          <a:bodyPr wrap="square" rtlCol="1">

        let pFontStyle: any = PPTXUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
        //console.log("genTextBody spNode: ", PPTXUtils.getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

        //var lstStyle = textBodyNode["a:lstStyle"];
        
        let apNode: any = textBodyNode["a:p"];
        if (apNode.constructor !== Array) {
            apNode = [apNode];
        }

        for (let i: number = 0; i < apNode.length; i++) {
            let pNode: any = apNode[i];
            let rNode: any = pNode["a:r"];
            let fldNode: any = pNode["a:fld"];
            let brNode: any = pNode["a:br"];
            if (rNode !== undefined) {
                rNode = (rNode.constructor === Array) ? rNode : [rNode];
            }
            if (rNode !== undefined && fldNode !== undefined) {
                fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                rNode = rNode.concat(fldNode)
            }
            if (rNode !== undefined && brNode !== undefined) {
                PPTXTextElementUtils.setFirstBreak(true);
                brNode = (brNode.constructor === Array) ? brNode : [brNode];
                brNode.forEach(function (item: any, indx: number) {
                    item.type = "br";
                });
                if (brNode.length > 1) {
                    brNode.shift();
                }
                rNode = rNode.concat(brNode)
                //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
                rNode.sort(function (a: any, b: any) {
                    return a.attrs.order - b.attrs.order;
                });
                //console.log("sorted rNode:",rNode)
            }
            //rtlStr = "";//"dir='"+isRTL+"'";
            let styleText: string = "";
            let marginsVer: string = PPTXTextStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
            if (marginsVer != "") {
                styleText = marginsVer;
            }
            if (type == "body" || type == "obj" || type == "shape") {
                styleText += "font-size: 0px;";
                //styleText += "line-height: 0;";
                styleText += "font-weight: 100;";
                styleText += "font-style: normal;";
            }
            let cssName: string = "";

            if (styleText in styleTable) {
                cssName = styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(styleTable).length + 1);
                styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }
            //console.log("textBodyNode: ", textBodyNode["a:lstStyle"])
            let prg_width_node: any = PPTXUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
            let prg_height_node: any;// = PPTXUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
            let sld_prg_width: string = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * PPTXConstants.SLIDE_FACTOR) + "px;") : "width:inherit;");
            let sld_prg_height: string = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * PPTXConstants.SLIDE_FACTOR) + "px;") : "");
            let prg_dir: string = PPTXTextStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
            text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + PPTXTextStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + " " +
                prg_dir + " " + cssName + "' >";
            let buText_ary: any = PPTXBulletUtils.genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
            let isBullate: boolean = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
            let bu_width: number = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
            text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
            //get text margin 
            let margin_ary: any[] = PPTXTextStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj);
            let margin: string = margin_ary[0];
            let mrgin_val: number = margin_ary[1];
            if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
                //sorce : table text
                prg_width_node = tbl_col_width;
            }

            let prgrph_text: string = "";
            //var prgr_txt_art = [];
            let total_text_len: number = 0;
            if (rNode === undefined && pNode !== undefined) {
                // without r
                let prgr_text: string = PPTXTextElementUtils.genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate, styleTable);
                if (isBullate) {
                    // Measure text width for proper layout calculation
                    // This works in browser environments
                    if (typeof document !== 'undefined' && document.createElement) {
                        const txt_obj = document.createElement('div');
                        txt_obj.innerHTML = prgr_text;
                        const span = txt_obj.firstChild as HTMLElement;
                        if (span && span.style) {
                            span.style.position = 'absolute';
                            span.style.float = 'left';
                            span.style.whiteSpace = 'nowrap';
                            span.style.visibility = 'hidden';
                            document.body.appendChild(span);
                            total_text_len += span.offsetWidth;
                            document.body.removeChild(span);
                        }
                    }
                }
                prgrph_text += prgr_text;
            } else if (rNode !== undefined) {
                // with multi r
                for (let j: number = 0; j < rNode.length; j++) {
                    const prgr_text: string = PPTXTextElementUtils.genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate, styleTable);
                    if (isBullate) {
                        // Measure text width for proper layout calculation
                        // This works in browser environments
                        if (typeof document !== 'undefined' && document.createElement) {
                            const txt_obj = document.createElement('div');
                            txt_obj.innerHTML = prgr_text;
                            const span = txt_obj.firstChild as HTMLElement;
                            if (span && span.style) {
                                span.style.position = 'absolute';
                                span.style.float = 'left';
                                span.style.whiteSpace = 'nowrap';
                                span.style.visibility = 'hidden';
                                document.body.appendChild(span);
                                total_text_len += span.offsetWidth;
                                document.body.removeChild(span);
                            }
                        }
                    }
                    prgrph_text += prgr_text;
                }
            }

            prg_width_node = parseInt(prg_width_node) * PPTXConstants.SLIDE_FACTOR - bu_width - mrgin_val;
            if (isBullate) {
                //get prg_width_node if there is a bulltes
                //console.log("total_text_len: ", total_text_len, "prg_width_node:", prg_width_node)

                if (total_text_len < prg_width_node ){
                    prg_width_node = total_text_len + bu_width;
                }
            }
            let prg_width: string = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node )) + "px;" : "width:inherit;");
            text += "<div style='height: 100%;direction: initial;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
            text += prgrph_text;
            text += "</div>";
            text += "</div>";
        }

        return text;
    };

    // Break line tracking state
    static #isFirstBreak: boolean = false;

    static isFirstBreak(): boolean {
        return PPTXTextElementUtils.#isFirstBreak;
    }

    static setFirstBreak(value: boolean): void {
        PPTXTextElementUtils.#isFirstBreak = value;
    }

    /**
     * 处理数学公式字段
     * @param {any} fieldNode - 字段节点
     * @returns {string} HTML表示
     */
    static processMathFormula(fieldNode: any): string {
        // 尝试从字段节点中提取数学内容
        if (fieldNode && fieldNode['m:f']) {
            // 检查是否包含OMath内容
            const oMathContent = fieldNode['m:f'];
            const mathText = PPTXTextElementUtils.extractSimpleMathText(oMathContent);
            if (mathText) {
                return `<span class='math-formula' title='${mathText}' style='font-style: italic; color: #0066cc;'>[数学公式: ${mathText}]</span>`;
            }
        }
        // 如果无法提取具体内容，返回通用占位符
        return "<span class='math-formula-placeholder' style='font-style: italic; color: #0066cc;'>[数学公式]</span>";
    }

    /**
     * 从OMath内容中提取简单文本
     * @param {any} oMathContent - OMath内容节点
     * @returns {string} 提取的文本
     */
    static extractSimpleMathText(oMathContent: any): string {
        if (!oMathContent) return '';

        // 递归搜索文本节点
        const findTextNodes = (node: any): string[] => {
            let texts: string[] = [];
            
            if (typeof node === 'object' && node !== null) {
                for (const key in node) {
                    if (key === 'm:t' && typeof node[key] === 'string') {
                        texts.push(node[key]);
                    } else if (typeof node[key] === 'object') {
                        texts = texts.concat(findTextNodes(node[key]));
                    }
                }
            }
            
            return texts;
        };

        const textNodes = findTextNodes(oMathContent);
        return textNodes.join('');
    }

    /**
     * 处理OMath内容
     * @param {any} oMathNode - OMath节点
     * @returns {string} HTML表示
     */
    static processOMath(oMathNode: any): string {
        // 尝试从OMath节点中提取数学内容
        const mathText = PPTXTextElementUtils.extractSimpleMathText(oMathNode);
        if (mathText) {
            return `<span class='omath-placeholder' title='${mathText}' style='font-style: italic; color: #0066cc;'>[OMath公式: ${mathText}]</span>`;
        }
        // 如果无法提取具体内容，返回通用占位符
        return "<span class='omath-placeholder' style='font-style: italic; color: #0066cc;'>[OMath公式]</span>";
    }
}

export { PPTXTextElementUtils };