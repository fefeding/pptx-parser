
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
    static genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate, styleTable) {
    // 需要的依赖变量: rtl_langs_array, styleTable, is_first_br
    // 这些变量需要通过参数传递或从模块中获取
    var text_style = "";
    var lstStyle = textBodyNode["a:lstStyle"];
    var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

    var text = node["a:t"];

    var openElemnt = "<span";
    var closeElemnt = "</span>";
    var styleText = "";
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

    var pPrNode = pNode["a:pPr"];
    var lvl = 1;
    var lvlNode = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }

    var layoutMasterNode = PPTXLayoutUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
    var pPrNodeLaout = layoutMasterNode.nodeLaout;
    var pPrNodeMaster = layoutMasterNode.nodeMaster;

    // Language check
    var lang = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
    var rtlLangs = PPTXConstants.RTL_LANGS;
    var isRtlLan = (lang !== undefined && rtlLangs.indexOf(lang) !== -1) ? true : false;

    // RTL
    var getRtlVal = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    var isRTL = false;
    if (getRtlVal !== undefined && getRtlVal == "1") {
        isRTL = true;
    }

    var linkID = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
    var linkTooltip = "";
    var defLinkClr;
    if (linkID !== undefined) {
        linkTooltip = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
        if (linkTooltip !== undefined) {
            linkTooltip = "title='" + linkTooltip + "'";
        }
        defLinkClr = PPTXColorUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

        var linkClrNode = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
        var rPrlinkClr = PPTXColorUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);

        if (rPrlinkClr !== undefined && rPrlinkClr != "") {
            defLinkClr = rPrlinkClr;
        }
    }

    // Get font color
    var fontClrPr = PPTXTextStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
    var fontClrType = fontClrPr[2];

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
            var colorAry = fontClrPr[0].color;
            var rot = fontClrPr[0].rot;

            styleText += "background: linear-gradient(" + rot + "deg,";
            for (var i = 0; i < colorAry.length; i++) {
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

    var font_size = PPTXTextStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
    
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
    var highlight = PPTXUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
    if (highlight !== undefined) {
        styleText += "background-color:#" + PPTXColorUtils.getSolidFill(highlight, undefined, undefined, warpObj) + ";";
    }

    // Letter spacing
    var spcNode = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
    if (spcNode === undefined) {
        spcNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
        }
    }
    if (spcNode !== undefined) {
        var ltrSpc = parseInt(spcNode) / 100;
        styleText += "letter-spacing: " + ltrSpc + "px;";
    }

    // Text Cap Types
    var capNode = PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
    if (capNode === undefined) {
        capNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
        }
    }
    if (capNode == "small" || capNode == "all") {
        styleText += "text-transform: uppercase";
    }

    var cssName = "";
    
    if (styleText in styleTable) {
        cssName = styleTable[styleText]["name"];
    } else {
        cssName = "_css_" + (Object.keys(styleTable).length + 1);
        styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }

    var linkColorSyle = "";
    if (fontClrType == "solid" && linkID !== undefined) {
        linkColorSyle = "style='color: inherit;'";
    }

    if (linkID !== undefined && linkID != "") {
        var linkURL = warpObj["slideResObj"][linkID]["target"];
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
    static genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width, styleTable) {
        let text = "";
        const slideMasterTextStyles = warpObj["slideMasterTextStyles"];

    if (textBodyNode === undefined) {
        return text;
    }
    //rtl : <p:txBody>
    //          <a:bodyPr wrap="square" rtlCol="1">

    var pFontStyle = PPTXUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
    //console.log("genTextBody spNode: ", PPTXUtils.getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

    //var lstStyle = textBodyNode["a:lstStyle"];
    
    var apNode = textBodyNode["a:p"];
    if (apNode.constructor !== Array) {
        apNode = [apNode];
    }

    for (var i = 0; i < apNode.length; i++) {
        var pNode = apNode[i];
        var rNode = pNode["a:r"];
        var fldNode = pNode["a:fld"];
        var brNode = pNode["a:br"];
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
            brNode.forEach(function (item, indx) {
                item.type = "br";
            });
            if (brNode.length > 1) {
                brNode.shift();
            }
            rNode = rNode.concat(brNode)
            //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
            rNode.sort(function (a, b) {
                return a.attrs.order - b.attrs.order;
            });
            //console.log("sorted rNode:",rNode)
        }
        //rtlStr = "";//"dir='"+isRTL+"'";
        let styleText = "";
        var marginsVer = PPTXTextStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
        if (marginsVer != "") {
            styleText = marginsVer;
        }
        if (type == "body" || type == "obj" || type == "shape") {
            styleText += "font-size: 0px;";
            //styleText += "line-height: 0;";
            styleText += "font-weight: 100;";
            styleText += "font-style: normal;";
        }
        let cssName = "";

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
        var prg_width_node = PPTXUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
        var prg_height_node;// = PPTXUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
        var sld_prg_width = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * PPTXConstants.SLIDE_FACTOR) + "px;") : "width:inherit;");
        var sld_prg_height = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * PPTXConstants.SLIDE_FACTOR) + "px;") : "");
        var prg_dir = PPTXTextStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
        text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + PPTXTextStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + " " +
            prg_dir + " " + cssName + "' >";
        var buText_ary = PPTXBulletUtils.genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
        var isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
        var bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
        text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
        //get text margin 
        var margin_ary = PPTXTextStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj);
        var margin = margin_ary[0];
        var mrgin_val = margin_ary[1];
        if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
            //sorce : table text
            prg_width_node = tbl_col_width;
        }

        var prgrph_text = "";
        //var prgr_txt_art = [];
        var total_text_len = 0;
        if (rNode === undefined && pNode !== undefined) {
            // without r
            var prgr_text = PPTXTextElementUtils.genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate, styleTable);
            if (isBullate) {
                var txt_obj = document.createElement('div');
                txt_obj.innerHTML = prgr_text;
                var span = txt_obj.firstChild;
                span.style.position = 'absolute';
                span.style.float = 'left';
                span.style.whiteSpace = 'nowrap';
                span.style.visibility = 'hidden';
                document.body.appendChild(span);
                total_text_len += span.offsetWidth;
                document.body.removeChild(span);
            }
            prgrph_text += prgr_text;
        } else if (rNode !== undefined) {
            // with multi r
            for (var j = 0; j < rNode.length; j++) {
                const prgr_text = PPTXTextElementUtils.genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate, styleTable);
                if (isBullate) {
                    const txt_obj = document.createElement('div');
                    txt_obj.innerHTML = prgr_text;
                    const span = txt_obj.firstChild;
                    span.style.position = 'absolute';
                    span.style.float = 'left';
                    span.style.whiteSpace = 'nowrap';
                    span.style.visibility = 'hidden';
                    document.body.appendChild(span);
                    total_text_len += span.offsetWidth;
                    document.body.removeChild(span);
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
        var prg_width = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node )) + "px;" : "width:inherit;");
        text += "<div style='height: 100%;direction: initial;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
        text += prgrph_text;
        text += "</div>";
        text += "</div>";
    }

    return text;
};

    // Break line tracking state
    static #isFirstBreak = false;

    static isFirstBreak() {
        return PPTXTextElementUtils.#isFirstBreak;
    }

    static setFirstBreak(value) {
        PPTXTextElementUtils.#isFirstBreak = value;
    }

}

// 为了保持向后兼容性和全局访问，保留全局赋值
// if (typeof window !== 'undefined') {
//     window.PPTXTextElementUtils = PPTXTextElementUtils;
// } // Removed for ES modules

export { PPTXTextElementUtils };
