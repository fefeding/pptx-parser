/**
 * PPTX 文本元素工具
 * 用于生成 span 元素和文本相关内容
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXTextElementUtils = {};

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
    PPTXTextElementUtils.genSpanElement = function(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate, styleTable) {
        // 需要的依赖变量: rtl_langs_array, styleTable, is_first_br
        // 这些变量需要通过参数传递或从模块中获取
        var text_style = "";
        var lstStyle = textBodyNode["a:lstStyle"];
        var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

        var text = node["a:t"];

        var openElemnt = "<sapn";
        var closeElemnt = "</sapn>";
        var styleText = "";
        if (text === undefined && node["type"] !== undefined) {
            if (window.PPTXTextElementUtils.isFirstBreak()) {
                window.PPTXTextElementUtils.setFirstBreak(false);
                return "<sapn class='line-break-br' ></sapn>";
            }

            styleText += "display: block;";
        } else {
            window.PPTXTextElementUtils.setFirstBreak(true);
        }

        if (typeof text !== 'string') {
            text = window.PPTXUtils.getTextByPathList(node, ["a:fld", "a:t"]);
            if (typeof text !== 'string') {
                text = "&nbsp;";
            }
        }

        var pPrNode = pNode["a:pPr"];
        var lvl = 1;
        var lvlNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }

        var layoutMasterNode = window.PPTXLayoutUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
        var pPrNodeLaout = layoutMasterNode.nodeLaout;
        var pPrNodeMaster = layoutMasterNode.nodeMaster;

        // Language check
        var lang = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
        var rtlLangs = window.PPTXConstants.RTL_LANGS;
        var isRtlLan = (lang !== undefined && rtlLangs.indexOf(lang) !== -1) ? true : false;

        // RTL
        var getRtlVal = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        var isRTL = false;
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
        }

        var linkID = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
        var linkTooltip = "";
        var defLinkClr;
        if (linkID !== undefined) {
            linkTooltip = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
            if (linkTooltip !== undefined) {
                linkTooltip = "title='" + linkTooltip + "'";
            }
            defLinkClr = window.PPTXColorUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

            var linkClrNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
            var rPrlinkClr = window.PPTXColorUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);

            if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                defLinkClr = rPrlinkClr;
            }
        }

        // Get font color
        var fontClrPr = window.PPTXTextStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
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

        var font_size = window.PPTXTextStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
        
        text_style += "font-size:" + font_size + ";" +
            "font-family:" + window.PPTXTextStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
            "font-weight:" + window.PPTXTextStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
            "font-style:" + window.PPTXTextStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
            "text-decoration:" + window.PPTXTextStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
            "text-align:" + window.PPTXTextStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
            "vertical-align:" + window.PPTXTextStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";

        // RTL language direction
        if (isRtlLan) {
            styleText += "direction:rtl;";
        } else {
            styleText += "direction:ltr;";
        }

        // Highlight
        var highlight = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
        if (highlight !== undefined) {
            styleText += "background-color:#" + window.PPTXColorUtils.getSolidFill(highlight, undefined, undefined, warpObj) + ";";
        }

        // Letter spacing
        var spcNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
        if (spcNode === undefined) {
            spcNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
            }
        }
        if (spcNode !== undefined) {
            var ltrSpc = parseInt(spcNode) / 100;
            styleText += "letter-spacing: " + ltrSpc + "px;";
        }

        // Text Cap Types
        var capNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
        if (capNode === undefined) {
            capNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
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
            linkURL = window.PPTXUtils.escapeHtml(linkURL);
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + " target='_blank'>" +
                    text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a>" + closeElemnt;
        } else {
            return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + closeElemnt;
        }
    };

    // Break line tracking state
    var _isFirstBreak = false;

    PPTXTextElementUtils.isFirstBreak = function() {
        return _isFirstBreak;
    };

    PPTXTextElementUtils.setFirstBreak = function(value) {
        _isFirstBreak = value;
    };

    // Export to window
    window.PPTXTextElementUtils = PPTXTextElementUtils;

})();
