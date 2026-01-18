import { PPTXColorUtils } from "./color.js";
import { PPTXUtils } from './utils.js';
var StyleManager = function() {
    this.styleTable = {};
};

    /**
 * 获取或创建样式对应的 CSS 类名
 * @param {String} styleText - 样式文本
 * @param {String} prefix - CSS 类名前缀 (可选)
 * @returns {String} CSS 类名
 */
StyleManager.prototype.getStyleClassName = function(styleText, prefix) {
    var cssName;
    if (styleText in this.styleTable) {
        cssName = this.styleTable[styleText]["name"];
    } else {
        prefix = prefix || "_css_";
        cssName = prefix + (Object.keys(this.styleTable).length + 1);
        this.styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }
    return cssName;
};

    /**
 * 获取样式表
 * @returns {Object} 样式表对象
 */
StyleManager.prototype.getStyleTable = function() {
    return this.styleTable;
};

    /**
 * 生成全局 CSS 样式
 * @returns {String} CSS 样式字符串
 */
StyleManager.prototype.generateGlobalCSS = function() {
    var css = "";
    for (var styleText in this.styleTable) {
        var cssName = this.styleTable[styleText]["name"];
        css += "." + cssName + " {" + styleText + "}\n";
    }
    return css;
};

    /**
 * 重置样式表
 */
StyleManager.prototype.reset = function() {
    this.styleTable = {};
};


    // 获取边框样式
StyleManager.prototype.getBorder = function(node, pNode, isSvgMode, bType, warpObj) {
    var cssText, lineNode;

    if (bType == "shape") {
        cssText = "border: ";
        lineNode = node["p:spPr"]["a:ln"];
    } else if (bType == "text") {
        cssText = "";
        lineNode = node["a:rPr"]["a:ln"];
    }

    var is_noFill = PPTXUtils.getTextByPathList(lineNode, ["a:noFill"]);
    if (is_noFill !== undefined) {
        return "hidden";
    }

    if (lineNode == undefined) {
        var lnRefNode = PPTXUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined){
            var lnIdx = PPTXUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
            lineNode = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
        }
    }
    if (lineNode == undefined) {
        cssText = "";
        lineNode = node;
    }

    var borderColor;
    if (lineNode !== undefined) {
        var borderWidth = parseInt(PPTXUtils.getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
        if (isNaN(borderWidth) || borderWidth < 1) {
            cssText += (4/3) + "px ";
        } else {
            cssText += borderWidth + "px ";
        }

        var borderType = PPTXUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
        if (borderType === undefined) {
            borderType = PPTXUtils.getTextByPathList(lineNode, ["attrs", "cmpd"]);
        }
        var strokeDasharray = "0";
        switch (borderType) {
            case "solid":
                cssText += "solid";
                strokeDasharray = "0";
                break;
            case "dash":
                cssText += "dashed";
                strokeDasharray = "5";
                break;
            case "dashDot":
                cssText += "dashed";
                strokeDasharray = "5, 5, 1, 5";
                break;
            case "dot":
                cssText += "dotted";
                strokeDasharray = "1, 5";
                break;
            case "lgDash":
                cssText += "dashed";
                strokeDasharray = "10, 5";
                break;
            case "dbl":
                cssText += "double";
                strokeDasharray = "0";
                break;
            case "lgDashDotDot":
                cssText += "dashed";
                strokeDasharray = "10, 5, 1, 5, 1, 5";
                break;
            case "sysDash":
                cssText += "dashed";
                strokeDasharray = "5, 2";
                break;
            case "sysDashDot":
                cssText += "dashed";
                strokeDasharray = "5, 2, 1, 5";
                break;
            case "sysDashDotDot":
                cssText += "dashed";
                strokeDasharray = "5, 2, 1, 5, 1, 5";
                break;
            case "sysDot":
                cssText += "dotted";
                strokeDasharray = "2, 5";
                break;
            default:
                cssText += "solid";
                strokeDasharray = "0";
        }

        var fillTyp = PPTXColorUtils.getFillType(lineNode);
        if (fillTyp == "NO_FILL") {
            borderColor = isSvgMode ? "none" : "";
        } else if (fillTyp == "SOLID_FILL") {
            borderColor = PPTXColorUtils.getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
        } else if (fillTyp == "GRADIENT_FILL") {
            borderColor = PPTXColorUtils.getGradientFill(lineNode["a:gradFill"], warpObj);
        } else if (fillTyp == "PATTERN_FILL") {
            borderColor = PPTXColorUtils.getPatternFill(lineNode["a:pattFill"], warpObj);
        }
    }

    if (borderColor === undefined) {
lnRefNode = PPTXUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined) {
            borderColor = PPTXColorUtils.getSolidFill(lnRefNode, undefined, undefined, warpObj);
        }
    }

    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        } else {
            borderColor = "hidden";
        }
    } else {
        borderColor = "#" + borderColor;
    }
    cssText += " " + borderColor + " ";

    if (isSvgMode) {
        return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
    } else {
        return cssText + ";";
    }
}

    // 单例模式
var PPTXStyleManager = new StyleManager();

export { 
    PPTXStyleManager
 };

// Also export to global scope for backward compatibility
// window.PPTXStyleManager = PPTXStyleManager; // Removed for ES modules
