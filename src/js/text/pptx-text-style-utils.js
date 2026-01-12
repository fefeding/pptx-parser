/**
 * pptx-text-style-utils.js
 * Utilities for handling text styles, fonts, and alignment
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXTextStyleUtils = {};

    /**
     * Get font bold style
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} "bold" or "inherit"
     */
    PPTXTextStyleUtils.getFontBold = function(node, type, slideMasterTextStyles) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "inherit";
    };

    /**
     * Get font italic style
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} "italic" or "inherit"
     */
    PPTXTextStyleUtils.getFontItalic = function(node, type, slideMasterTextStyles) {
        return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
    };

    /**
     * Get font decoration (underline, strikethrough)
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} CSS text-decoration value
     */
    PPTXTextStyleUtils.getFontDecoration = function(node, type, slideMasterTextStyles) {
        if (node["a:rPr"] !== undefined) {
            var underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
            var strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';

            if (underLine != "none" && strikethrough == "noStrike") {
                return "underline";
            } else if (underLine == "none" && strikethrough != "noStrike") {
                return "line-through";
            } else if (underLine != "none" && strikethrough != "noStrike") {
                return "underline line-through";
            } else {
                return "inherit";
            }
        } else {
            return "inherit";
        }
    };

    /**
     * Get text vertical align (baseline)
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} CSS vertical-align value
     */
    PPTXTextStyleUtils.getTextVerticalAlign = function(node, type, slideMasterTextStyles) {
        var baseline = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
        return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
    };

    /**
     * Get font typeface
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} warpObj - The warp object
     * @param {Object} pFontStyle - Parent font style
     * @returns {String} Font family name
     */
    PPTXTextStyleUtils.getFontType = function(node, type, warpObj, pFontStyle) {
        var typeface = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

        if (typeface === undefined) {
            var fontIdx = "";
            var fontGrup = "";
            if (pFontStyle !== undefined) {
                fontIdx = window.PPTXUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
            }
            var fontSchemeNode = window.PPTXUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
            if (fontIdx == "") {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    fontIdx = "major";
                } else {
                    fontIdx = "minor";
                }
            }
            fontGrup = "a:" + fontIdx + "Font";
            typeface = window.PPTXUtils.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
        }

        return (typeface === undefined) ? "inherit" : typeface;
    };

    /**
     * Get text horizontal align
     * @param {Object} node - The paragraph node
     * @param {Object} pNode - The parent node
     * @param {String} type - The text type
     * @param {Object} warpObj - The warp object
     * @returns {String} CSS text-align value
     */
    PPTXTextStyleUtils.getTextHorizontalAlign = function(node, pNode, type, warpObj) {
        var getAlgn = window.PPTXUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
        if (getAlgn === undefined) {
            getAlgn = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
        }
        if (getAlgn === undefined) {
            if (type == "title" || type == "ctrTitle" || type == "subTitle") {
                var lvlIdx = 1;
                var lvlNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
                if (lvlNode !== undefined) {
                    lvlIdx = parseInt(lvlNode) + 1;
                }
                var lvlStr = "a:lvl" + lvlIdx + "pPr";
                getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (getAlgn === undefined) {
                    getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined) {
                        getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                        if (getAlgn === undefined && type === "subTitle") {
                            getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                        }
                    }
                }
            } else if (type == "body") {
                getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
            } else {
                getAlgn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
            }

        }

        var align = "inherit";
        if (getAlgn !== undefined) {
            switch (getAlgn) {
                case "l":
                    align = "left";
                    break;
                case "r":
                    align = "right";
                    break;
                case "ctr":
                    align = "center";
                    break;
                case "just":
                    align = "justify";
                    break;
                case "dist":
                    align = "justify";
                    break;
                default:
                    align = "inherit";
            }
        }
        return align;
    };

    // Export to window
    window.PPTXTextStyleUtils = PPTXTextStyleUtils;

})();
