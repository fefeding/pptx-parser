/**
 * pptx-text-style-utils.js
 * Utilities for handling text styles, fonts, and alignment
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXTextStyleUtils = {};
    var fontSizeFactor = 4 / 3.2;
    var slideFactor = 96 / 914400;

    /**
     * Get font bold style
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} "bold" or "inherit"
     */
    PPTXTextStyleUtils.getFontBold = function(node, type, slideMasterTextStyles) {
        if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] && node["a:rPr"]["attrs"]["b"] === "1") {
            return "bold";
        }
        return "inherit";
    };

    /**
     * Get font italic style
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} "italic" or "inherit"
     */
    PPTXTextStyleUtils.getFontItalic = function(node, type, slideMasterTextStyles) {
        if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] && node["a:rPr"]["attrs"]["i"] === "1") {
            return "italic";
        }
        return "inherit";
    };

    /**
     * Get font decoration (underline, strikethrough)
     * @param {Object} node - The text run node
     * @param {String} type - The text type
     * @param {Object} slideMasterTextStyles - Slide master text styles
     * @returns {String} CSS text-decoration value
     */
    PPTXTextStyleUtils.getFontDecoration = function(node, type, slideMasterTextStyles) {
        if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]) {
            var attrs = node["a:rPr"]["attrs"];
            var underLine = attrs["u"] !== undefined ? attrs["u"] : "none";
            var strikethrough = attrs["strike"] !== undefined ? attrs["strike"] : 'noStrike';

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

    /**
     * Get font size
     * @param {Object} node - The text run node
     * @param {Object} textBodyNode - The text body node
     * @param {Object} pFontStyle - Parent font style
     * @param {Number} lvl - Level
     * @param {String} type - Text type
     * @param {Object} warpObj - Warp object
     * @returns {String} Font size with unit
     */
    PPTXTextStyleUtils.getFontSize = function(node, textBodyNode, pFontStyle, lvl, type, warpObj) {
        // if(type == "sldNum")
        //console.log("getFontSize node:", node, "lstStyle", lstStyle, "lvl:", lvl, 'type:', type, "warpObj:", warpObj)
        var lstStyle = (textBodyNode !== undefined)? textBodyNode["a:lstStyle"] : undefined;
        var lvlpPr = "a:lvl" + lvl + "pPr";
        var fontSize = undefined;
        var sz, kern;
        if (node["a:rPr"] !== undefined) {
            fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
        }
        if (isNaN(fontSize) || fontSize === undefined && node["a:fld"] !== undefined) {
            sz = window.PPTXUtils.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
            sz = window.PPTXUtils.getTextByPathList(node["a:endParaRPr"], [ "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
            sz = window.PPTXUtils.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
        }
        //a:spAutoFit
        var isAutoFit = false;
        var isKerning = false;
        if (textBodyNode !== undefined){
            var spAutoFitNode = window.PPTXUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
            // if (spAutoFitNode === undefined) {
            //     spAutoFitNode = window.PPTXUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit"]);
            // }
            if (spAutoFitNode !== undefined){
                isAutoFit = true;
                isKerning = true;
            }
        }
        if (isNaN(fontSize) || fontSize === undefined) {
            // if (type == "shape" || type == "textBox") {
            //     type = "body";
            //     lvlpPr = "a:lvl1pPr";
            // }
            sz = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            fontSize = parseInt(sz) / 100;
            kern = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (isKerning && kern !== undefined && !isNaN(fontSize) && (fontSize - parseInt(kern) / 100) > 0){
                fontSize = fontSize - parseInt(kern) / 100;
            }
        }

        if (isNaN(fontSize) || fontSize === undefined) {
            // if (type == "shape" || type == "textBox") {
            //     type = "body";
            //     lvlpPr = "a:lvl1pPr";
            // }
            sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
            kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
            if (sz === undefined) {
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                    sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                }
                else if (type == "shape") {
                    //textBox and shape text does not indent
                    sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    isKerning = false;
                }

                if (sz === undefined) {
                    sz = window.PPTXUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                    kern = (kern === undefined)? window.PPTXUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                    isKerning = false;
                }
                //  else if (type === undefined || type == "shape") {
                //     sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                //     kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                // } 
                // else if (type == "textBox") {
                //     sz = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                //     kern = window.PPTXUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                // }
            } 
            fontSize = parseInt(sz) / 100;
            if (isKerning && kern !== undefined && !isNaN(fontSize) && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100 )) {
                fontSize = fontSize - parseInt(kern) / 100;
                //fontSize =  parseInt(kern) / 100;
            }
        }

        var baseline = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
        if (baseline !== undefined && !isNaN(fontSize)) {
            var baselineVl = parseInt(baseline) / 100000;
            //fontSize -= 10; 
            // fontSize = fontSize * baselineVl;
            fontSize -= baselineVl;
        }

        if (!isNaN(fontSize)){
            var normAutofit = window.PPTXUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
            if (normAutofit !== undefined && normAutofit != 0){
                //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
                fontSize = Math.round(fontSize * (normAutofit / 100000))
            }
        }

        return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * fontSizeFactor + "px");// + "pt");
    };

    /**
     * Get font color with text border and shadow effects
     * @param {Object} node - The text run node
     * @param {Object} pNode - The parent node
     * @param {Object} lstStyle - The list style node
     * @param {Object} pFontStyle - Parent font style
     * @param {Number} lvl - Level
     * @param {Number} idx - Index
     * @param {String} type - Text type
     * @param {Object} warpObj - Warp object
     * @returns {Array} [color, txt_effects, colorType, highlightColor]
     */
    PPTXTextStyleUtils.getFontColorPr = function(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj) {
        //text border using: text-shadow: -1px 0 black, 0 1px black, 1px 0 black, 0 -1px black;
        //{getFontColor(..) return color} -> getFontColorPr(..) return array[color,textBordr/shadow]
        //https://stackoverflow.com/questions/2570972/css-font-border
        //https://www.w3schools.com/cssref/css3_pr_text-shadow.asp
        //themeContent
        //console.log("getFontColorPr>> type:", type, ", node: ", node)
        var rPrNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr"]);
        var filTyp, color, textBordr, colorType = "", highlightColor = "";
        //console.log("getFontColorPr type:", type, ", node: ", node, "pNode:", pNode, "pFontStyle:", pFontStyle)
        if (rPrNode !== undefined) {
            filTyp = window.PPTXColorUtils.getFillType(rPrNode);
            if (filTyp == "SOLID_FILL") {
                var solidFillNode = rPrNode["a:solidFill"];// window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                color = window.PPTXColorUtils.getSolidFill(solidFillNode, undefined, undefined, warpObj);
                var highlightNode = rPrNode["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = window.PPTXColorUtils.getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                var pattFill = rPrNode["a:pattFill"];// window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                color = window.PPTXColorUtils.getPatternFill(pattFill, warpObj);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = window.PPTXColorUtils.getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                var shpFill = rPrNode["a:gradFill"];
                color = window.PPTXColorUtils.getGradientFill(shpFill, warpObj);
                colorType = "gradient";
            } 
        }
        if (color === undefined && window.PPTXUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
            //lstStyle
            var lstStyledefRPr = window.PPTXUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
            filTyp = window.PPTXColorUtils.getFillType(lstStyledefRPr);
            if (filTyp == "SOLID_FILL") {
                var solidFillNode = lstStyledefRPr["a:solidFill"];// window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                color = window.PPTXColorUtils.getSolidFill(solidFillNode, undefined, undefined, warpObj);
                var highlightNode = lstStyledefRPr["a:highlight"];
                if (highlightNode !== undefined) {
                    highlightColor = window.PPTXColorUtils.getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
                colorType = "solid";
            } else if (filTyp == "PATTERN_FILL") {
                var pattFill = lstStyledefRPr["a:pattFill"];// window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                color = window.PPTXColorUtils.getPatternFill(pattFill, warpObj);
                colorType = "pattern";
            } else if (filTyp == "PIC_FILL") {
                color = window.PPTXColorUtils.getPicFill("slideBg", lstStyledefRPr["a:blipFill"], warpObj);
                colorType = "pic";
            } else if (filTyp == "GRADIENT_FILL") {
                var shpFill = lstStyledefRPr["a:gradFill"];
                color = window.PPTXColorUtils.getGradientFill(shpFill, warpObj);
                colorType = "gradient";
            }

        }
        if (color === undefined) {
            var sPstyle = window.PPTXUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
            if (sPstyle !== undefined) {
                color = window.PPTXColorUtils.getSolidFill(sPstyle, undefined, undefined, warpObj);
                if (color !== undefined) {
                    colorType = "solid";
                }
                var highlightNode = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
                if (highlightNode !== undefined) {
                    highlightColor = window.PPTXColorUtils.getSolidFill(highlightNode, undefined, undefined, warpObj);
                }
            }
            if (color === undefined) {
                if (pFontStyle !== undefined) {
                    color = window.PPTXColorUtils.getSolidFill(pFontStyle, undefined, undefined, warpObj);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                }
            }
        }
        //console.log("getFontColorPr node", node, "colorType: ", colorType,"color: ",color)

        if (color === undefined) {

            var layoutMasterNode = window.PPTXLayoutUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            if (pPrNodeLaout !== undefined) {
                var defRpRLaout = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                if (defRpRLaout !== undefined) {
                    color = window.PPTXColorUtils.getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                    var highlightNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                    if (highlightNode !== undefined) {
                        highlightColor = window.PPTXColorUtils.getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                }
            }
            if (color === undefined) {

                if (pPrNodeMaster !== undefined) {
                    var defRprMaster = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                    if (defRprMaster !== undefined) {
                        color = window.PPTXColorUtils.getSolidFill(defRprMaster, undefined, undefined, warpObj);
                        var highlightNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = window.PPTXColorUtils.getSolidFill(highlightNode, undefined, undefined, warpObj);
                        }
                        colorType = "solid";
                    }
                }
            }
        }
        var txtEffects = [];
        var txtEffObj = {}
        //textBordr
        var txtBrdrNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:ln"]);
        var textBordr = "";
        if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
            var txBrd = window.PPTXShapeFillsUtils.getBorder(node, pNode, false, "text", warpObj);
            var txBrdAry = txBrd.split(" ");
            //var brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("pt")))) + "px";
            var brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px")))) + "px";
            var brdClr = txBrdAry[2];
            //var brdTyp = txBrdAry[1]; //not in use
            //console.log("getFontColorPr txBrdAry:", txBrdAry)
            if (colorType == "solid") {
                textBordr = "-" + brdSize + " 0 " + brdClr + ", 0 " + brdSize + " " + brdClr + ", " + brdSize + " 0 " + brdClr + ", 0 -" + brdSize + " " + brdClr;
                // if (oShadowStr != "") {
                //     textBordr += "," + oShadowStr;
                // } else {
                //     textBordr += ";";
                // }
                txtEffects.push(textBordr);
            } else {
                //textBordr = brdSize + " " + brdClr;
                txtEffObj.border = brdSize + " " + brdClr;
            }
        }
        // else {
        //     //if no border but exist/not exist shadow
        //     if (colorType == "solid") {
        //         textBordr = oShadowStr;
        //     } else {
        //         //TODO
        //     }
        // }
        var txtGlowNode = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
        var oGlowStr = "";
        if (txtGlowNode !== undefined) {
            var glowClr = window.PPTXColorUtils.getSolidFill(txtGlowNode, undefined, undefined, warpObj);
            var rad = (txtGlowNode["attrs"]["rad"]) ? (txtGlowNode["attrs"]["rad"] * slideFactor) : 0;
            oGlowStr = "0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr +
                ", 0 0 " + rad + "px #" + glowClr;
            if (colorType == "solid") {
                txtEffects.push(oGlowStr);
            } else {
                // txtEffObj.glow = {
                //     radiuse: rad,
                //     color: glowClr
                // } 
                txtEffects.push(
                    "drop-shadow(0 0 " + rad / 3 + "px #" + glowClr + ") " +
                    "drop-shadow(0 0 " + rad * 2 / 3 + "px #" + glowClr + ") " +
                    "drop-shadow(0 0 " + rad + "px #" + glowClr + ")"
                );
            }
        }
        var txtShadow = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
        var oShadowStr = "";
        if (txtShadow !== undefined) {
            //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
            //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
            //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
            //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

            var shadowClr = window.PPTXColorUtils.getSolidFill(txtShadow, undefined, undefined, warpObj);
            var outerShdwAttrs = txtShadow["attrs"];
            // algn: "bl"
            // dir: "2640000"
            // dist: "38100"
            // rotWithShape: "0/1" - Specifies whether the shadow rotates with the shape if the shape is rotated.
            //blurRad (Blur Radius) - Specifies the blur radius of the shadow.
            //kx (Horizontal Skew) - Specifies the horizontal skew angle.
            //ky (Vertical Skew) - Specifies the vertical skew angle.
            //sx (Horizontal Scaling Factor) - Specifies the horizontal scaling slideFactor; negative scaling causes a flip.
            //sy (Vertical Scaling Factor) - Specifies the vertical scaling slideFactor; negative scaling causes a flip.
            var algn = outerShdwAttrs["algn"];
            var dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
            var dist = parseInt(outerShdwAttrs["dist"]) * slideFactor;//(px) //* (3 / 4); //(pt)
            var rotWithShape = outerShdwAttrs["rotWithShape"];
            var blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * slideFactor + "px") : "";
            var sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
            var sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
            var vx = dist * Math.sin(dir * Math.PI / 180);
            var hx = dist * Math.cos(dir * Math.PI / 180);

            //console.log("getFontColorPr outerShdwAttrs:", outerShdwAttrs, ", shadowClr:", shadowClr, ", algn: ", algn, ",dir: ", dir, ", dist: ", dist, ",rotWithShape: ", rotWithShape, ", color: ", color)

            if (!isNaN(vx) && !isNaN(hx)) {
                oShadowStr = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;// + ";";
                if (colorType == "solid") {
                    txtEffects.push(oShadowStr);
                } else {

                    // txtEffObj.oShadow = {
                    //     hx: hx,
                    //     vx: vx,
                    //     radius: blurRad,
                    //     color: shadowClr
                    // }

                    //txtEffObj.oShadow = hx + "px " + vx + "px " + blurRad + " #" + shadowClr;

                    txtEffects.push("drop-shadow(" + hx + "px " + vx + "px " + blurRad + " #" + shadowClr + ")");
                }
            }
            //console.log("getFontColorPr vx:", vx, ", hx: ", hx, ", sx: ", sx, ", sy: ", sy, ",oShadowStr: ", oShadowStr)
        }
        //console.log("getFontColorPr>>> color:", color)
        // if (color === undefined || color === "FFF") {
        //     color = "#000";
        // } else {
        //     color = "" + color;
        // }
        var text_effcts = "", txt_effects;
        if (colorType == "solid") {
            if (txtEffects.length > 0) {
                text_effcts = txtEffects.join(",");
            }
            txt_effects = text_effcts + ";"
        } else {
            if (txtEffects.length > 0) {
                text_effcts = txtEffects.join(" ");
            }
            txtEffObj.effcts = text_effcts;
            txt_effects = txtEffObj
        }
        //console.log("getFontColorPr txt_effects:", txt_effects)

        //return [color, textBordr, colorType];
        return [color, txt_effects, colorType, highlightColor];
    };

    /**
     * Get content direction (LTR/RTL)
     * @param {Object} node - The node
     * @param {String} type - The text type
     * @param {Object} warpObj - The warp object
     * @returns {String} "content" or "content-rtl"
     */
    PPTXTextStyleUtils.getContentDir = function(node, type, warpObj) {
        var defRtl = window.PPTXUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "a:defPPr", "attrs", "rtl"]);
        if (defRtl !== undefined) {
            if (defRtl == "1"){
                return "content-rtl";
            } else if (defRtl == "0") {
                return "content";
            }
        }
        //var lvl1Rtl = window.PPTXUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "lvl1pPr", "attrs", "rtl"]);
        // if (lvl1Rtl !== undefined) {
        //     if (lvl1Rtl == "1") {
        //         return "content-rtl";
        //     } else if (lvl1Rtl == "0") {
        //         return "content";
        //     }
        // }
        var rtlCol = window.PPTXUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "rtlCol"]);
        if (rtlCol !== undefined) {
            if (rtlCol == "1") {
                return "content-rtl";
            } else if (rtlCol == "0") {
                return "content";
            }
        }
        //console.log("getContentDir node:", node, "rtlCol:", rtlCol)

        if (type === undefined) {
            return "content";
        }
        var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
        var dirLoc = "";

        switch (type) {
            case "title":
            case "ctrTitle":
                dirLoc = "p:titleStyle";
                break;
            case "body":
            case "dt":
            case "ftr":
            case "sldNum":
            case "textBox":
                dirLoc = "p:bodyStyle";
                break;
            case "shape":
                dirLoc = "p:otherStyle";
        }
        if (slideMasterTextStyles !== undefined && dirLoc !== "") {
            var dirVal = window.PPTXUtils.getTextByPathList(slideMasterTextStyles[dirLoc], ["a:lvl1pPr", "attrs", "rtl"]);
            if (dirVal == "1") {
                return "content-rtl";
            }
        } 
        // else {
        //     if (type == "textBox") {
        //         var dirVal = window.PPTXUtils.getTextByPathList(warpObj, ["defaultTextStyle", "a:lvl1pPr", "attrs", "rtl"]);
        //         if (dirVal == "1") {
        //             return "content-rtl";
        //         }
        //     }
        // }
        return "content";
        //console.log("getContentDir() type:", type, "slideMasterTextStyles:", slideMasterTextStyles,"dirNode:",dirVal)
    };

    /**
     * Get vertical alignment for block elements
     * @param {Object} node - The node
     * @param {Object} slideLayoutSpNode - Slide layout shape node
     * @param {Object} slideMasterSpNode - Slide master shape node
     * @param {String} type - The text type
     * @returns {String} "v-mid", "v-down", or "v-up"
     */
    PPTXTextStyleUtils.getVerticalAlign = function(node, slideLayoutSpNode, slideMasterSpNode, type) {
        //X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
        var anchor = window.PPTXUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        //console.log("getVerticalAlign anchor:", anchor, "slideLayoutSpNode: ", slideLayoutSpNode)
        if (anchor === undefined) {
            //console.log("getVerticalAlign type:", type," node:", node, "slideLayoutSpNode:", slideLayoutSpNode, "slideMasterSpNode:", slideMasterSpNode)
            anchor = window.PPTXUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            if (anchor === undefined) {
                anchor = window.PPTXUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    //"If this attribute is omitted, then a value of t, or top is implied."
                    anchor = "t";//window.PPTXUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                }
            }
        }
        //console.log("getVerticalAlign:", node, slideLayoutSpNode, slideMasterSpNode, type, anchor)
        return (anchor === "ctr")?"v-mid" : ((anchor === "b") ? "v-down" : "v-up");
    };

    /**
     * Get paragraph direction (LTR/RTL)
     * @param {Object} node - The node
     * @param {Object} textBodyNode - The text body node
     * @param {Number} idx - Index
     * @param {String} type - The text type
     * @param {Object} warpObj - The warp object
     * @returns {String} "pregraph-ltr", "pregraph-rtl", or "pregraph-inherit"
     */
    PPTXTextStyleUtils.getPregraphDir = function(node, textBodyNode, idx, type, warpObj) {
        var rtl = window.PPTXUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
        //console.log("getPregraphDir node:", node, "textBodyNode", textBodyNode, "rtl:", rtl, "idx", idx, "type", type, "warpObj", warpObj)

        if (rtl === undefined) {
            var layoutMasterNode = window.PPTXLayoutUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;
            rtl = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (rtl === undefined && type != "shape") {
                rtl = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtrl"]);
            }
        }

        if (rtl == "1") {
            return "pregraph-rtl";
        } else if (rtl == "0") {
            return "pregraph-ltr";
        }
        return "pregraph-inherit";

        // var contentDir = getContentDir(type, warpObj);
        // console.log("getPregraphDir node:", node["a:r"], "rtl:", rtl, "idx", idx, "type", type, "contentDir:", contentDir)

        // if (contentDir == "content"){
        //     return "pregraph-ltr";
        // } else if (contentDir == "content-rtl"){ 
        //     return "pregraph-rtl";
        // }
        // return "";
    };

    /**
     * Get horizontal alignment for paragraph
     * @param {Object} node - The node
     * @param {Object} textBodyNode - The text body node
     * @param {Number} idx - Index
     * @param {String} type - The text type
     * @param {String} prg_dir - Paragraph direction
     * @param {Object} warpObj - The warp object
     * @returns {String} Alignment class name
     */
    PPTXTextStyleUtils.getHorizontalAlign = function(node, textBodyNode, idx, type, prg_dir, warpObj) {
        var algn = window.PPTXUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
            //var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
            // var pPrNodeLaout = layoutMasterNode.nodeLaout;
            // var pPrNodeMaster = layoutMasterNode.nodeMaster;
            var lvlIdx = 1;
            var lvlNode = window.PPTXUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvlIdx = parseInt(lvlNode) + 1;
            }
            var lvlStr = "a:lvl" + lvlIdx + "pPr";

            var lstStyle = textBodyNode["a:lstStyle"];
            algn = window.PPTXUtils.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);

            if (algn === undefined && idx !== undefined ) {
                //slidelayout
                algn = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                if (algn === undefined) {
                    algn = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                    if (algn === undefined) {
                        algn = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                    }
                }
            }
            if (algn === undefined) {
                if (type !== undefined) {
                    //slidelayout
                    algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                    if (algn === undefined) {
                        //masterlayout
                        if (type == "title" || type == "ctrTitle") {
                            algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "body" || type == "obj" || type == "subTitle") {
                            algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "shape" || type == "diagram") {
                            algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                        } else if (type == "textBox") {
                            algn = window.PPTXUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                        } else {
                            algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        }
                    }
                } else {
                    algn = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                }
            }
        }

        if (algn === undefined) {
            if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                return "h-mid";
            } else if (type == "sldNum") {
                return "h-right";
            }
        }
        if (algn !== undefined) {
            switch (algn) {
                case "l":
                    if (prg_dir == "pregraph-rtl"){
                        //return "h-right";
                        return "h-left-rtl";
                    }else{
                        return "h-left";
                    }
                    break;
                case "r":
                    if (prg_dir == "pregraph-rtl") {
                        //return "h-left";
                        return "h-right-rtl";
                    }else{
                        return "h-right";
                    }
                    break;
                case "ctr":
                    return "h-mid";
                    break;
                case "just":
                case "dist":
                default:
                    return "h-" + algn;
            }
        }
        //return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
    };

    /**
     * Get paragraph margins for bullet points
     * @param {Object} pNode - The paragraph node
     * @param {Number} idx - Index
     * @param {String} type - The text type
     * @param {Boolean} isBullate - Whether it's a bullet point
     * @param {Object} warpObj - The warp object
     * @returns {Array} [margin style string, margin value]
     */
    PPTXTextStyleUtils.getPregraphMargn = function(pNode, idx, type, isBullate, warpObj){
        if (!isBullate){
            return ["",0];
        }
        var marLStr = "", marRStr = "" , maginVal = 0;
        var pPrNode = pNode["a:pPr"];
        var layoutMasterNode = window.PPTXLayoutUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
        var pPrNodeLaout = layoutMasterNode.nodeLaout;
        var pPrNodeMaster = layoutMasterNode.nodeMaster;
        
        // var lang = window.PPTXUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
        // var isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1) ? true : false;
        //rtl
        var getRtlVal = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
        if (getRtlVal === undefined) {
            getRtlVal = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
            if (getRtlVal === undefined && type != "shape") {
                getRtlVal = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
            }
        }
        var isRTL = false;
        var dirStr = "ltr";
        if (getRtlVal !== undefined && getRtlVal == "1") {
            isRTL = true;
            dirStr = "rtl";
        }

        //align
        var alignNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
        if (alignNode === undefined) {
            alignNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
            if (alignNode === undefined) {
                alignNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
            }
        }
        //indent?
        var indentNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
            }
        }
        var indent = 0;
        if (indentNode !== undefined) {
            indent = parseInt(indentNode) * slideFactor;
        }
        //
        //marL
        var marLNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
            }
        }
        var marginLeft = 0;
        if (marLNode !== undefined) {
            marginLeft = parseInt(marLNode) * slideFactor;
        }
        if ((indentNode !== undefined || marLNode !== undefined)) {
            //var lvlIndent = defTabSz * lvl;

            if (isRTL) {// && alignNode == "r") {
                //marLStr = "margin-right: ";
                marLStr = "padding-right: ";
            } else {
                //marLStr = "margin-left: ";
                marLStr = "padding-left: ";
            }
            if (isBullate) {
                maginVal = Math.abs(0 - indent);
                marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
            } else {
                maginVal = Math.abs(marginLeft + indent);
                marLStr += maginVal + "px;";  // (minus bullate numeric lenght/size - TODO
            }
        }

        //marR?
        var marRNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
        if (marRNode === undefined && marLNode === undefined) {
            //need to check if this posble - TODO
            marRNode = window.PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
            if (marRNode === undefined) {
                marRNode = window.PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
            }
        }
        if (marRNode !== undefined && isBullate) {
            var marginRight = parseInt(marRNode) * slideFactor;
            if (isRTL) {// && alignNode == "r") {
                //marRStr = "margin-right: ";
                marRStr = "padding-right: ";
            } else {
                //marRStr = "margin-left: ";
                marRStr = "padding-left: ";
            }
            marRStr += Math.abs(0 - indent) + "px;";
        }


        return [marLStr, maginVal];
    };

    /**
     * Get vertical margins for paragraphs
     * @param {Object} pNode - The paragraph node
     * @param {Object} textBodyNode - The text body node
     * @param {String} type - The text type
     * @param {Number} idx - Index
     * @param {Object} warpObj - The warp object
     * @returns {String} CSS margin and padding styles
     */
    PPTXTextStyleUtils.getVerticalMargins = function(pNode, textBodyNode, type, idx, warpObj) {
        //margin-top ; 
        //a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
        //margin-bottom
        //a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
        //+
        //a:pPr =>a:lnSpc => a:spcPts (/?) | a:spcPct (/?)
        //console.log("getVerticalMargins ", pNode, type,idx, warpObj)
        //var lstStyle = textBodyNode["a:lstStyle"];
        var lvl = 1
        var spcBefNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
        var spcAftNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
        var lnSpcNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
        var lnSpcNodeType = "Pct";
        if (lnSpcNode === undefined) {
            lnSpcNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
            if (lnSpcNode !== undefined) {
                lnSpcNodeType = "Pts";
            }
        }
        var lvlNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        var fontSize;
        if (window.PPTXUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
            var fontSizeStr = window.PPTXTextStyleUtils.getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
            if (fontSizeStr != "inherit") {
                fontSize = parseInt(fontSizeStr, "px"); //pt
            }
        }
        //var spcBef = "";
        //console.log("getVerticalMargins 1", fontSizeStr, fontSize, lnSpcNode, parseInt(lnSpcNode) / 100000, spcBefNode, spcAftNode)
        // if(spcBefNode !== undefined){
        //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
        // }
        // else{
        //    //i did not found case with percentage 
        //     spcBefNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct","attrs","val"]);
        //     if(spcBefNode !== undefined){
        //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
        //     }
        // }
        //var spcAft = "";
        // if(spcAftNode !== undefined){
        //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
        // }
        // else{
        //    //i did not found case with percentage 
        //     spcAftNode = window.PPTXUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct","attrs","val"]);
        //     if(spcAftNode !== undefined){
        //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
        //     }
        // }
        // if(spcAftNode !== undefined){
        //     //check in layout and then in master
        // }
        var isInLayoutOrMaster = true;
        if(type == "shape" || type == "textBox"){
            isInLayoutOrMaster = false;
        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            //check in layout
            if (idx !== undefined) {
                var laypPrNode = window.PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

                if (spcBefNode === undefined) {
                    spcBefNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                    // if(spcBefNode !== undefined){
                    //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                    // } 
                    // else{
                    //    //i did not found case with percentage 
                    //     spcBefNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                    //     if(spcBefNode !== undefined){
                    //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                    //     }
                    // }
                }

                if (spcAftNode === undefined) {
                    spcAftNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                    // if(spcAftNode !== undefined){
                    //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                    // }
                    // else{
                    //    //i did not found case with percentage 
                    //     spcAftNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                    //     if(spcAftNode !== undefined){
                    //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                    //     }
                    // }
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = window.PPTXUtils.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }

        }
        if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
            //check in master
            //slideMasterTextStyles
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
            var dirLoc = "";
            var lvl = "a:lvl" + lvl + "pPr";
            switch (type) {
                case "title":
                case "ctrTitle":
                    dirLoc = "p:titleStyle";
                    break;
                case "body":
                case "obj":
                case "dt":
                case "ftr":
                case "sldNum":
                case "textBox":
                // case "shape":
                    dirLoc = "p:bodyStyle";
                    break;
                case "shape":
                //case "textBox":
                default:
                    dirLoc = "p:otherStyle";
            }
            // if (type == "shape" || type == "textBox") {
            //     lvl = "a:lvl1pPr";
            // }
            var inLvlNode = window.PPTXUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
            if (inLvlNode !== undefined) {
                if (spcBefNode === undefined) {
                    spcBefNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                    // if(spcBefNode !== undefined){
                    //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                    // } 
                    // else{
                    //    //i did not found case with percentage 
                    //     spcBefNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                    //     if(spcBefNode !== undefined){
                    //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                    //     }
                    // }
                }

                if (spcAftNode === undefined) {
                    spcAftNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                    // if(spcAftNode !== undefined){
                    //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                    // }
                    // else{
                    //    //i did not found case with percentage 
                    //     spcAftNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                    //     if(spcAftNode !== undefined){
                    //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                    //     }
                    // }
                }

                if (lnSpcNode === undefined) {
                    lnSpcNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                    if (lnSpcNode === undefined) {
                        lnSpcNode = window.PPTXUtils.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                        if (lnSpcNode !== undefined) {
                            lnSpcNodeType = "Pts";
                        }
                    }
                }
            }
        }
        var spcBefor = 0, spcAfter = 0, spcLines = 0;
        var marginTopBottomStr = "";
        if (spcBefNode !== undefined) {
            spcBefor = parseInt(spcBefNode) / 100;
        }
        if (spcAftNode !== undefined) {
            spcAfter = parseInt(spcAftNode) / 100;
        }
        
        if (lnSpcNode !== undefined && fontSize !== undefined) {
            if (lnSpcNodeType == "Pts") {
                marginTopBottomStr += "padding-top: " + ((parseInt(lnSpcNode) / 100) - fontSize) + "px;";//+ "pt;";
            } else {
                var fct = parseInt(lnSpcNode) / 100000;
                spcLines = fontSize * (fct - 1) - fontSize;// fontSize *
                var pTop = (fct > 1) ? spcLines : 0;
                var pBottom = (fct > 1) ? fontSize : 0;
                // marginTopBottomStr += "padding-top: " + spcLines + "pt;";
                // marginTopBottomStr += "padding-bottom: " + pBottom + "pt;";
                marginTopBottomStr += "padding-top: " + pBottom + "px;";// + "pt;";
                marginTopBottomStr += "padding-bottom: " + spcLines + "px;";// + "pt;";
            }
        }

        //if (spcBefNode !== undefined || lnSpcNode !== undefined) {
        marginTopBottomStr += "margin-top: " + (spcBefor - 1) + "px;";// + "pt;"; //margin-top: + spcLines // minus 1 - to fix space
        //}
        if (spcAftNode !== undefined || lnSpcNode !== undefined) {
            //marginTopBottomStr += "margin-bottom: " + ((spcAfter - fontSize < 0) ? 0 : (spcAfter - fontSize)) + "pt;"; //margin-bottom: + spcLines
            //marginTopBottomStr += "margin-bottom: " + spcAfter * (1 / 4) + "px;";// + "pt;";
            marginTopBottomStr += "margin-bottom: " + spcAfter  + "px;";// + "pt;";
        }

        //console.log("getVerticalMargins 2 fontSize:", fontSize, "lnSpcNode:", lnSpcNode, "spcLines:", spcLines, "spcBefor:", spcBefor, "spcAfter:", spcAfter)
        //console.log("getVerticalMargins 3 ", marginTopBottomStr, pNode, warpObj)

        //return spcAft + spcBef;
        return marginTopBottomStr;
    };

    // Export to window
    window.PPTXTextStyleUtils = PPTXTextStyleUtils;

})();