

var PPTXStyleUtils = (function() {
    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;
    
    var rtl_langs_array = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR","ur-PK"]

    function getFillType(node) {
            //Need to test/////////////////////////////////////////////
            //SOLID_FILL
            //PIC_FILL
            //GRADIENT_FILL
            //PATTERN_FILL
            //NO_FILL
            var fillType = "";
            if (node["a:noFill"] !== undefined) {
                fillType = "NO_FILL";
            }
            if (node["a:solidFill"] !== undefined) {
                fillType = "SOLID_FILL";
            }
            if (node["a:gradFill"] !== undefined) {
                fillType = "GRADIENT_FILL";
            }
            if (node["a:pattFill"] !== undefined) {
                fillType = "PATTERN_FILL";
            }
            if (node["a:blipFill"] !== undefined) {
                fillType = "PIC_FILL";
            }
            if (node["a:grpFill"] !== undefined) {
                fillType = "GROUP_FILL";
            }


            return fillType;
        }
    // function hexToRgbNew(hex) {
        //     var arrBuff = new ArrayBuffer(4);
        //     var vw = new DataView(arrBuff);
        //     vw.setUint32(0, parseInt(hex, 16), false);
        //     var arrByte = new Uint8Array(arrBuff);
        //     return arrByte[1] + "," + arrByte[2] + "," + arrByte[3];
        // }
        function getShapeFill(node, pNode, isSvgMode, warpObj, source) {

            // 1. presentationML
            // p:spPr/ [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
            // From slide
            //Fill Type:
            //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode)
            var fillType = getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
            //var noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            var fillColor;
            if (fillType == "NO_FILL") {
                return isSvgMode ? "none" : "";//"background-color: initial;";
            } else if (fillType == "SOLID_FILL") {
                var shpFill = node["p:spPr"]["a:solidFill"];
                fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
            } else if (fillType == "GRADIENT_FILL") {
                var shpFill = node["p:spPr"]["a:gradFill"];
                fillColor = getGradientFill(shpFill, warpObj);
                //console.log("shpFill",shpFill,grndColor.color)
            } else if (fillType == "PATTERN_FILL") {
                var shpFill = node["p:spPr"]["a:pattFill"];
                fillColor = getPatternFill(shpFill, warpObj);
            } else if (fillType == "PIC_FILL") {
                var shpFill = node["p:spPr"]["a:blipFill"];
                fillColor = getPicFill(source, shpFill, warpObj);
            }
            //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode, ", fillType: ", fillType, ", fillColor: ", fillColor, ", source: ", source)


            // 2. drawingML namespace
            if (fillColor === undefined) {
                var clrName = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
                var idx = parseInt (PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
                if (idx == 0 || idx == 1000) {
                    //no fill
                    return isSvgMode ? "none" : "";
                } else if (idx > 0 && idx < 1000) {
                    // <a:fillStyleLst> fill
                } else if (idx > 1000) {
                    //<a:bgFillStyleLst>
                }
                fillColor = getSolidFill(clrName, undefined, undefined, warpObj);
            }
            // 3. is group fill
            if (fillColor === undefined) {
                var grpFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
                if (grpFill !== undefined) {
                    //fillColor = getSolidFill(clrName, undefined, undefined, undefined, warpObj);
                    //get parent fill style - TODO
                    //console.log("ShapeFill: grpFill: ", grpFill, ", pNode: ", pNode)
                    var grpShpFill = pNode["p:grpSpPr"];
                    var spShpNode = { "p:spPr": grpShpFill }
                    return getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
                } else if (fillType == "NO_FILL") {
                    return isSvgMode ? "none" : "";
                }
            }
            //console.log("ShapeFill: fillColor: ", fillColor, ", fillType; ", fillType)

            if (fillColor !== undefined) {
                if (fillType == "GRADIENT_FILL") {
                    if (isSvgMode) {
                        // console.log("GRADIENT_FILL color", fillColor.color[0])
                        return fillColor;
                    } else {
                        var colorAry = fillColor.color;
                        var rot = fillColor.rot;

                        var bgcolor = "background: linear-gradient(" + rot + "deg,";
                        for (var i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bgcolor += "#" + colorAry[i] + ");";
                            } else {
                                bgcolor += "#" + colorAry[i] + ", ";
                            }

                        }
                        return bgcolor;
                    }
                } else if (fillType == "PIC_FILL") {
                    if (isSvgMode) {
                        return fillColor;
                    } else {

                        return "background-image:url(" + fillColor + ");";
                    }
                } else if (fillType == "PATTERN_FILL") {
                    /////////////////////////////////////////////////////////////Need to check -----------TODO
                    // if (isSvgMode) {
                    //     var color = tinycolor(fillColor);
                    //     fillColor = color.toRgbString();

                    //     return fillColor;
                    // } else {
                    var bgPtrn = "", bgSize = "", bgPos = "";
                    bgPtrn = fillColor[0];
                    if (fillColor[1] !== null && fillColor[1] !== undefined && fillColor[1] != "") {
                        bgSize = " background-size:" + fillColor[1] + ";";
                    }
                    if (fillColor[2] !== null && fillColor[2] !== undefined && fillColor[2] != "") {
                        bgPos = " background-position:" + fillColor[2] + ";";
                    }
                    return "background: " + bgPtrn + ";" + bgSize + bgPos;
                    //}
                } else {
                    if (isSvgMode) {
                        var color = tinycolor(fillColor);
                        fillColor = color.toRgbString();

                        return fillColor;
                    } else {
                        //console.log(node,"fillColor: ",fillColor,"fillType: ",fillType,"isSvgMode: ",isSvgMode)
                        return "background-color: #" + fillColor + ";";
                    }
                }
            } else {
                if (isSvgMode) {
                    return "none";
                } else {
                    return "background-color: inherit;";
                }

            }

        }

        
        function getFontType(node, type, warpObj, pFontStyle) {
            var typeface = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

            if (typeface === undefined) {
                var fontIdx = "";
                var fontGrup = "";
                if (pFontStyle !== undefined) {
                    fontIdx = PPTXXmlUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
                }
                var fontSchemeNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
                if (fontIdx == "") {
                    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                        fontIdx = "major";
                    } else {
                        fontIdx = "minor";
                    }
                }
                fontGrup = "a:" + fontIdx + "Font";
                typeface = PPTXXmlUtils.getTextByPathList(fontSchemeNode, [fontGrup, "a:latin", "attrs", "typeface"]);
            }

            return (typeface === undefined) ? "inherit" : typeface;
        }

        function getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj) {
            //text border using: text-shadow: -1px 0 black, 0 1px black, 1px 0 black, 0 -1px black;
            //{getFontColor(..) return color} -> getFontColorPr(..) return array[color,textBordr/shadow]
            //https://stackoverflow.com/questions/2570972/css-font-border
            //https://www.w3schools.com/cssref/css3_pr_text-shadow.asp
            //themeContent
            //console.log("getFontColorPr>> type:", type, ", node: ", node)
            var rPrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr"]);
            var filTyp, color, textBordr, colorType = "", highlightColor = "";
            //console.log("getFontColorPr type:", type, ", node: ", node, "pNode:", pNode, "pFontStyle:", pFontStyle)
            if (rPrNode !== undefined) {
                filTyp = getFillType(rPrNode);
                if (filTyp == "SOLID_FILL") {
                    var solidFillNode = rPrNode["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    var highlightNode = rPrNode["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    var pattFill = rPrNode["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = getBgPicFill(rPrNode, "slideBg", warpObj, undefined, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    var shpFill = rPrNode["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                } 
            }
            if (color === undefined && PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
                //lstStyle
                var lstStyledefRPr = PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
                filTyp = getFillType(lstStyledefRPr);
                if (filTyp == "SOLID_FILL") {
                    var solidFillNode = lstStyledefRPr["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    var highlightNode = lstStyledefRPr["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    var pattFill = lstStyledefRPr["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = getBgPicFill(lstStyledefRPr, "slideBg", warpObj, undefined, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    var shpFill = lstStyledefRPr["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                }

            }
            if (color === undefined) {
                var sPstyle = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
                if (sPstyle !== undefined) {
                    color = getSolidFill(sPstyle, undefined, undefined, warpObj);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                    var highlightNode = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                }
                if (color === undefined) {
                    if (pFontStyle !== undefined) {
                        color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
                        if (color !== undefined) {
                            colorType = "solid";
                        }
                    }
                }
            }
            //console.log("getFontColorPr node", node, "colorType: ", colorType,"color: ",color)

            if (color === undefined) {

                var layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
                var pPrNodeLaout = layoutMasterNode.nodeLaout;
                var pPrNodeMaster = layoutMasterNode.nodeMaster;

                if (pPrNodeLaout !== undefined) {
                    var defRpRLaout = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                    if (defRpRLaout !== undefined) {
                        color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                        var highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                        }
                        colorType = "solid";
                    }
                }
                if (color === undefined) {

                    if (pPrNodeMaster !== undefined) {
                        var defRprMaster = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                        if (defRprMaster !== undefined) {
                            color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                            var highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                            if (highlightNode !== undefined) {
                                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                            }
                            colorType = "solid";
                        }
                    }
                }
            }
            var txtEffects = [];
            var txtEffObj = {}
            //textBordr
            var txtBrdrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:ln"]);
            var textBordr = "";
            if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
                var txBrd = getBorder(node, pNode, false, "text", warpObj);
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
            var txtGlowNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
            var oGlowStr = "";
            if (txtGlowNode !== undefined) {
                var glowClr = getSolidFill(txtGlowNode, undefined, undefined, warpObj);
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
            var txtShadow = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
            var oShadowStr = "";
            if (txtShadow !== undefined) {
                //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
                //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
                //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
                //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

                var shadowClr = getSolidFill(txtShadow, undefined, undefined, warpObj);
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
        }
        function getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj) {
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
                sz = PPTXXmlUtils.getTextByPathList(node["a:fld"], ["a:rPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            if ((isNaN(fontSize) || fontSize === undefined) && node["a:t"] === undefined) {
                sz = PPTXXmlUtils.getTextByPathList(node["a:endParaRPr"], [ "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            if ((isNaN(fontSize) || fontSize === undefined) && lstStyle !== undefined) {
                sz = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlpPr, "a:defRPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
            }
            //a:spAutoFit
            var isAutoFit = false;
            var isKerning = false;
            if (textBodyNode !== undefined){
                var spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
                // if (spAutoFitNode === undefined) {
                //     spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit"]);
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
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                fontSize = parseInt(sz) / 100;
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                if (isKerning && kern !== undefined && !isNaN(fontSize) && (fontSize - parseInt(kern) / 100) > 0){
                    fontSize = fontSize - parseInt(kern) / 100;
                }
            }

            if (isNaN(fontSize) || fontSize === undefined) {
                // if (type == "shape" || type == "textBox") {
                //     type = "body";
                //     lvlpPr = "a:lvl1pPr";
                // }
                sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTables"], ["typeTable", type, "p:txBody", "a:lstStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                if (sz === undefined) {
                    if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:titleStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    } else if (type == "body" || type == "obj" || type == "dt" || type == "sldNum" || type === "textBox") {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:bodyStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    }
                    else if (type == "shape") {
                        //textBox and shape text does not indent
                        // 普通形状使用 otherStyle，与原始库保持一致
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                        isKerning = false;
                    }

                    if (sz === undefined) {
                        sz = PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "sz"]);
                        kern = (kern === undefined)? PPTXXmlUtils.getTextByPathList(warpObj["defaultTextStyle"], [lvlpPr, "a:defRPr", "attrs", "kern"]) : undefined;
                        isKerning = false;
                    }
                    //  else if (type === undefined || type == "shape") {
                    //     sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    //     kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    // } 
                    // else if (type == "textBox") {
                    //     sz = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "sz"]);
                    //     kern = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterTextStyles"], ["p:otherStyle", lvlpPr, "a:defRPr", "attrs", "kern"]);
                    // }
                } 
                fontSize = parseInt(sz) / 100;
                if (isKerning && kern !== undefined && !isNaN(fontSize) && ((fontSize - parseInt(kern) / 100) > parseInt(kern) / 100 )) {
                    fontSize = fontSize - parseInt(kern) / 100;
                    //fontSize =  parseInt(kern) / 100;
                }
            }

            var baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            if (baseline !== undefined && !isNaN(fontSize)) {
                var baselineVl = parseInt(baseline) / 100000;
                //fontSize -= 10; 
                // fontSize = fontSize * baselineVl;
                fontSize -= baselineVl;
            }

            if (!isNaN(fontSize)){
                var normAutofit = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
                if (normAutofit !== undefined && normAutofit != 0){
                    //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
                    fontSize = Math.round(fontSize * (normAutofit / 100000))
                }
            }

            return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * fontSizeFactor + "px");// + "pt");
        }

        function getFontBold(node, type, slideMasterTextStyles) {
            return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "inherit";
        }

        function getFontItalic(node, type, slideMasterTextStyles) {
            return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "inherit";
        }

        function getFontDecoration(node, type, slideMasterTextStyles) {
            ///////////////////////////////Amir///////////////////////////////
            if (node["a:rPr"] !== undefined) {
                var underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
                var strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';
                //console.log("strikethrough: "+strikethrough);

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
            /////////////////////////////////////////////////////////////////
            //return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "inherit";
        }
        ////////////////////////////////////Amir/////////////////////////////////////
        function getTextHorizontalAlign(node, pNode, type, warpObj) {
            //console.log("getTextHorizontalAlign: type: ", type, ", node: ", node)
            var getAlgn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (getAlgn === undefined) {
                getAlgn = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
            }
            if (getAlgn === undefined) {
                if (type == "title" || type == "ctrTitle" || type == "subTitle") {
                    var lvlIdx = 1;
                    var lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
                    if (lvlNode !== undefined) {
                        lvlIdx = parseInt(lvlNode) + 1;
                    }
                    var lvlStr = "a:lvl" + lvlIdx + "pPr";
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (getAlgn === undefined) {
                        getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        if (getAlgn === undefined) {
                            getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                            if (getAlgn === undefined && type === "subTitle") {
                                getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                            }
                        }
                    }
                } else if (type == "body") {
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", "a:lvl1pPr", "attrs", "algn"]);
                } else {
                    getAlgn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", "a:lvl1pPr", "attrs", "algn"]);
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
        }
        /////////////////////////////////////////////////////////////////////
        function getTextVerticalAlign(node, type, slideMasterTextStyles) {
            var baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
        }

        function getTableBorders(node, warpObj) {
            var borderStyle = "";
            if (node["a:bottom"] !== undefined) {
                var obj = {
                    "p:spPr": {
                        "a:ln": node["a:bottom"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-bottom");
            }
            if (node["a:top"] !== undefined) {
                var obj = {
                    "p:spPr": {
                        "a:ln": node["a:top"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-top");
            }
            if (node["a:right"] !== undefined) {
                var obj = {
                    "p:spPr": {
                        "a:ln": node["a:right"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-right");
            }
            if (node["a:left"] !== undefined) {
                var obj = {
                    "p:spPr": {
                        "a:ln": node["a:left"]["a:ln"]
                    }
                }
                var borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-left");
            }

            return borderStyle;
        }
        //////////////////////////////////////////////////////////////////
        function getBorder(node, pNode, isSvgMode, bType, warpObj) {
            //console.log("getBorder", node, pNode, isSvgMode, bType)
            var cssText, lineNode, subNodeTxt;

            if (bType == "shape") {
                cssText = "border: ";
                lineNode = node["p:spPr"]["a:ln"];
                //subNodeTxt = "p:spPr";
                //node["p:style"]["a:lnRef"] = 
            } else if (bType == "text") {
                cssText = "";
                lineNode = node["a:rPr"]["a:ln"];
                //subNodeTxt = "a:rPr";
            }

            //var is_noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            var is_noFill = PPTXXmlUtils.getTextByPathList(lineNode, ["a:noFill"]);
            if (is_noFill !== undefined) {
                return "hidden";
            }

            //console.log("lineNode: ", lineNode)
            if (lineNode == undefined) {
                var lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"])
                if (lnRefNode !== undefined){
                    var lnIdx = PPTXXmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
                    //console.log("lnIdx:", lnIdx, "lnStyleLst:", warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) -1])
                    lineNode = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"][Number(lnIdx) - 1];
                }
            }
            if (lineNode == undefined) {
                //is table
                cssText = "";
                lineNode = node
            }

            var borderColor;
            var borderWidth = 0;
            var borderType = "solid";
            var strokeDasharray = "0";
            if (lineNode !== undefined) {
                // Border width: 1pt = 12700, default = 0.75pt
                borderWidth = parseInt (PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
                if (isNaN(borderWidth) || borderWidth < 1) {
                    cssText += (4/3) + "px ";//"1pt ";
                } else {
                    cssText += borderWidth + "px ";// + "pt ";
                }
                // Border type
                borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
                if (borderType === undefined) {
                    borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "cmpd"]);
                }
                strokeDasharray = "0";
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
                    case undefined:
                    //console.log(borderType);
                    default:
                        cssText += "solid";
                        strokeDasharray = "0";
                }
                // Border color
                var fillTyp = getFillType(lineNode);
                //console.log("getBorder:node : fillTyp", fillTyp)
                if (fillTyp == "NO_FILL") {
                    borderColor = isSvgMode ? "none" : "";//"background-color: initial;";
                } else if (fillTyp == "SOLID_FILL") {
                    borderColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
                } else if (fillTyp == "GRADIENT_FILL") {
                    borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
                    //console.log("shpFill",shpFill,grndColor.color)
                } else if (fillTyp == "PATTERN_FILL") {
                    borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
                }

            }

            //console.log("getBorder:node : borderColor", borderColor)
            // 2. drawingML namespace
            if (borderColor === undefined) {
                //var schemeClrNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
                // if (schemeClrNode !== undefined) {
                //     var schemeClr = "a:" + PPTXXmlUtils.getTextByPathList(schemeClrNode, ["attrs", "val"]);
                //     var borderColor = getSchemeColorFromTheme(schemeClr, undefined, undefined);
                // }
                var lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
                //console.log("getBorder: lnRef : ", lnRefNode)
                if (lnRefNode !== undefined) {
                    borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                }

                // if (borderColor !== undefined) {
                //     var shade = PPTXXmlUtils.getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
                //     if (shade !== undefined) {
                //         shade = parseInt(shade) / 10000;
                //         var color = tinycolor("#" + borderColor);
                //         borderColor = color.darken(shade).toHex8();//.replace("#", "");
                //     }
                // }

            }

            //console.log("getBorder: borderColor : ", borderColor)
            if (borderColor === undefined) {
                if (isSvgMode) {
                    borderColor = "none";
                } else {
                    borderColor = "hidden";
                }
            } else {
                borderColor = "#" + borderColor; //wrong if not solid fill - TODO

            }
            cssText += " " + borderColor + " ";//wrong if not solid fill - TODO

            if (isSvgMode) {
                return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
            } else {
                return cssText + ";";
            }
            // } else {
            //     if (isSvgMode) {
            //         return { "color": 'none', "width": '0', "type": 'none', "strokeDasharray": '0' };
            //     } else {
            //         return "hidden";
            //     }
            // }
        }
        function getBackground(warpObj, slideSize, index, settings) {
            //var rslt = "";
            var slideContent = warpObj["slideContent"];
            var slideLayoutContent = warpObj["slideLayoutContent"];
            var slideMasterContent = warpObj["slideMasterContent"];

            var nodesSldLayout = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
            var nodesSldMaster = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
            // console.log("slideContent : ", slideContent)
            // console.log("slideLayoutContent : ", slideLayoutContent)
            // console.log("slideMasterContent : ", slideMasterContent)
            //console.log("warpObj : ", warpObj)
            var showMasterSp = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
            //console.log("slideLayoutContent : ", slideLayoutContent, ", showMasterSp: ", showMasterSp)
            var bgColor = getSlideBackgroundFill(warpObj, index);
            var result = "<div class='slide-background-" + index + "' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>"
            var node_ph_type_ary = [];
            if (nodesSldLayout !== undefined) {
                for (var nodeKey in nodesSldLayout) {
                    if (nodesSldLayout[nodeKey].constructor === Array) {
                        for (var i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                            var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                            // if (ph_type !== undefined && ph_type != "pic") {
                            //     node_ph_type_ary.push(ph_type);
                            // }
                            if (ph_type != "pic") {
                                result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg", 'group', settings); //slideLayoutBg , slideMasterBg
                            }
                        }
                    } else {
                        var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        // if (ph_type !== undefined && ph_type != "pic") {
                        //     node_ph_type_ary.push(ph_type);
                        // }
                        if (ph_type != "pic") {
                            result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg", 'group', settings); //slideLayoutBg, slideMasterBg
                        }
                    }
                }
            }
            if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
                for (var nodeKey in nodesSldMaster) {
                    if (nodesSldMaster[nodeKey].constructor === Array) {
                        for (var i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                            var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                            //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                            result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg", 'group', settings); //slideLayoutBg , slideMasterBg
                            //}
                        }
                    } else {
                        var ph_type = PPTXXmlUtils.getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                        //if (node_ph_type_ary.indexOf(ph_type) > -1) {
                        result += PPTXNodeUtils.processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg", 'group', settings); //slideLayoutBg, slideMasterBg
                        //}
                    }
                }
            }
            return result;

        }
        function getSlideBackgroundFill(warpObj, index) {
            var slideContent = warpObj["slideContent"];
            var slideLayoutContent = warpObj["slideLayoutContent"];
            var slideMasterContent = warpObj["slideMasterContent"];

            //console.log("slideContent: ", slideContent)
            //console.log("slideLayoutContent: ", slideLayoutContent)
            //console.log("slideMasterContent: ", slideMasterContent)
            //PPTXShapeUtils.getFillType(node)
            var bgPr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
            var bgRef = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
            //console.log("slideContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
            var bgcolor;
            if (bgPr !== undefined) {
                //bgcolor = "background-color: blue;";
                var bgFillTyp = getFillType(bgPr);

                if (bgFillTyp == "SOLID_FILL") {
                    var sldFill = bgPr["a:solidFill"];
                    var clrMapOvr;
                    var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                        if (sldClrMapOvr !== undefined) {
                            clrMapOvr = sldClrMapOvr;
                        } else {
                            clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                        }

                    }
                    var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                    //var sldTint = getColorOpacity(sldFill);
                    //console.log("bgColor: ", bgColor)
                    //bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                    bgcolor = "background: #" + sldBgClr + ";";

                } else if (bgFillTyp == "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                } else if (bgFillTyp == "PIC_FILL") {
                    //console.log("PIC_FILL - ", bgFillTyp, bgPr, warpObj);
                    bgcolor = getBgPicFill(bgPr, "slideBg", warpObj, undefined, index);

                }
                //console.log(slideContent,slideMasterContent,color_ary,tint_ary,rot,bgcolor)
            } else if (bgRef !== undefined) {
                //console.log("slideContent",bgRef)
                var clrMapOvr;
                var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
                var phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);

                // if (bgRef["a:srgbClr"] !== undefined) {
                //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                //     var schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);
                //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                // }
                var idx = Number(bgRef["attrs"]["idx"]);


                if (idx == 0 || idx == 1000) {
                    //no background
                } else if (idx > 0 && idx < 1000) {
                    //fillStyleLst in themeContent
                    //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                    //bgcolor = "background: red;";
                } else if (idx > 1000) {
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    var trueIdx = idx - 1000;
                    // themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    var sortblAry = [];
                    Object.keys(bgFillLst).forEach(function (key) {
                        var bgFillLstTyp = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (var i = 0; i < bgFillLstTyp.length; i++) {
                                    var obj = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp[i]["attrs"]["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            } else {
                                var obj = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]["order"]
                                }
                                sortblAry.push(obj)
                            }
                        }
                    });
                    var sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort(function (a, b) {
                        return a.idex - b.idex;
                    });
                    var bgFillLstIdx = sortByOrder[trueIdx - 1];
                    var bgFillTyp = getFillType(bgFillLstIdx);
                    if (bgFillTyp == "SOLID_FILL") {
                        var sldFill = bgFillLstIdx["a:solidFill"];
                        var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                        bgcolor = "background: #" + sldBgClr + ";";
                        //console.log("slideMasterContent - sldFill",sldFill)
                    } else if (bgFillTyp == "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                    } else {
                        console.log(bgFillTyp)
                    }
                }

            }
            else {
                bgPr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
                bgRef = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);
                //console.log("slideLayoutContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
                var clrMapOvr;
                var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }
                if (bgPr !== undefined) {
                    var bgFillTyp = getFillType(bgPr);
                    if (bgFillTyp == "SOLID_FILL") {
                        var sldFill = bgPr["a:solidFill"];

                        var sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                        bgcolor = "background: #" + sldBgClr + ";";
                    } else if (bgFillTyp == "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                    } else if (bgFillTyp == "PIC_FILL") {
                        bgcolor = getBgPicFill(bgPr, "slideLayoutBg", warpObj, undefined, index);

                    }
                    //console.log("slideLayoutContent",bgcolor)
                } else if (bgRef !== undefined) {
                    console.log("slideLayoutContent: bgRef", bgRef)
                    //bgcolor = "background: white;";
                    var phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
                    var idx = Number(bgRef["attrs"]["idx"]);
                    //console.log("phClr=", phClr, "idx=", idx)

                    if (idx == 0 || idx == 1000) {
                        //no background
                    } else if (idx > 0 && idx < 1000) {
                        //fillStyleLst in themeContent
                        //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                        //bgcolor = "background: red;";
                    } else if (idx > 1000) {
                        //bgFillStyleLst  in themeContent
                        //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                        var trueIdx = idx - 1000;
                        var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                        var sortblAry = [];
                        Object.keys(bgFillLst).forEach(function (key) {
                            //console.log("cubicBezTo[" + key + "]:");
                            var bgFillLstTyp = bgFillLst[key];
                            if (key != "attrs") {
                                if (bgFillLstTyp.constructor === Array) {
                                    for (var i = 0; i < bgFillLstTyp.length; i++) {
                                        var obj = {};
                                        obj[key] = bgFillLstTyp[i];
                                        obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp[i]["attrs"]["order"]
                                        }
                                        sortblAry.push(obj)
                                    }
                                } else {
                                    var obj = {};
                                    obj[key] = bgFillLstTyp;
                                    obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp["attrs"]["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            }
                        });
                        var sortByOrder = sortblAry.slice(0);
                        sortByOrder.sort(function (a, b) {
                            return a.idex - b.idex;
                        });
                        var bgFillLstIdx = sortByOrder[trueIdx - 1];
                        var bgFillTyp = getFillType(bgFillLstIdx);
                        if (bgFillTyp == "SOLID_FILL") {
                            var sldFill = bgFillLstIdx["a:solidFill"];
                            //console.log("sldFill: ", sldFill)
                            //var sldTint = getColorOpacity(sldFill);
                            //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                            var sldBgClr = getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
                            //console.log("bgcolor: ", bgcolor)
                            bgcolor = "background: #" + sldBgClr + ";";
                        } else if (bgFillTyp == "GRADIENT_FILL") {
                            //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                            bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                        } else if (bgFillTyp == "PIC_FILL") {
                            //theme rels
                            //console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, warpObj);
                            bgcolor = getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, index);
                        } else {
                            console.log(bgFillTyp)
                        }
                    }
                } else {
                    bgPr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
                    bgRef = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);

                    var clrMap = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                    //console.log("slideMasterContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
                    if (bgPr !== undefined) {
                        var bgFillTyp = getFillType(bgPr);
                        if (bgFillTyp == "SOLID_FILL") {
                            var sldFill = bgPr["a:solidFill"];
                            var sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj);
                            // var sldTint = getColorOpacity(sldFill);
                            // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                            bgcolor = "background: #" + sldBgClr + ";";
                        } else if (bgFillTyp == "GRADIENT_FILL") {
                            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                        } else if (bgFillTyp == "PIC_FILL") {
                            bgcolor = getBgPicFill(bgPr, "slideMasterBg", warpObj, undefined, index);
                        }
                    } else if (bgRef !== undefined) {
                        //var obj={
                        //    "a:solidFill": bgRef
                        //}
                        var phClr = getSolidFill(bgRef, clrMap, undefined, warpObj);
                        // var phClr;
                        // if (bgRef["a:srgbClr"] !== undefined) {
                        //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                        // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                        //     var schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);

                        //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                        // }
                        var idx = Number(bgRef["attrs"]["idx"]);
                        //console.log("phClr=", phClr, "idx=", idx)

                        if (idx == 0 || idx == 1000) {
                            //no background
                        } else if (idx > 0 && idx < 1000) {
                            //fillStyleLst in themeContent
                            //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                            //bgcolor = "background: red;";
                        } else if (idx > 1000) {
                            //bgFillStyleLst  in themeContent
                            //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                            var trueIdx = idx - 1000;
                            var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                            var sortblAry = [];
                            Object.keys(bgFillLst).forEach(function (key) {
                                //console.log("cubicBezTo[" + key + "]:");
                                var bgFillLstTyp = bgFillLst[key];
                                if (key != "attrs") {
                                    if (bgFillLstTyp.constructor === Array) {
                                        for (var i = 0; i < bgFillLstTyp.length; i++) {
                                            var obj = {};
                                            obj[key] = bgFillLstTyp[i];
                                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                            obj["attrs"] = {
                                                "order": bgFillLstTyp[i]["attrs"]["order"]
                                            }
                                            sortblAry.push(obj)
                                        }
                                    } else {
                                        var obj = {};
                                        obj[key] = bgFillLstTyp;
                                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp["attrs"]["order"]
                                        }
                                        sortblAry.push(obj)
                                    }
                                }
                            });
                            var sortByOrder = sortblAry.slice(0);
                            sortByOrder.sort(function (a, b) {
                                return a.idex - b.idex;
                            });
                            var bgFillLstIdx = sortByOrder[trueIdx - 1];
                            var bgFillTyp = getFillType(bgFillLstIdx);
                            //console.log("bgFillLstIdx: ", bgFillLstIdx, ", bgFillTyp: ", bgFillTyp, ", phClr: ", phClr);
                            if (bgFillTyp == "SOLID_FILL") {
                                var sldFill = bgFillLstIdx["a:solidFill"];
                                //console.log("sldFill: ", sldFill)
                                //var sldTint = getColorOpacity(sldFill);
                                //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                                var sldBgClr = getSolidFill(sldFill, clrMap, phClr, warpObj);
                                //console.log("bgcolor: ", bgcolor)
                                bgcolor = "background: #" + sldBgClr + ";";
                            } else if (bgFillTyp == "GRADIENT_FILL") {
                                //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                                bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                            } else if (bgFillTyp == "PIC_FILL") {
                                //theme rels
                                // console.log("PIC_FILL - ", bgFillTyp, bgFillLstIdx, bgFillLst, warpObj);
                                bgcolor = getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, index);
                            } else {
                                console.log(bgFillTyp)
                            }
                        }
                    }
                }
            }

            //console.log("bgcolor: ", bgcolor)
            return bgcolor;
        }
        function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
            var bgcolor = "";
            if (bgPr !== undefined) {
                var grdFill = bgPr["a:gradFill"];
                var gsLst = grdFill["a:gsLst"]["a:gs"];
                //var startColorNode, endColorNode;
                var color_ary = [];
                var pos_ary = [];
                //var tint_ary = [];
                for (var i = 0; i < gsLst.length; i++) {
                    var lo_tint;
                    var lo_color = "";
                    var lo_color = getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
                    var pos = PPTXXmlUtils.getTextByPathList(gsLst[i], ["attrs", "pos"])
                    //console.log("pos: ", pos)
                    if (pos !== undefined) {
                        pos_ary[i] = pos / 1000 + "%";
                    } else {
                        pos_ary[i] = "";
                    }
                    //console.log("lo_color", lo_color)
                    color_ary[i] = "#" + lo_color;
                    //tint_ary[i] = (lo_tint !== undefined) ? parseInt(lo_tint) / 100000 : 1;
                }
                //get rot
                var lin = grdFill["a:lin"];
                var rot = 90;
                if (lin !== undefined) {
                    rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]);// + 270;
                    //console.log("rot: ", rot)
                    rot = rot + 90;
                }
                bgcolor = "background: linear-gradient(" + rot + "deg,";
                for (var i = 0; i < gsLst.length; i++) {
                    if (i == gsLst.length - 1) {
                        //if (phClr === undefined) {
                        //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ");";
                        bgcolor += color_ary[i] + " " + pos_ary[i] + ");";
                        //} else {
                        //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ");";
                        // bgcolor += "" + phClr + ";";;
                        //}
                    } else {
                        //if (phClr === undefined) {
                        //bgcolor += "rgba(" + hexToRgbNew(color_ary[i]) + "," + tint_ary[i] + ")" + ", ";
                        bgcolor += color_ary[i] + " " + pos_ary[i] + ", ";;
                        //} else {
                        //bgcolor += "rgba(" + hexToRgbNew(phClr) + "," + tint_ary[i] + ")" + ", ";
                        // bgcolor += phClr + ", ";
                        //}
                    }
                }
            } else {
                if (phClr !== undefined) {
                    //bgcolor = "rgba(" + hexToRgbNew(phClr) + ",0);";
                    //bgcolor = phClr + ");";
                    bgcolor = "background: #" + phClr + ";";
                }
            }
            return bgcolor;
        }
        function getBgPicFill(bgPr, sorce, warpObj, phClr, index) {
            //console.log("getBgPicFill bgPr", bgPr)
            var bgcolor;
            var picFillBase64 = getPicFill(sorce, bgPr["a:blipFill"], warpObj);
            var ordr = bgPr["attrs"]["order"];
            var aBlipNode = bgPr["a:blipFill"]["a:blip"];
            //a:duotone
            var duotone = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
            if (duotone !== undefined) {
                //console.log("pic duotone: ", duotone)
                var clr_ary = [];
                // duotone.forEach(function (clr) {
                //     console.log("pic duotone clr: ", clr)
                // }) 
                Object.keys(duotone).forEach(function (clr_type) {
                    //console.log("pic duotone clr: clr_type: ", clr_type, duotone[clr_type])
                    if (clr_type != "attrs") {
                        var obj = {};
                        obj[clr_type] = duotone[clr_type];
                        clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    }
                    // Object.keys(duotone[clr_type]).forEach(function (clr) {
                    //     if (clr != "order") {
                    //         var obj = {};
                    //         obj[clr_type] = duotone[clr_type][clr];
                    //         clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    //     }
                    // })
                })
                //console.log("pic duotone clr_ary: ", clr_ary);
                //filter: url(file.svg#filter-element-id)
                //https://codepen.io/bhenbe/pen/QEZOvd
                //https://www.w3schools.com/cssref/css3_pr_filter.asp

                // var color1 = clr_ary[0];
                // var color2 = clr_ary[1];
                // var cssName = "";

                // var styleText_before_after = "content: '';" +
                //     "display: block;" +
                //     "width: 100%;" +
                //     "height: 100%;" +
                //     // "z-index: 1;" +
                //     "position: absolute;" +
                //     "top: 0;" +
                //     "left: 0;";

                // var cssName = "slide-background-" + index + "::before," + " .slide-background-" + index + "::after";
                // styleTable[styleText_before_after] = {
                //     "name": cssName,
                //     "text": styleText_before_after
                // };


                // var styleText_after = "background-color: #" + clr_ary[1] + ";" +
                //     "mix-blend-mode: darken;";

                // cssName = "slide-background-" + index + "::after";
                // styleTable[styleText_after] = {
                //     "name": cssName,
                //     "text": styleText_after
                // };

                // var styleText_before = "background-color: #" + clr_ary[0] + ";" +
                //     "mix-blend-mode: lighten;";

                // cssName = "slide-background-" + index + "::before";
                // styleTable[styleText_before] = {
                //     "name": cssName,
                //     "text": styleText_before
                // };

            }
            //a:alphaModFix
            var aphaModFixNode = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"])
            var imgOpacity = "";
            if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
                var amt = parseInt(aphaModFixNode["amt"]) / 100000;
                //var opacity = amt;
                imgOpacity = "opacity:" + amt + ";";

            }
            //a:tile

            var tileNode = PPTXXmlUtils.getTextByPathList(bgPr, ["a:blipFill", "a:tile", "attrs"])
            var prop_style = "";
            if (tileNode !== undefined && tileNode["sx"] !== undefined) {
                var sx = (parseInt(tileNode["sx"]) / 100000);
                var sy = (parseInt(tileNode["sy"]) / 100000);
                var tx = (parseInt(tileNode["tx"]) / 100000);
                var ty = (parseInt(tileNode["ty"]) / 100000);
                var algn = tileNode["algn"]; //tl(top left),t(top), tr(top right), l(left), ctr(center), r(right), bl(bottom left), b(bottm) , br(bottom right)
                var flip = tileNode["flip"]; //none,x,y ,xy

                prop_style += "background-repeat: round;"; //repeat|repeat-x|repeat-y|no-repeat|space|round|initial|inherit;
                //prop_style += "background-size: 300px 100px;"; size (w,h,sx, sy) -TODO
                //prop_style += "background-position: 50% 40%;"; //offset (tx, ty) -TODO
            }
            //a:srcRect
            //a:stretch => a:fillRect =>attrs (l:-17000, r:-17000)
            var stretch = PPTXXmlUtils.getTextByPathList(bgPr, ["a:blipFill", "a:stretch"]);
            if (stretch !== undefined) {
                var fillRect = PPTXXmlUtils.getTextByPathList(stretch, ["a:fillRect", "attrs"]);
                //console.log("getBgPicFill=>bgPr: ", bgPr)
                // var top = fillRect["t"], right = fillRect["r"], bottom = fillRect["b"], left = fillRect["l"];
                prop_style += "background-repeat: no-repeat;";
                prop_style += "background-position: center;";
                if (fillRect !== undefined) {
                    //prop_style += "background-size: contain, cover;";
                    prop_style += "background-size:  100% 100%;;";
                }
            }
            bgcolor = "background: url(" + picFillBase64 + ");  z-index: " + ordr + ";" + prop_style + imgOpacity;

            return bgcolor;
        }
      
        
        function getGradientFill(node, warpObj) {
            //console.log("getGradientFill: node", node)
            var gsLst = node["a:gsLst"]["a:gs"];
            //get start color
            var color_ary = [];
            var tint_ary = [];
            for (var i = 0; i < gsLst.length; i++) {
                var lo_tint;
                var lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
                //console.log("lo_color",lo_color)
                color_ary[i] = lo_color;
            }
            //get rot
            var lin = node["a:lin"];
            var rot = 0;
            if (lin !== undefined) {
                rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]) + 90;
            }
            return {
                "color": color_ary,
                "rot": rot
            }
        }
        function getPicFill(type, node, warpObj) {
            //Need to test/////////////////////////////////////////////
            //rId
            //TODO - Image Properties - Tile, Stretch, or Display Portion of Image
            //(http://officeopenxml.com/drwPic-tile.php)
            var img;
            var rId = node["a:blip"]["attrs"]["r:embed"];
            var imgPath;
            //console.log("getPicFill(...) rId: ", rId, ", warpObj: ", warpObj, ", type: ", type)
            if (type == "slideBg" || type == "slide") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
            } else if (type == "slideLayoutBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
            } else if (type == "slideMasterBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
            } else if (type == "themeBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
            } else if (type == "diagramBg") {
                imgPath = PPTXXmlUtils.getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
            }
            if (imgPath === undefined) {
                return undefined;
            }
            img = PPTXXmlUtils.getTextByPathList(warpObj, ["loaded-images", imgPath]); //, type, rId
            if (img === undefined) {
                 // 确定上下文类型用于路径解析
                var context = 'slide';
                if (type == "slideMasterBg") {
                    context = 'master';
                } else if (type == "slideLayoutBg") {
                    context = 'layout';
                }
                imgPath = PPTXXmlUtils.resolveMediaPath(imgPath, context, '');

                var imgExt = imgPath.split(".").pop();
                if (imgExt == "xml") {
                    return undefined;
                }
                var imgFile = warpObj["zip"].file(imgPath);
                if (imgFile === null || imgFile === undefined) {
                    console.warn("Image file not found:", imgPath);
                    return undefined;
                }
                var imgArrayBuffer = imgFile.asArrayBuffer();
                var imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                img = "data:" + imgMimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
                //warpObj["loaded-images"][imgPath] = img; //"defaultTextStyle": defaultTextStyle,
                setTextByPathList(warpObj, ["loaded-images", imgPath], img); //, type, rId
            }
            return img;
        }
        function getPatternFill(node, warpObj) {
            //https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Images/Using_CSS_gradients
            //https://cssgradient.io/blog/css-gradient-text/
            //https://css-tricks.com/background-patterns-simplified-by-conic-gradients/
            //https://stackoverflow.com/questions/6705250/how-to-get-a-pattern-into-a-written-text-via-css
            //https://stackoverflow.com/questions/14072142/striped-text-in-css
            //https://css-tricks.com/stripes-css/
            //https://yuanchuan.dev/gradient-shapes/
            var fgColor = "", bgColor = "", prst = "";
            var bgClr = node["a:bgClr"];
            var fgClr = node["a:fgClr"];
            prst = node["attrs"]["prst"];
            fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
            bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
            //var angl_ary = getAnglefromParst(prst);
            //var ptrClr = "repeating-linear-gradient(" + angl + "deg,  #" + bgColor + ",#" + fgColor + " 2px);"
            //linear-gradient(0deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black), 
            //linear-gradient(90deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black);
            var linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
            //console.log("getPatternFill: node:", node, ", prst: ", prst, ", fgColor: ", fgColor, ", bgColor:", bgColor, ', linear_gradient: ', linear_gradient)
            return linear_gradient;
        }

        function getLinerGrandient(prst, bgColor, fgColor) {
            // dashDnDiag (Dashed Downward Diagonal)-V
            // dashHorz (Dashed Horizontal)-V
            // dashUpDiag(Dashed Upward DIagonal)-V
            // dashVert(Dashed Vertical)-V
            // diagBrick(Diagonal Brick)-V
            // divot(Divot)-VX
            // dkDnDiag(Dark Downward Diagonal)-V
            // dkHorz(Dark Horizontal)-V
            // dkUpDiag(Dark Upward Diagonal)-V
            // dkVert(Dark Vertical)-V
            // dotDmnd(Dotted Diamond)-VX
            // dotGrid(Dotted Grid)-V
            // horzBrick(Horizontal Brick)-V
            // lgCheck(Large Checker Board)-V
            // lgConfetti(Large Confetti)-V
            // lgGrid(Large Grid)-V
            // ltDnDiag(Light Downward Diagonal)-V
            // ltHorz(Light Horizontal)-V
            // ltUpDiag(Light Upward Diagonal)-V
            // ltVert(Light Vertical)-V
            // narHorz(Narrow Horizontal)-V
            // narVert(Narrow Vertical)-V
            // openDmnd(Open Diamond)-V
            // pct10(10 %)-V
            // pct20(20 %)-V
            // pct25(25 %)-V
            // pct30(30 %)-V
            // pct40(40 %)-V
            // pct5(5 %)-V
            // pct50(50 %)-V
            // pct60(60 %)-V
            // pct70(70 %)-V
            // pct75(75 %)-V
            // pct80(80 %)-V
            // pct90(90 %)-V
            // smCheck(Small Checker Board) -V
            // smConfetti(Small Confetti)-V
            // smGrid(Small Grid) -V
            // solidDmnd(Solid Diamond)-V
            // sphere(Sphere)-V
            // trellis(Trellis)-VX
            // wave(Wave)-V
            // wdDnDiag(Wide Downward Diagonal)-V
            // wdUpDiag(Wide Upward Diagonal)-V
            // weave(Weave)-V
            // zigZag(Zig Zag)-V
            // shingle(Shingle)-V
            // plaid(Plaid)-V
            // cross (Cross)
            // diagCross(Diagonal Cross)
            // dnDiag(Downward Diagonal)
            // horz(Horizontal)
            // upDiag(Upward Diagonal)
            // vert(Vertical)
            switch (prst) {
                case "smGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "4px 4px"];
                    break
                case "dotGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "8px 8px"];
                    break
                case "lgGrid":
                    return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1.5px ), " +
                        "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1.5px)  #" + bgColor + ";", "8px 8px"];
                    break
                case "wdUpDiag":
                    //return ["repeating-linear-gradient(-45deg,  #" + bgColor + ", #" + bgColor + " 1px,#" + fgColor + " 5px);"];
                    return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                    // return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 1px),  #" + fgColor + " 50%, transparent calc(50% + 1px),  transparent 100%) " +
                    //     "#" + bgColor + ";", "6px 6px"];
                    break
                case "dkUpDiag":
                    return ["repeating-linear-gradient(-45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
                    break
                case "ltUpDiag":
                    return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                    break
                case "wdDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                    break
                case "dkDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
                    break
                case "ltDnDiag":
                    return ["repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                    break
                case "dkHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
                    break
                case "ltHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                    break
                case "narHorz":
                    return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                    break
                case "dkVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
                    break
                case "ltVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
                    break
                case "narVert":
                    return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
                    break
                case "lgCheck":
                case "smCheck":
                    var size = "";
                    var pos = "";
                    if (prst == "lgCheck") {
                        size = "8px 8px";
                        pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
                    } else {
                        size = "4px 4px";
                        pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
                    }
                    return ["linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0) " +
                        "#" + bgColor + ";", size, pos];
                    break
                // case "smCheck":
                //     return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                //         "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%)  " +
                //         "#" + bgColor + ";", "4px 4px"];
                //     break 

                case "dashUpDiag":
                    return ["repeating-linear-gradient(152deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "dashDnDiag":
                    return ["repeating-linear-gradient(45deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "diagBrick":
                    return ["linear-gradient(45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                        "linear-gradient(-45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                        "linear-gradient(-45deg, transparent 65%,  #" + fgColor + " 80%, transparent 0) " +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "horzBrick":
                    return ["linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                        "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px) " +
                        "#" + fgColor + ";", "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];
                    break

                case "dashVert":
                    return ["linear-gradient(0deg,  #" + bgColor + " 30%, transparent 30%)," +
                        "linear-gradient(90deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "dashHorz":
                    return ["linear-gradient(90deg,  #" + bgColor + " 30%, transparent 30%)," +
                        "linear-gradient(0deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "solidDmnd":
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", "8px 8px"];
                    break
                case "openDmnd":
                    return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                        "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%) " +
                        "#" + bgColor + ";", "8px 8px"];
                    break

                case "dotDmnd":
                    return ["radial-gradient(#" + fgColor + " 15%, transparent 0), " +
                        "radial-gradient(#" + fgColor + " 15%, transparent 0) " +
                        "#" + bgColor + ";", "4px 4px", "0 0, 2px 2px"];
                    break
                case "zigZag":
                case "wave":
                    var size = "";
                    if (prst == "zigZag") size = "0";
                    else size = "1px";
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "lgConfetti":
                case "smConfetti":
                    var size = "";
                    if (prst == "lgConfetti") size = "4px 4px";
                    else size = "2px 2px";
                    return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px, " +
                        "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                        "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px , " +
                        "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                        "#" + bgColor + ";", size];
                    break
                // case "weave":
                //     return ["linear-gradient(45deg,  #" + bgColor + " 5%, transparent 25%) 50px 0, " +
                //         "linear-gradient(135deg,  #" + bgColor + " 25%, transparent 25%) 50px 0, " +
                //         "linear-gradient(45deg,  #" + bgColor + " 25%, transparent 25%) " +
                //         "#" + fgColor + ";", "4px 4px"];
                //     //background: linear-gradient(45deg, #dca 12%, transparent 0, transparent 88%, #dca 0),
                //     //linear-gradient(135deg, transparent 37 %, #a85 0, #a85 63 %, transparent 0),
                //     //linear-gradient(45deg, transparent 37 %, #dca 0, #dca 63 %, transparent 0) #753;
                //     // background-size: 25px 25px;
                //     break;

                case "plaid":
                    return ["linear-gradient(0deg, transparent, transparent 25%, #" + fgColor + "33 25%, #" + fgColor + "33 50%)," +
                        "linear-gradient(90deg, transparent, transparent 25%, #" + fgColor + "66 25%, #" + fgColor + "66 50%) " +
                        "#" + bgColor + ";", "4px 4px"];
                    /**
                        background-color: #6677dd;
                        background-image: 
                        repeating-linear-gradient(0deg, transparent, transparent 35px, rgba(255, 255, 255, 0.2) 35px, rgba(255, 255, 255, 0.2) 70px), 
                        repeating-linear-gradient(90deg, transparent, transparent 35px, rgba(255,255,255,0.4) 35px, rgba(255,255,255,0.4) 70px);
                     */
                    break;
                case "sphere":
                    return ["radial-gradient(#" + fgColor + " 50%, transparent 50%)," +
                        "#" + bgColor + ";", "4px 4px"];
                    break
                case "weave":
                case "shingle":
                    return ["linear-gradient(45deg, #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px), " +
                        "linear-gradient(-45deg,  #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px) 0 4px, " +
                        "#" + bgColor + ";", "4px 8px"];
                    break
                //background:
                //linear-gradient(45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px),
                //linear-gradient(-45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px)0 4px;
                //background-color:#708090;
                //background-size: 4px 8px;
                case "pct5":
                case "pct10":
                case "pct20":
                case "pct25":
                case "pct30":
                case "pct40":
                case "pct50":
                case "pct60":
                case "pct70":
                case "pct75":
                case "pct80":
                case "pct90":
                //case "dotDmnd":
                case "trellis":
                case "divot":
                    var px_pr_ary;
                    switch (prst) {
                        case "pct5":
                            px_pr_ary = ["0.3px", "10%", "2px 2px"];
                            break
                        case "divot":
                            px_pr_ary = ["0.3px", "40%", "4px 4px"];
                            break
                        case "pct10":
                            px_pr_ary = ["0.3px", "20%", "2px 2px"];
                            break
                        case "pct20":
                            //case "dotDmnd":
                            px_pr_ary = ["0.2px", "40%", "2px 2px"];
                            break
                        case "pct25":
                            px_pr_ary = ["0.2px", "50%", "2px 2px"];
                            break
                        case "pct30":
                            px_pr_ary = ["0.5px", "50%", "2px 2px"];
                            break
                        case "pct40":
                            px_pr_ary = ["0.5px", "70%", "2px 2px"];
                            break
                        case "pct50":
                            px_pr_ary = ["0.09px", "90%", "2px 2px"];
                            break
                        case "pct60":
                            px_pr_ary = ["0.3px", "90%", "2px 2px"];
                            break
                        case "pct70":
                        case "trellis":
                            px_pr_ary = ["0.5px", "95%", "2px 2px"];
                            break
                        case "pct75":
                            px_pr_ary = ["0.65px", "100%", "2px 2px"];
                            break
                        case "pct80":
                            px_pr_ary = ["0.85px", "100%", "2px 2px"];
                            break
                        case "pct90":
                            px_pr_ary = ["1px", "100%", "2px 2px"];
                            break
                    }
                    return ["radial-gradient(#" + fgColor + " " + px_pr_ary[0] + ", transparent " + px_pr_ary[1] + ")," +
                        "#" + bgColor + ";", px_pr_ary[2]];
                    break
                default:
                    return [0, 0];
            }
        }

        function getSolidFill(node, clrMap, phClr, warpObj) {

            if (node === undefined) {
                return undefined;
            }

            //console.log("getSolidFill node: ", node)
            var color = "";
            var clrNode;
            if (node["a:srgbClr"] !== undefined) {
                clrNode = node["a:srgbClr"];
                color = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //#...
            } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
                clrNode = node["a:schemeClr"];
                var schemeClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
                color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
                //console.log("schemeClr: ", schemeClr, "color: ", color)
            } else if (node["a:scrgbClr"] !== undefined) {
                clrNode = node["a:scrgbClr"];
                //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
                var defBultColorVals = clrNode["attrs"];
                var red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
                var green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
                var blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
                //var scrgbClr = red + "," + green + "," + blue;
                color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
                //console.log("scrgbClr: " + scrgbClr);

            } else if (node["a:prstClr"] !== undefined) {
                clrNode = node["a:prstClr"];
                //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
                var prstClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]["val"];
                color = getColorName2Hex(prstClr);
                //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
            } else if (node["a:hslClr"] !== undefined) {
                clrNode = node["a:hslClr"];
                //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
                var defBultColorVals = clrNode["attrs"];
                var hue = Number(defBultColorVals["hue"]) / 100000;
                var sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
                var lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
                //var hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
                var hsl2rgb = hslToRgb(hue, sat, lum);
                color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
                //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
                // console.log("hslClr: " + hslClr);
            } else if (node["a:sysClr"] !== undefined) {
                clrNode = node["a:sysClr"];
                //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
                var sysClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
                if (sysClr !== undefined) {
                    color = sysClr;
                }
            }
            //console.log("color: [%cstart]", "color: #" + color, tinycolor(color).toHslString(), color)

            //fix color -------------------------------------------------------- TODO 
            //
            //1. "alpha":
            //Specifies the opacity as expressed by a percentage value.
            // [Example: The following represents a green solid fill which is 50 % opaque
            // < a: solidFill >
            //     <a:srgbClr val="00FF00">
            //         <a:alpha val="50%" />
            //     </a:srgbClr>
            // </a: solidFill >
            var isAlpha = false;
            var alpha = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
            //console.log("alpha: ", alpha)
            if (!isNaN(alpha)) {
                // var al_color = new colz.Color(color);
                // al_color.setAlpha(alpha);
                // var ne_color = al_color.rgba.toString();
                // color = (rgba2hex(ne_color))
                var al_color = tinycolor(color);
                al_color.setAlpha(alpha);
                color = al_color.toHex8()
                isAlpha = true;
                //console.log("al_color: ", al_color, ", color: ", color)
            }
            //2. "alphaMod":
            // Specifies the opacity as expressed by a percentage relative to the input color.
            //     [Example: The following represents a green solid fill which is 50 % opaque
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:alphaMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //3. "alphaOff":
            // Specifies the opacity as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the opacity beyond 100 %, decreases never decrease
            // the opacity below 0 %.
            // [Example: The following represents a green solid fill which is 90 % opaque
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:alphaOff val="-10%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //4. "blue":
            //Specifies the value of the blue component.The assigned value is specified as a
            //percentage with 0 % indicating minimal blue and 100 % indicating maximum blue.
            //  [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //      to value RRGGBB = (00, FF, FF)
            //          <a: solidFill >
            //              <a:srgbClr val="00FF00">
            //                  <a:blue val="100%" />
            //              </a:srgbClr>
            //          </a: solidFill >
            //5. "blueMod"
            // Specifies the blue component as expressed by a percentage relative to the input color
            // component.Increases never increase the blue component beyond 100 %, decreases
            // never decrease the blue component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            //     to value RRGGBB = (00, 00, 80)
            //     < a: solidFill >
            //         <a:srgbClr val="0000FF">
            //             <a:blueMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //6. "blueOff"
            // Specifies the blue component as expressed by a percentage offset increase or decrease
            // to the input color component.Increases never increase the blue component
            // beyond 100 %, decreases never decrease the blue component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            // to value RRGGBB = (00, 00, CC)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:blueOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //7. "comp" - This element specifies that the color rendered should be the complement of its input color with the complement
            // being defined as such.Two colors are called complementary if, when mixed they produce a shade of grey.For
            // instance, the complement of red which is RGB(255, 0, 0) is cyan.(<a:comp/>)

            //8. "gamma" - This element specifies that the output color rendered by the generating application should be the sRGB gamma
            //              shift of the input color.

            //9. "gray" - This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
            //              primaries.

            //10. "green":
            // Specifies the value of the green component. The assigned value is specified as a
            // percentage with 0 % indicating minimal green and 100 % indicating maximum green.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
            // to value RRGGBB = (00, FF, FF)
            //     < a: solidFill >
            //         <a:srgbClr val="0000FF">
            //             <a:green val="100%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //11. "greenMod":
            // Specifies the green component as expressed by a percentage relative to the input color
            // component.Increases never increase the green component beyond 100 %, decreases
            // never decrease the green component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, 80, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:greenMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //12. "greenOff":
            // Specifies the green component as expressed by a percentage offset increase or decrease
            // to the input color component.Increases never increase the green component
            // beyond 100 %, decreases never decrease the green component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, CC, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:greenOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //13. "hue" (This element specifies a color using the HSL color model):
            // This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
            // < a: solidFill >
            //     <a:hslClr hue="14400000" sat="100%" lum="50%">
            // </a:solidFill>
            // <a:solidFill>
            //     <a:hslClr hue="0" sat="100%" lum="50%">
            //         <a:hue val="14400000"/>
            //     <a:hslClr/>
            // </a:solidFill>

            //14. "hueMod" (This element specifies a color using the HSL color model):
            // Specifies the hue as expressed by a percentage relative to the input color.
            // [Example: The following manipulates the fill color from having RGB value RRGGBB = (00, FF, 00) to value RRGGBB = (FF, FF, 00)
            //         < a: solidFill >
            //             <a:srgbClr val="00FF00">
            //                 <a:hueMod val="50%" />
            //             </a:srgbClr>
            //         </a: solidFill >

            var hueMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
            //console.log("hueMod: ", hueMod)
            if (!isNaN(hueMod)) {
                color = applyHueMod(color, hueMod, isAlpha);
            }
            //15. "hueOff"(This element specifies a color using the HSL color model):
            // Specifies the actual angular value of the shift.The result of the shift shall be between 0
            // and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. Shifts
            // resulting in angular values greater than 360 are treated as 360.
            // [Example:
            //     The following increases the hue angular value by 10 degrees.
            //     < a: solidFill >
            //         <a:hslClr hue="0" sat="100%" lum="50%"/>
            //             <a:hueOff val="600000"/>
            //     </a: solidFill >
            //var hueOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
            // if (!isNaN(hueOff)) {
            //     //console.log("hueOff: ", hueOff, " (TODO)")
            //     //color = applyHueOff(color, hueOff, isAlpha);
            // }

            //16. "inv" (inverse)
            //This element specifies the inverse of its input color.
            //The inverse of red (1, 0, 0) is cyan (0, 1, 1 ).
            // The following represents cyan, the inverse of red:
            // <a:solidFill>
            //     <a:srgbClr val="FF0000">
            //         <a:inv />
            //     </a:srgbClr>
            // </a:solidFill>

            //17. "invGamma" - This element specifies that the output color rendered by the generating application should be the inverse sRGB
            //                  gamma shift of the input color.

            //18. "lum":
            // This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
            // Typically luminance values fall in the range[0 %, 100 %].
            // The following two solid fills are equivalent:
            // <a:solidFill>
            //     <a:hslClr hue="14400000" sat="100%" lum="50%">
            // </a:solidFill>
            // <a:solidFill>
            //     <a:hslClr hue="14400000" sat="100%" lum="0%">
            //         <a:lum val="50%" />
            //     <a:hslClr />
            // </a:solidFill>
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            // to value RRGGBB = (00, 66, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lum val="20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            //19. "lumMod":
            // Specifies the luminance as expressed by a percentage relative to the input color.
            // Increases never increase the luminance beyond 100 %, decreases never decrease the
            // luminance below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, 75, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lumMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            var lumMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
            //console.log("lumMod: ", lumMod)
            if (!isNaN(lumMod)) {
                color = applyLumMod(color, lumMod, isAlpha);
            }
            //var lumMod_color = applyLumMod(color, 0.5);
            //console.log("lumMod_color: ", lumMod_color)
            //20. "lumOff"
            // Specifies the luminance as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the luminance beyond 100 %, decreases never
            // decrease the luminance below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, 99, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:lumOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            var lumOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
            //console.log("lumOff: ", lumOff)
            if (!isNaN(lumOff)) {
                color = applyLumOff(color, lumOff, isAlpha);
            }


            //21. "red":
            // Specifies the value of the red component.The assigned value is specified as a percentage
            // with 0 % indicating minimal red and 100 % indicating maximum red.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (FF, FF, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:red val="100%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //22. "redMod":
            // Specifies the red component as expressed by a percentage relative to the input color
            // component.Increases never increase the red component beyond 100 %, decreases never
            // decrease the red component below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
            //     to value RRGGBB = (80, 00, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="FF0000">
            //             <a:redMod val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            //23. "redOff":
            // Specifies the red component as expressed by a percentage offset increase or decrease to
            // the input color component.Increases never increase the red component beyond 100 %,
            //     decreases never decrease the red component below 0 %.
            //     [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
            //     to value RRGGBB = (CC, 00, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="FF0000">
            //             <a:redOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >

            //23. "sat":
            // This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
            // Typically saturation values fall in the range[0 %, 100 %].
            // [Example:
            //     The following two solid fills are equivalent:
            //     <a:solidFill>
            //         <a:hslClr hue="14400000" sat="100%" lum="50%">
            //     </a:solidFill>
            //     <a:solidFill>
            //         <a:hslClr hue="14400000" sat="0%" lum="50%">
            //             <a:sat val="100000" />
            //         <a:hslClr />
            //     </a:solidFill>
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (40, C0, 40)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:sat val="50%" />
            //         </a:srgbClr>
            //     <a: solidFill >
            // end example]

            //24. "satMod":
            // Specifies the saturation as expressed by a percentage relative to the input color.
            // Increases never increase the saturation beyond 100 %, decreases never decrease the
            // saturation below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (66, 99, 66)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:satMod val="20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            var satMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
            if (!isNaN(satMod)) {
                color = applySatMod(color, satMod, isAlpha);
            }
            //25. "satOff":
            // Specifies the saturation as expressed by a percentage offset increase or decrease to the
            // input color.Increases never increase the saturation beyond 100 %, decreases never
            // decrease the saturation below 0 %.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (19, E5, 19)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:satOff val="-20%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // var satOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
            // if (!isNaN(satOff)) {
            //     console.log("satOff: ", satOff, " (TODO)")
            // }

            //26. "shade":
            // This element specifies a darker version of its input color.A 10 % shade is 10 % of the input color combined with 90 % black.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (00, BC, 00)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:shade val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            // end example]
            var shade = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
            if (!isNaN(shade)) {
                color = applyShade(color, shade, isAlpha);
            }
            //27.  "tint":
            // This element specifies a lighter version of its input color.A 10 % tint is 10 % of the input color combined with
            // 90 % white.
            // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
            //     to value RRGGBB = (BC, FF, BC)
            //     < a: solidFill >
            //         <a:srgbClr val="00FF00">
            //             <a:tint val="50%" />
            //         </a:srgbClr>
            //     </a: solidFill >
            var tint = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
            if (!isNaN(tint)) {
                color = applyTint(color, tint, isAlpha);
            }
            //console.log("color [%cfinal]: ", "color: #" + color, tinycolor(color).toHslString(), color)

            return color;
        }
        function toHex(n) {
            var hex = n.toString(16);
            while (hex.length < 2) { hex = "0" + hex; }
            return hex;
        }
        function hslToRgb(hue, sat, light) {
            var t1, t2, r, g, b;
            hue = hue / 60;
            if (light <= 0.5) {
                t2 = light * (sat + 1);
            } else {
                t2 = light + sat - (light * sat);
            }
            t1 = light * 2 - t2;
            r = hueToRgb(t1, t2, hue + 2) * 255;
            g = hueToRgb(t1, t2, hue) * 255;
            b = hueToRgb(t1, t2, hue - 2) * 255;
            return { r: r, g: g, b: b };
        }
        function hueToRgb(t1, t2, hue) {
            if (hue < 0) hue += 6;
            if (hue >= 6) hue -= 6;
            if (hue < 1) return (t2 - t1) * hue + t1;
            else if (hue < 3) return t2;
            else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
            else return t1;
        }
        function getColorName2Hex(name) {
            var hex;
            var colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
            var colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
            var findIndx = colorName.indexOf(name);
            if (findIndx != -1) {
                hex = colorHex[findIndx];
            }
            return hex;
        }
        function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
            //<p:clrMap ...> in slide master
            // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
            //console.log("getSchemeColorFromTheme: schemeClr: ", schemeClr, ",clrMap: ", clrMap)
            var slideLayoutClrOvride;
            if (clrMap !== undefined) {
                slideLayoutClrOvride = clrMap;//getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
            } else {
                var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    slideLayoutClrOvride = sldClrMapOvr;
                } else {
                    var sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        slideLayoutClrOvride = sldClrMapOvr;
                    } else {
                        slideLayoutClrOvride = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
            }
            //console.log("getSchemeColorFromTheme slideLayoutClrOvride: ", slideLayoutClrOvride);
            var schmClrName = schemeClr.substr(2);
            if (schmClrName == "phClr" && phClr !== undefined) {
                color = phClr;
            } else {
                if (slideLayoutClrOvride !== undefined) {
                    switch (schmClrName) {
                        case "tx1":
                        case "tx2":
                        case "bg1":
                        case "bg2":
                            schemeClr = "a:" + slideLayoutClrOvride[schmClrName];
                            break;
                    }
                } else {
                    switch (schmClrName) {
                        case "tx1":
                            schemeClr = "a:dk1";
                            break;
                        case "tx2":
                            schemeClr = "a:dk2";
                            break;
                        case "bg1":
                            schemeClr = "a:lt1";
                            break;
                        case "bg2":
                            schemeClr = "a:lt2";
                            break;
                    }
                }
                //console.log("getSchemeColorFromTheme:  schemeClr: ", schemeClr);
                var refNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
                var color = PPTXXmlUtils.getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
                //console.log("themeContent: color", color);
                if (color === undefined && refNode !== undefined) {
                    color = PPTXXmlUtils.getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
                }
            }
            //console.log(color)
            return color;
        }

        function extractChartData(serNode) {

            var dataMat = new Array();

            if (serNode === undefined) {
                return dataMat;
            }

            if (serNode["c:xVal"] !== undefined) {
                var dataRow = new Array();
                eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
                dataRow = new Array();
                eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push(parseFloat(innerNode["c:v"]));
                    return "";
                });
                dataMat.push(dataRow);
            } else {
                eachElement(serNode, function (innerNode, index) {
                    var dataRow = new Array();
                    var colName = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                    // Category (string or number)
                    var rowNames = {};
                    if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function (innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    } else if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                            rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                            return "";
                        });
                    }

                    // Value
                    if  (PPTXXmlUtils.getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                        eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                            dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                            return "";
                        });
                    }

                    dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
                    return "";
                });
            }

            return dataMat;
        }
        /**
         * setTextByPathList
         * @param {Object} node
         * @param {string Array} path
         * @param {string} value
         */
        function setTextByPathList(node, path, value) {

            if (path.constructor !== Array) {
                throw Error("Error of path type! path is not array.");
            }

            if (node === undefined) {
                return undefined;
            }

            Object.prototype.set = function (parts, value) {
                if(!parts) return this;
                //var parts = prop.split('.');
                var obj = this;
                var lent = parts.length;
                for (var i = 0; i < lent; i++) {
                    var p = parts[i];
                    if (obj[p] === undefined) {
                        if (i == lent - 1) {
                            obj[p] = value;
                        } else {
                            obj[p] = {};
                        }
                    }
                    obj = obj[p];
                }
                return obj;
            }

            node.set(path, value)
        }

        /**
         * eachElement
         * @param {Object} node
         * @param {function} doFunction
         */
        function eachElement(node, doFunction) {
            if (node === undefined) {
                return;
            }
            var result = "";
            if (node.constructor === Array) {
                var l = node.length;
                for (var i = 0; i < l; i++) {
                    result += doFunction(node[i], i);
                }
            } else {
                result += doFunction(node, 0);
            }
            return result;
        }

        // ===== Color functions =====
        /**
         * applyShade
         * @param {string} rgbStr
         * @param {number} shadeValue
         */
        function applyShade(rgbStr, shadeValue, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applyShade  color: ", color, ", shadeValue: ", shadeValue)
            // 确保shadeValue在0-1之间
            shadeValue = Math.max(0, Math.min(1, shadeValue));
            // PPTX标准：Shade = L * shadeValue
            var cacl_l = Math.max(0, Math.min(1, color.l * shadeValue));
            if (isAlpha)
                return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }

        /**
         * applyTint
         * @param {string} rgbStr
         * @param {number} tintValue
         */
        function applyTint(rgbStr, tintValue, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applyTint  color: ", color, ", tintValue: ", tintValue)
            // 确保tintValue在0-1之间
            tintValue = Math.max(0, Math.min(1, tintValue));
            // PPTX标准：Tint = L * tintValue + (1 - tintValue)
            var cacl_l = Math.max(0, Math.min(1, color.l * tintValue + (1 - tintValue)));
            if (isAlpha)
                return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }

        /**
         * applyLumOff
         * @param {string} rgbStr
         * @param {number} offset
         */
        function applyLumOff(rgbStr, offset, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applyLumOff  color.l: ", color.l, ", offset: ", offset, ", color.l + offset : ", color.l + offset)
            var lum = offset + color.l;
            if (lum >= 1) {
                if (isAlpha)
                    return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
                return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
            }
            if (isAlpha)
                return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
        }

        /**
         * applyLumMod
         * @param {string} rgbStr
         * @param {number} multiplier
         */
        function applyLumMod(rgbStr, multiplier, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applyLumMod  color.l: ", color.l, ", multiplier: ", multiplier, ", color.l * multiplier : ", color.l * multiplier)
            var cacl_l = color.l * multiplier;
            if (cacl_l >= 1) {
                cacl_l = 1;
            }
            if (isAlpha)
                return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
        }


        // /**
        //  * applyHueMod
        //  * @param {string} rgbStr
        //  * @param {number} multiplier
        //  */
        function applyHueMod(rgbStr, multiplier, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applyLumMod  color.h: ", color.h, ", multiplier: ", multiplier, ", color.h * multiplier : ", color.h * multiplier)

            var cacl_h = color.h * multiplier;
            if (cacl_h >= 360) {
                cacl_h = cacl_h - 360;
            }
            if (isAlpha)
                return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
            return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
        }


        // /**
        //  * applyHueOff
        //  * @param {string} rgbStr
        //  * @param {number} offset
        //  */
        // function applyHueOff(rgbStr, offset, isAlpha) {
        //     var color = tinycolor(rgbStr).toHsl();
        //     //console.log("applyLumMod  color.h: ", color.h, ", offset: ", offset, ", color.h * offset : ", color.h * offset)

        //     var cacl_h = color.h * offset;
        //     if (cacl_h >= 360) {
        //         cacl_h = cacl_h - 360;
        //     }
        //     if (isAlpha)
        //         return tinycolor({ h: cocacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
        //     return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
        // }
        // /**
        //  * applySatMod
        //  * @param {string} rgbStr
        //  * @param {number} multiplier
        //  */
        function applySatMod(rgbStr, multiplier, isAlpha) {
            var color = tinycolor(rgbStr).toHsl();
            //console.log("applySatMod  color.s: ", color.s, ", multiplier: ", multiplier, ", color.s * multiplier : ", color.s * multiplier)
            var cacl_s = color.s * multiplier;
            if (cacl_s >= 1) {
                cacl_s = 1;
            }
            //return;
            // if (isAlpha)
            //     return tinycolor(rgbStr).saturate(multiplier * 100).toHex8();
            // return tinycolor(rgbStr).saturate(multiplier * 100).toHex();
            if (isAlpha)
                return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
            return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
        }

        /**
         * rgba2hex
         * @param {string} rgbaStr
         */
        function rgba2hex(rgbaStr) {
            var a,
                rgb = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i),
                alpha = (rgb && rgb[4] || "").trim(),
                hex = rgb ?
                    (rgb[1] | 1 << 8).toString(16).slice(1) +
                    (rgb[2] | 1 << 8).toString(16).slice(1) +
                    (rgb[3] | 1 << 8).toString(16).slice(1) : rgbaStr;

            if (alpha !== "") {
                a = alpha;
            } else {
                a = 0o1;
            }
            // multiply before convert to HEX
            a = ((a * 255) | 1 << 8).toString(16).slice(1)
            hex = hex + a;

            return hex;
        }
        // function degreesToRadians(degrees) {
        //     //Math.PI
        //     if (degrees == "" || degrees == null || degrees == undefined) {
        //         return 0;
        //     }
        //     return degrees * (Math.PI / 180);
        // }
        
        function getSvgGradient(w, h, angl, color_arry, shpId) {
            var stopsArray = getMiddleStops(color_arry - 2);

            var svgAngle = '',
                svgHeight = h,
                svgWidth = w,
                svg = '',
                xy_ary = SVGangle(angl, svgHeight, svgWidth),
                x1 = xy_ary[0],
                y1 = xy_ary[1],
                x2 = xy_ary[2],
                y2 = xy_ary[3];

            var sal = stopsArray.length,
                sr = sal < 20 ? 100 : 1000;
            svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
            svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
            svg += svgAngle;

            for (var i = 0; i < sal; i++) {
                var tinClr = tinycolor("#" + color_arry[i]);
                var alpha = tinClr.getAlpha();
                //console.log("color: ", color_arry[i], ", rgba: ", tinClr.toHexString(), ", alpha: ", alpha)
                svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
                svg += '/>\n'
            }

            svg += '</linearGradient>\n' + '';

            return svg
        }
        function getMiddleStops(s) {
            var sArry = ['0%', '100%'];
            if (s == 0) {
                return sArry;
            } else {
                var i = s;
                while (i--) {
                    var middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                        middleStopString = middleStop + "%";
                    sArry.splice(-1, 0, middleStopString);
                } // AM: add into stopsArray before 100%
            }
            return sArry
        }
        function SVGangle(deg, svgHeight, svgWidth) {
            var w = parseFloat(svgWidth),
                h = parseFloat(svgHeight),
                ang = parseFloat(deg),
                o = 2,
                n = 2,
                wc = w / 2,
                hc = h / 2,
                tx1 = 2,
                ty1 = 2,
                tx2 = 2,
                ty2 = 2,
                k = (((ang % 360) + 360) % 360),
                j = (360 - k) * Math.PI / 180,
                i = Math.tan(j),
                l = hc - i * wc;

            if (k == 0) {
                tx1 = w,
                    ty1 = hc,
                    tx2 = 0,
                    ty2 = hc
            } else if (k < 90) {
                n = w,
                    o = 0
            } else if (k == 90) {
                tx1 = wc,
                    ty1 = 0,
                    tx2 = wc,
                    ty2 = h
            } else if (k < 180) {
                n = 0,
                    o = 0
            } else if (k == 180) {
                tx1 = 0,
                    ty1 = hc,
                    tx2 = w,
                    ty2 = hc
            } else if (k < 270) {
                n = 0,
                    o = h
            } else if (k == 270) {
                tx1 = wc,
                    ty1 = h,
                    tx2 = wc,
                    ty2 = 0
            } else {
                n = w,
                    o = h;
            }
            // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
            var m = o + (n / i),
                tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
                ty1 = ty1 == 2 ? i * tx1 + l : ty1,
                tx2 = tx2 == 2 ? w - tx1 : tx2,
                ty2 = ty2 == 2 ? h - ty1 : ty2,
                x1 = Math.round(tx2 / w * 100 * 100) / 100,
                y1 = Math.round(ty2 / h * 100 * 100) / 100,
                x2 = Math.round(tx1 / w * 100 * 100) / 100,
                y2 = Math.round(ty1 / h * 100 * 100) / 100;
            return [x1, y1, x2, y2];
        }
        function getSvgImagePattern(node, fill, shpId, warpObj) {
            var pic_dim = getBase64ImageDimensions(fill);
            var width = pic_dim[0];
            var height = pic_dim[1];
            //console.log("getSvgImagePattern node:", node);
            var blipFillNode = node["p:spPr"]["a:blipFill"];
            var tileNode = PPTXXmlUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"])
            if (tileNode !== undefined && tileNode["sx"] !== undefined) {
                var sx = (parseInt(tileNode["sx"]) / 100000) * width;
                var sy = (parseInt(tileNode["sy"]) / 100000) * height;
            }

            var blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
            var tialphaModFixNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"])
            var imgOpacity = "";
            if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
                var amt = parseInt(tialphaModFixNode["amt"]) / 100000;
                var opacity = amt;
                var imgOpacity = "opacity='" + opacity + "'";

            }
            if (sx !== undefined && sx != 0) {
                var ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
            } else {
                var ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
            }
            var duotoneNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:duotone"])
            var fillterNode = "";
            var filterUrl = "";
            if (duotoneNode !== undefined) {
                //console.log("pic duotoneNode: ", duotoneNode)
                var clr_ary = [];
                Object.keys(duotoneNode).forEach(function (clr_type) {
                    //Object.keys(duotoneNode[clr_type]).forEach(function (clr) {
                    //console.log("blip pic duotone clr: ", duotoneNode[clr_type][clr], clr)
                    if (clr_type != "attrs") {
                        var obj = {};
                        obj[clr_type] = duotoneNode[clr_type];
                        //console.log("blip pic duotone obj: ", obj)
                        var hexClr = getSolidFill(obj, undefined, undefined, warpObj)
                        //clr_ary.push();

                        var color = tinycolor("#" + hexClr);
                        clr_ary.push(color.toRgb()); // { r: 255, g: 0, b: 0, a: 1 }
                    }
                    // })
                })

                if (clr_ary.length == 2) {

                    fillterNode = '<filter id="svg_image_duotone"> ' +
                        '<feColorMatrix type="matrix" values=".33 .33 .33 0 0' +
                        '.33 .33 .33 0 0' +
                        '.33 .33 .33 0 0' +
                        '0 0 0 1 0">' +
                        '</feColorMatrix>' +
                        '<feComponentTransfer color-interpolation-filters="sRGB">' +
                        //clr_ary.forEach(function(clr){
                        '<feFuncR type="table" tableValues="' + clr_ary[0].r / 255 + ' ' + clr_ary[1].r / 255 + '"></feFuncR>' +
                        '<feFuncG type="table" tableValues="' + clr_ary[0].g / 255 + ' ' + clr_ary[1].g / 255 + '"></feFuncG>' +
                        '<feFuncB type="table" tableValues="' + clr_ary[0].b / 255 + ' ' + clr_ary[1].b / 255 + '"></feFuncB>' +
                        //});
                        '</feComponentTransfer>' +
                        ' </filter>';
                }

                filterUrl = 'filter="url(#svg_image_duotone)"';

                ptrn += fillterNode;
            }

            fill = PPTXXmlUtils.escapeHtml(fill);
            if (sx !== undefined && sx != 0) {
                ptrn += '<image  xlink:href="' + fill + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
            } else {
                ptrn += '<image  xlink:href="' + fill + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
            }
            ptrn += '</pattern>';

            //console.log("getSvgImagePattern(...) pic_dim:", pic_dim, ", fillColor: ", fill, ", blipNode: ", blipNode, ",sx: ", sx, ", sy: ", sy, ", clr_ary: ", clr_ary, ", ptrn: ", ptrn)

            return ptrn;
        }

        function getBase64ImageDimensions(imgSrc) {
            var image = new Image();
            var w, h;
            image.onload = function () {
                w = image.width;
                h = image.height;
            };
            image.src = imgSrc;

            do {
                if (image.width !== undefined) {
                    return [image.width, image.height];
                }
            } while (image.width === undefined);

            //return [w, h];
        }

        function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {

            //X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
            var anchor = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
            //console.log("getVerticalAlign anchor:", anchor, "slideLayoutSpNode: ", slideLayoutSpNode)
            if (anchor === undefined) {
                //console.log("getVerticalAlign type:", type," node:", node, "slideLayoutSpNode:", slideLayoutSpNode, "slideMasterSpNode:", slideMasterSpNode)
                anchor = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                if (anchor === undefined) {
                    anchor = PPTXXmlUtils.getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    if (anchor === undefined) {
                        //"If this attribute is omitted, then a value of t, or top is implied."
                        anchor = "t";//getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
                    }
                }
            }
            //console.log("getVerticalAlign:", node, slideLayoutSpNode, slideMasterSpNode, type, anchor)
            return (anchor === "ctr")?"v-mid" : ((anchor === "b") ? "v-down" : "v-up");
        }

    function getContentDir(node, type, warpObj) {
            return "content";
            var defRtl = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "a:defPPr", "attrs", "rtl"]);
            if (defRtl !== undefined) {
                if (defRtl == "1"){
                    return "content-rtl";
                } else if (defRtl == "0") {
                    return "content";
                }
            }
            //var lvl1Rtl = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "lvl1pPr", "attrs", "rtl"]);
            // if (lvl1Rtl !== undefined) {
            //     if (lvl1Rtl == "1") {
            //         return "content-rtl";
            //     } else if (lvl1Rtl == "0") {
            //         return "content";
            //     }
            // }
            var rtlCol = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "rtlCol"]);
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
                var dirVal = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles[dirLoc], ["a:lvl1pPr", "attrs", "rtl"]);
                if (dirVal == "1") {
                    return "content-rtl";
                }
            } 
            // else {
            //     if (type == "textBox") {
            //         var dirVal = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", "a:lvl1pPr", "attrs", "rtl"]);
            //         if (dirVal == "1") {
            //             return "content-rtl";
            //         }
            //     }
            // }
            return "content";
            //console.log("getContentDir() type:", type, "slideMasterTextStyles:", slideMasterTextStyles,"dirNode:",dirVal)
        }

        function getVerticalMargins(pNode, textBodyNode, type, idx, warpObj) {
            //margin-top ; 
            //a:pPr => a:spcBef => a:spcPts (/100) | a:spcPct (/?)
            //margin-bottom
            //a:pPr => a:spcAft => a:spcPts (/100) | a:spcPct (/?)
            //+
            //a:pPr =>a:lnSpc => a:spcPts (/?) | a:spcPct (/?)
            //console.log("getVerticalMargins ", pNode, type,idx, warpObj)
            //var lstStyle = textBodyNode["a:lstStyle"];
            var lvl = 1
            var spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
            var spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
            var lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
            var lnSpcNodeType = "Pct";
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                if (lnSpcNode !== undefined) {
                    lnSpcNodeType = "Pts";
                }
            }
            var lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            var fontSize;
            if  (PPTXXmlUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
                var fontSizeStr = getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
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
            //     spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPct","attrs","val"]);
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
            //     spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPct","attrs","val"]);
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
                    var laypPrNode = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

                    if (spcBefNode === undefined) {
                        spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                        // if(spcBefNode !== undefined){
                        //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                        // } 
                        // else{
                        //    //i did not found case with percentage 
                        //     spcBefNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                        //     if(spcBefNode !== undefined){
                        //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                        //     }
                        // }
                    }

                    if (spcAftNode === undefined) {
                        spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                        // if(spcAftNode !== undefined){
                        //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                        // }
                        // else{
                        //    //i did not found case with percentage 
                        //     spcAftNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                        //     if(spcAftNode !== undefined){
                        //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                        //     }
                        // }
                    }

                    if (lnSpcNode === undefined) {
                        lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                        if (lnSpcNode === undefined) {
                            lnSpcNode = PPTXXmlUtils.getTextByPathList(laypPrNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
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
                var inLvlNode = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
                if (inLvlNode !== undefined) {
                    if (spcBefNode === undefined) {
                        spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPts", "attrs", "val"]);
                        // if(spcBefNode !== undefined){
                        //     spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "pt;"
                        // } 
                        // else{
                        //    //i did not found case with percentage 
                        //     spcBefNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcBef", "a:spcPct","attrs","val"]);
                        //     if(spcBefNode !== undefined){
                        //         spcBef = "margin-top:" + parseInt(spcBefNode)/100 + "%;"
                        //     }
                        // }
                    }

                    if (spcAftNode === undefined) {
                        spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPts", "attrs", "val"]);
                        // if(spcAftNode !== undefined){
                        //     spcAft = "margin-bottom:" + parseInt(spcAftNode)/100 + "pt;"
                        // }
                        // else{
                        //    //i did not found case with percentage 
                        //     spcAftNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:spcAft", "a:spcPct","attrs","val"]);
                        //     if(spcAftNode !== undefined){
                        //         spcBef = "margin-bottom:" + parseInt(spcAftNode)/100 + "%;"
                        //     }
                        // }
                    }

                    if (lnSpcNode === undefined) {
                        lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:lnSpc", "a:spcPct", "attrs", "val"]);
                        if (lnSpcNode === undefined) {
                            lnSpcNode = PPTXXmlUtils.getTextByPathList(inLvlNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
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
        }
        function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj) {
            var algn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (algn === undefined) {
                //var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                // var pPrNodeLaout = layoutMasterNode.nodeLaout;
                // var pPrNodeMaster = layoutMasterNode.nodeMaster;
                var lvlIdx = 1;
                var lvlNode = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
                if (lvlNode !== undefined) {
                    lvlIdx = parseInt(lvlNode) + 1;
                }
                var lvlStr = "a:lvl" + lvlIdx + "pPr";

                var lstStyle = textBodyNode["a:lstStyle"];
                algn = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);

                if (algn === undefined && idx !== undefined ) {
                    //slidelayout
                    algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                    if (algn === undefined) {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                        if (algn === undefined) {
                            algn = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                        }
                    }
                }
                if (algn === undefined) {
                    if (type !== undefined) {
                        //slidelayout
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);

                        if (algn === undefined) {
                            //masterlayout
                            if (type == "title" || type == "ctrTitle") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "body" || type == "obj" || type == "subTitle") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "shape" || type == "diagram") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                            } else if (type == "textBox") {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                            } else {
                                algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                            }
                        }
                    } else {
                        algn = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
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
        }

        function getLayoutAndMasterNode(node, idx, type, warpObj) {
            var pPrNodeLaout, pPrNodeMaster;
            var pPrNode = node["a:pPr"];
            //lvl
            var lvl = 1;
            var lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            if (idx !== undefined) {
                //slidelayout
                pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
                    if (pPrNodeLaout === undefined) {
                        pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
                    }
                }
            }
            if (type !== undefined) {
                //slidelayout
                var lvlStr = "a:lvl" + lvl + "pPr";
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
                }
                //masterlayout
                if (type == "title" || type == "ctrTitle") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
                } else if (type == "body" || type == "obj" || type == "subTitle") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
                } else if (type == "shape" || type == "diagram") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
                } else if (type == "textBox") {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
                } else {
                    pPrNodeMaster = PPTXXmlUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
                }
            }
            return {
                "nodeLaout": pPrNodeLaout,
                "nodeMaster": pPrNodeMaster
            };
        }
    function getPregraphDir(node, textBodyNode, idx, type, warpObj) {
            var rtl = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
            //console.log("getPregraphDir node:", node, "textBodyNode", textBodyNode, "rtl:", rtl, "idx", idx, "type", type, "warpObj", warpObj)
          

            if (rtl === undefined) {
                var layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                var pPrNodeLaout = layoutMasterNode.nodeLaout;
                var pPrNodeMaster = layoutMasterNode.nodeMaster;
                rtl = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (rtl === undefined && type != "shape") {
                    rtl = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }

            if (rtl == "1") {
                return "pregraph-rtl";
            } else if (rtl == "0") {
                return "pregraph-ltr";
            }
            return "pregraph-inherit";

            // var contentDir = PPTXStyleUtils.getContentDir(type, warpObj);
            // console.log("getPregraphDir node:", node["a:r"], "rtl:", rtl, "idx", idx, "type", type, "contentDir:", contentDir)

            // if (contentDir == "content"){
            //     return "pregraph-ltr";
            // } else if (contentDir == "content-rtl"){ 
            //     return "pregraph-rtl";
            // }
            // return "";
        }
    function getPregraphMargn(pNode, idx, type, isBullate, warpObj){
            if (!isBullate){
                return ["",0];
            }
            var marLStr = "", marRStr = "" , maginVal = 0;
            var pPrNode = pNode["a:pPr"];
            var layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;
            
            // var lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
            // var isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1) ? true : false;
            //rtl
            var getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            var isRTL = false;
            var dirStr = "ltr";
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
                dirStr = "rtl";
            }

            //align
            var alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
            if (alignNode === undefined) {
                alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                if (alignNode === undefined) {
                    alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }
            //indent?
            var indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
                if (indentNode === undefined) {
                    indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
                }
            }
            var indent = 0;
            if (indentNode !== undefined) {
                indent = parseInt(indentNode) * slideFactor;
            }
            //
            //marL
            var marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
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
            var marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
            if (marRNode === undefined && marLNode === undefined) {
                //need to check if this posble - TODO
                marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
                if (marRNode === undefined) {
                    marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
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
        }
    return {
        getFillType,
        getShapeFill,
        getFontType,
        getFontColorPr,
        getFontSize,
        getFontBold,
        getFontItalic,
        getFontDecoration,
        getTextHorizontalAlign,
        getTextVerticalAlign,
        getTableBorders,
        getBorder,
        getBackground,
        getSlideBackgroundFill,
        getBgGradientFill,
        getBgPicFill,
        getGradientFill,
        getPicFill,
        getPatternFill,
        getLinerGrandient,
        getSolidFill,
        toHex,
        hslToRgb,
        hueToRgb,
        getColorName2Hex,
        getSchemeColorFromTheme,
        extractChartData,
        setTextByPathList,
        eachElement,
        applyShade,
        applyTint,
        applyLumOff,
        applyLumMod,
        applyHueMod,
        applySatMod,
        rgba2hex,
        getSvgGradient,
        getMiddleStops,
        SVGangle,
        getSvgImagePattern,
        getBase64ImageDimensions,
        getVerticalAlign,
        getContentDir,
        getHorizontalAlign,
        getVerticalMargins,
        getLayoutAndMasterNode,
        getPregraphDir,
        getPregraphMargn,
    };
})();

window.PPTXStyleUtils = PPTXStyleUtils;
