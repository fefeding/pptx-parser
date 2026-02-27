/**
 * 样式处理模块
 * 
 * 处理 PPTX 文件中的各种样式属性，包括：
 * - 填充类型（纯色、渐变、图片、图案等）
 * - 边框样式
 * - 阴影效果
 * - 3D 效果
 * - 反射效果
 * 
 * @module utils/style
 */

import { PPTXXmlUtils } from './xml.js';
import { SLIDE_FACTOR, FONT_SIZE_FACTOR, RTL_LANGS_ARRAY } from '../core/constants.js';
import tinycolor from '../core/tinycolor.js';



function getFillType(node) {
            //Need to test/////////////////////////////////////////////
            //SOLID_FILL
            //PIC_FILL
            //GRADIENT_FILL
            //PATTERN_FILL
            //NO_FILL
            let fillType = "";
            if (node === undefined) {
                return fillType;
            }
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
        //     let arrBuff = new ArrayBuffer(4);
        //     let vw = new DataView(arrBuff);
        //     vw.setUint32(0, parseInt(hex, 16), false);
        //     let arrByte = new Uint8Array(arrBuff);
        //     return arrByte[1] + "," + arrByte[2] + "," + arrByte[3];
        // }
        function getShapeFill(node, pNode, isSvgMode, warpObj, source) {

            // 1. presentationML
            // p:spPr/ [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
            // From slide
            //Fill Type:
            //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode)
            let fillType = getFillType (PPTXXmlUtils.getTextByPathList(node, ["p:spPr"]));
            //let noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            let fillColor;
            if (fillType === "NO_FILL") {
                return isSvgMode ? "none" : "";
            } else if (fillType === "SOLID_FILL") {
                let shpFill = node["p:spPr"]["a:solidFill"];
                fillColor = getSolidFill(shpFill, undefined, undefined, warpObj);
            } else if (fillType === "GRADIENT_FILL") {
                let shpFill = node["p:spPr"]["a:gradFill"];
                fillColor = getGradientFill(shpFill, warpObj);
            } else if (fillType === "PATTERN_FILL") {
                let shpFill = node["p:spPr"]["a:pattFill"];
                fillColor = getPatternFill(shpFill, warpObj);
            } else if (fillType === "PIC_FILL") {
                let shpFill = node["p:spPr"]["a:blipFill"];
                fillColor = getPicFill(source, shpFill, warpObj);
            }
            //console.log("getShapeFill ShapeFill: ", node, ", isSvgMode; ", isSvgMode, ", fillType: ", fillType, ", fillColor: ", fillColor, ", source: ", source)


            // 2. drawingML namespace
            if (fillColor === undefined) {
                let clrName = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef"]);
                let idx = parseInt (PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:fillRef", "attrs", "idx"]));
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
                let grpFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:grpFill"]);
                if (grpFill !== undefined) {
                    //fillColor = getSolidFill(clrName, undefined, undefined, undefined, warpObj);
                    //get parent fill style - TODO
                    //console.log("ShapeFill: grpFill: ", grpFill, ", pNode: ", pNode)
                    let grpShpFill = pNode["p:grpSpPr"];
                    let spShpNode = { "p:spPr": grpShpFill }
                    return getShapeFill(spShpNode, node, isSvgMode, warpObj, source);
                } else if (fillType === "NO_FILL") {
                    return isSvgMode ? "none" : "";
                }
            }
            //console.log("ShapeFill: fillColor: ", fillColor, ", fillType; ", fillType)

            if (fillColor !== undefined) {
                if (fillType === "GRADIENT_FILL") {
                    if (isSvgMode) {
                        // console.log("GRADIENT_FILL color", fillColor.color[0])
                        return fillColor;
                    } else {
                        let colorAry = fillColor.color;
                        let rot = fillColor.rot;

                        let bgcolor = `background: linear-gradient(${rot}deg,`;
                        for (let i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bgcolor += "#" + colorAry[i] + ");";
                            } else {
                                bgcolor += "#" + colorAry[i] + ", ";
                            }

                        }
                        return bgcolor;
                    }
                } else if (fillType === "PIC_FILL") {
                    if (isSvgMode) {
                        // 当 isSvgMode 为 true 时，返回图像 URL 而不是整个对象
                        if (typeof fillColor === 'object' && fillColor.img) {
                            return fillColor.img;
                        } else {
                            return fillColor;
                        }
                    } else {
                        if (typeof fillColor === 'object' && fillColor.img) {
                            return `background-image:url(${fillColor.img}); background-size: ${fillColor.backgroundSize}; background-position: ${fillColor.backgroundPosition}; background-repeat: ${fillColor.backgroundRepeat};`;
                        } else {
                            return `background-image:url(${fillColor});`;
                        }
                    }
                } else if (fillType === "PATTERN_FILL") {
                    /////////////////////////////////////////////////////////////Need to check -----------TODO
                    // if (isSvgMode) {
                    //     let color = tinycolor(fillColor);
                    //     fillColor = color.toRgbString();

                    //     return fillColor;
                    // } else {
                    let bgPtrn = "", bgSize = "", bgPos = "";
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
                        let color = tinycolor(fillColor);
                        fillColor = color.toRgbString();

                        return fillColor;
                    } else {
                        //console.log(node,"fillColor: ",fillColor,"fillType: ",fillType,"isSvgMode: ",isSvgMode)
                        return `background-color: #${fillColor};`;
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
            let typeface = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);

            if (typeface === undefined) {
                let fontIdx = "";
                let fontGrup = "";
                if (pFontStyle !== undefined) {
                    fontIdx = PPTXXmlUtils.getTextByPathList(pFontStyle, ["attrs", "idx"]);
                }
                let fontSchemeNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fontScheme"]);
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
            let rPrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr"]);
            let filTyp, color, textBordr, colorType = "", highlightColor = "";
            //console.log("getFontColorPr type:", type, ", node: ", node, "pNode:", pNode, "pFontStyle:", pFontStyle)
            if (rPrNode !== undefined) {
                filTyp = getFillType(rPrNode);
                if (filTyp == "SOLID_FILL") {
                    let solidFillNode = rPrNode["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    let highlightNode = rPrNode["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    let pattFill = rPrNode["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = getBgPicFill(rPrNode, "slideBg", warpObj, undefined, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    let shpFill = rPrNode["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                } 
            }
            if (color === undefined && PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]) !== undefined) {
                //lstStyle
                let lstStyledefRPr = PPTXXmlUtils.getTextByPathList(lstStyle, ["a:lvl" + lvl + "pPr", "a:defRPr"]);
                filTyp = getFillType(lstStyledefRPr);
                if (filTyp == "SOLID_FILL") {
                    let solidFillNode = lstStyledefRPr["a:solidFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                    color = getSolidFill(solidFillNode, undefined, undefined, warpObj);
                    let highlightNode = lstStyledefRPr["a:highlight"];
                    if (highlightNode !== undefined) {
                        highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                    }
                    colorType = "solid";
                } else if (filTyp == "PATTERN_FILL") {
                    let pattFill = lstStyledefRPr["a:pattFill"];// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:pattFill"]);
                    color = getPatternFill(pattFill, warpObj);
                    colorType = "pattern";
                } else if (filTyp == "PIC_FILL") {
                    color = getBgPicFill(lstStyledefRPr, "slideBg", warpObj, undefined, undefined);
                    //color = getPicFill("slideBg", rPrNode["a:blipFill"], warpObj);
                    colorType = "pic";
                } else if (filTyp == "GRADIENT_FILL") {
                    let shpFill = lstStyledefRPr["a:gradFill"];
                    color = getGradientFill(shpFill, warpObj);
                    colorType = "gradient";
                }

            }
            if (color === undefined) {
                let sPstyle = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:fontRef"]);
                if (sPstyle !== undefined) {
                    color = getSolidFill(sPstyle, undefined, undefined, warpObj);
                    if (color !== undefined) {
                        colorType = "solid";
                    }
                    let highlightNode = sPstyle["a:highlight"]; //is "a:highlight" node in 'a:fontRef' ?
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
                let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;

                if (pPrNodeLaout !== undefined) {
                    let defRpRLaout = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:solidFill"]);
                    if (defRpRLaout !== undefined) {
                        color = getSolidFill(defRpRLaout, undefined, undefined, warpObj);
                        let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "a:highlight"]);
                        if (highlightNode !== undefined) {
                            highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                        }
                        colorType = "solid";
                    }
                }
                if (color === undefined) {
                    if (pPrNodeMaster !== undefined) {
                        let defRprMaster = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:solidFill"]);
                        if (defRprMaster !== undefined) {
                            color = getSolidFill(defRprMaster, undefined, undefined, warpObj);
                            let highlightNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "a:highlight"]);
                            if (highlightNode !== undefined) {
                                highlightColor = getSolidFill(highlightNode, undefined, undefined, warpObj);
                            }
                            colorType = "solid";
                        }
                    }
                }
            }
            let txtEffects = [];
            let txtEffObj = {}
            //textBordr
            let txtBrdrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:ln"]);
            textBordr = "";
            if (txtBrdrNode !== undefined && txtBrdrNode["a:noFill"] === undefined) {
                let txBrd = getBorder(node, pNode, false, "text", warpObj);
                let txBrdAry = txBrd.split(" ");
                //let brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("pt")))) + "px";
                let brdSize = (parseInt(txBrdAry[0].substring(0, txBrdAry[0].indexOf("px")))) + "px";
                let brdClr = txBrdAry[2];
                //let brdTyp = txBrdAry[1]; //not in use
                //console.log("getFontColorPr txBrdAry:", txBrdAry)
                if (colorType == "solid") {
                    textBordr = `-${brdSize} 0 ${brdClr}, 0 ${brdSize} ${brdClr}, ${brdSize} 0 ${brdClr}, 0 -${brdSize} ${brdClr}`;
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
            let txtGlowNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:glow"]);
            let oGlowStr = "";
            if (txtGlowNode !== undefined) {
                let glowClr = getSolidFill(txtGlowNode, undefined, undefined, warpObj);
                let rad = (txtGlowNode["attrs"]["rad"]) ? (txtGlowNode["attrs"]["rad"] * SLIDE_FACTOR) : 0;
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
            // Check for direct shadow effect in text run properties
            let txtShadow = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:effectLst", "a:outerShdw"]);
            let oShadowStr = "";
            
            // If no direct shadow, check effectRef from p:style
            if (txtShadow === undefined) {
                var effectRefNode = PPTXXmlUtils.getTextByPathList(pNode, ["p:style", "a:effectRef"]);
                if (effectRefNode !== undefined) {
                    var effectIdx = PPTXXmlUtils.getTextByPathList(effectRefNode, ["attrs", "idx"]);
                    if (effectIdx !== undefined && warpObj["themeContent"] !== undefined) {
                        // Access the effect style from the theme
                        var effectStyleLst = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:fmtScheme", "a:effectStyleLst", "a:effectStyle"]);
                        if (effectStyleLst !== undefined) {
                            // Ensure effectStyleLst is an array
                            if (!Array.isArray(effectStyleLst)) {
                                effectStyleLst = [effectStyleLst];
                            }
                            var idx = Number(effectIdx); // idx is 0-based, not 1-based
                            if (idx >= 0 && effectStyleLst[idx] !== undefined) {
                                txtShadow = PPTXXmlUtils.getTextByPathList(effectStyleLst[idx], ["a:effectLst", "a:outerShdw"]);
                            }
                        }
                    }
                }
            }
            
            if (txtShadow !== undefined) {
                //https://developer.mozilla.org/en-US/docs/Web/CSS/filter-function/drop-shadow()
                //https://stackoverflow.com/questions/60468487/css-text-with-linear-gradient-shadow-and-text-outline
                //https://css-tricks.com/creating-playful-effects-with-css-text-shadows/
                //https://designshack.net/articles/css/12-fun-css-text-shadows-you-can-copy-and-paste/

                let shadowClr = getSolidFill(txtShadow, undefined, undefined, warpObj);
                let outerShdwAttrs = txtShadow["attrs"];
                // algn: "bl"
                // dir: "2640000"
                // dist: "38100"
                // rotWithShape: "0/1" - Specifies whether the shadow rotates with the shape if the shape is rotated.
                //blurRad (Blur Radius) - Specifies the blur radius of the shadow.
                //kx (Horizontal Skew) - Specifies the horizontal skew angle.
                //ky (Vertical Skew) - Specifies the vertical skew angle.
                //sx (Horizontal Scaling Factor) - Specifies the horizontal scaling SLIDE_FACTOR; negative scaling causes a flip.
                //sy (Vertical Scaling Factor) - Specifies the vertical scaling SLIDE_FACTOR; negative scaling causes a flip.
                let algn = outerShdwAttrs["algn"];
                let dir = (outerShdwAttrs["dir"]) ? (parseInt(outerShdwAttrs["dir"]) / 60000) : 0;
                let dist = parseInt(outerShdwAttrs["dist"]) * SLIDE_FACTOR;//(px) //* (3 / 4); //(pt)
                let rotWithShape = outerShdwAttrs["rotWithShape"];
                let blurRad = (outerShdwAttrs["blurRad"]) ? (parseInt(outerShdwAttrs["blurRad"]) * SLIDE_FACTOR + "px") : "";
                let sx = (outerShdwAttrs["sx"]) ? (parseInt(outerShdwAttrs["sx"]) / 100000) : 1;
                let sy = (outerShdwAttrs["sy"]) ? (parseInt(outerShdwAttrs["sy"]) / 100000) : 1;
                let vx = dist * Math.sin(dir * Math.PI / 180);
                let hx = dist * Math.cos(dir * Math.PI / 180);

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
            let text_effcts = "", txt_effects;
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
            let lstStyle = (textBodyNode !== undefined)? textBodyNode["a:lstStyle"] : undefined;
            let lvlpPr = "a:lvl" + lvl + "pPr";
            let fontSize = undefined;
            let sz, kern;
            if (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"] && node["a:rPr"]["attrs"]["sz"] !== undefined) {
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
            let isAutoFit = false;
            let isKerning = false;
            if (textBodyNode !== undefined){
                let spAutoFitNode = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:spAutoFit"]);
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
                if (isKerning && kern !== undefined && !isNaN(fontSize) && kern > 0) {
                    // 字距调整不应该减小字体大小，而是调整字符间距
                    // 移除字体大小的减少
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
                if (isKerning && kern !== undefined && !isNaN(fontSize) && kern > 0) {
                    // 字距调整不应该减小字体大小，而是调整字符间距
                    // 移除字体大小的减少
                }
            }

            let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            if (baseline !== undefined && !isNaN(fontSize)) {
                let baselineVl = parseInt(baseline) / 100000;
                //fontSize -= 10; 
                // fontSize = fontSize * baselineVl;
                fontSize -= baselineVl;
            }

            // 如果仍然没有找到字体大小，尝试从段落的默认样式中获取
            if (isNaN(fontSize) || fontSize === undefined) {
                let pPrNode = node.parentNode && node.parentNode["a:pPr"];
                if (pPrNode) {
                    let defRPrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:defRPr"]);
                    if (defRPrNode && defRPrNode["attrs"] && defRPrNode["attrs"]["sz"]) {
                        fontSize = parseInt(defRPrNode["attrs"]["sz"]) / 100;
                    }
                }
            }

            // 确保字体大小有效
            if (isNaN(fontSize) || fontSize === undefined) {
                fontSize = 18; // 默认字体大小为 18pt，与第一个片段保持一致
            }

            if (!isNaN(fontSize)){
                let normAutofit = PPTXXmlUtils.getTextByPathList(textBodyNode, ["a:bodyPr", "a:normAutofit", "attrs", "fontScale"]);
                if (normAutofit !== undefined && normAutofit != 0){
                    //console.log("fontSize", fontSize, "normAutofit: ", normAutofit, normAutofit/100000)
                    fontSize = Math.round(fontSize * (normAutofit / 100000))
                }
            }

            return isNaN(fontSize) ? ((type == "br") ? "initial" : "inherit") : (fontSize * FONT_SIZE_FACTOR + "px");// + "pt");
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
                let underLine = node["a:rPr"]["attrs"]["u"] !== undefined ? node["a:rPr"]["attrs"]["u"] : "none";
                let strikethrough = node["a:rPr"]["attrs"]["strike"] !== undefined ? node["a:rPr"]["attrs"]["strike"] : 'noStrike';
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
            let getAlgn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (getAlgn === undefined) {
                getAlgn = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
            }
            if (getAlgn === undefined) {
                if (type == "title" || type == "ctrTitle" || type == "subTitle") {
                    let lvlIdx = 1;
                    let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
                    if (lvlNode !== undefined) {
                        lvlIdx = parseInt(lvlNode) + 1;
                    }
                    let lvlStr = "a:lvl" + lvlIdx + "pPr";
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

            let align = "inherit";
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
            let baseline = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
            return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
        }

        function getTableBorders(node, warpObj) {
            let borderStyle = "";
            if (node["a:bottom"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:bottom"]["a:ln"]
                    }
                }
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-bottom");
            }
            if (node["a:top"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:top"]["a:ln"]
                    }
                }
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-top");
            }
            if (node["a:right"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:right"]["a:ln"]
                    }
                }
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-right");
            }
            if (node["a:left"] !== undefined) {
                let obj = {
                    "p:spPr": {
                        "a:ln": node["a:left"]["a:ln"]
                    }
                }
                let borders = getBorder(obj, undefined, false, "shape", warpObj);
                borderStyle += borders.replace("border", "border-left");
            }

            return borderStyle;
        }
        //////////////////////////////////////////////////////////////////
        function getBorder(node, pNode, isSvgMode, bType, warpObj) {
            // DEBUG: Log for TextBox 5
            let shapeName = PPTXXmlUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
            if (shapeName === "TextBox 5") {
                console.log("=== TextBox 5 getBorder Debug ===");
                console.log("bType:", bType, "isSvgMode:", isSvgMode);
            }
            //console.log("getBorder", node, pNode, isSvgMode, bType)
            let cssText, lineNode, subNodeTxt;

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

            //let is_noFill = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:noFill"]);
            let is_noFill = PPTXXmlUtils.getTextByPathList(lineNode, ["a:noFill"]);
            if (is_noFill !== undefined) {
                return "hidden";
            }

            //console.log("lineNode: ", lineNode)
            let lnRefNode;
            let phClr = undefined;
            if (lineNode == undefined) {
                lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"])
                if (lnRefNode !== undefined){
                    let lnIdx = PPTXXmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
                    // DEBUG: Log for TextBox 5
                    if (shapeName === "TextBox 5") {
                        console.log("lnRefNode found, lnIdx:", lnIdx);
                    }
                    // Extract phClr from lnRef to replace placeholder colors in lnStyleLst
                    if (lnRefNode !== undefined) {
                        phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                        if (shapeName === "TextBox 5") {
                            console.log("phClr from lnRef:", phClr);
                        }
                    }
                    // 检查lnStyleLst的结构
                    const lnStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"];
                    // 处理lnStyleLst可能是对象而不是数组的情况
                    if (Array.isArray(lnStyleLst)) {
                        lineNode = lnStyleLst[Number(lnIdx)];
                        if (shapeName === "TextBox 5") {
                            console.log("lineNode from lnStyleLst:", JSON.stringify(lineNode).substring(0, 200));
                        }
                    } else {
                        // 如果是对象而不是数组，直接使用
                        lineNode = lnStyleLst;
                    }
                }
            }
            if (lineNode == undefined) {
                //is table
                cssText = "";
                lineNode = node
            }

            let borderColor;
            let borderWidth = 0;
            let borderType = "solid";
            let strokeDasharray = "0";
            if (lineNode !== undefined) {
                // Border width: 1pt = 12700, default = 0.75pt
                let w = PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "w"]);
                borderWidth = (w !== undefined) ? parseInt(w) / 12700 : (4/3);
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
                let fillTyp = getFillType(lineNode);
                //console.log("getBorder:node : fillTyp", fillTyp)
                if (fillTyp === "NO_FILL") {
                    borderColor = isSvgMode ? "none" : "";//"background-color: initial;";
                } else if (fillTyp === "SOLID_FILL") {
                    // 获取lnRef中的颜色作为phClr参数
                    if (!lnRefNode) {
                        lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
                    }
                    // phClr should already be extracted at line 780, but fallback here if not
                    if (phClr === undefined && lnRefNode !== undefined) {
                        phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                    }
                    borderColor = getSolidFill(lineNode["a:solidFill"], undefined, phClr, warpObj);
                } else if (fillTyp === "GRADIENT_FILL") {
                    borderColor = getGradientFill(lineNode["a:gradFill"], warpObj);
                    //console.log("shpFill",shpFill,grndColor.color)
                } else if (fillTyp === "PATTERN_FILL") {
                    borderColor = getPatternFill(lineNode["a:pattFill"], warpObj);
                }

            }

            //console.log("getBorder:node : borderColor", borderColor)
            // 2. drawingML namespace
            if (borderColor === undefined) {
                //let schemeClrNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
                // if (schemeClrNode !== undefined) {
                //     let schemeClr = "a:" + PPTXXmlUtils.getTextByPathList(schemeClrNode, ["attrs", "val"]);
                //     let borderColor = getSchemeColorFromTheme(schemeClr, undefined, undefined);
                // }
                let lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
                //console.log("getBorder: lnRef : ", lnRefNode)
                if (lnRefNode !== undefined) {
                    borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                }

                // if (borderColor !== undefined) {
                //     let shade = PPTXXmlUtils.getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
                //     if (shade !== undefined) {
                //         shade = parseInt(shade) / 10000;
                //         let color = tinycolor("#" + borderColor);
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
                // 检查borderColor是否已经是有效的颜色值
                if (borderColor && typeof borderColor === 'string') {
                    // 如果不是以#开头的十六进制颜色，添加#前缀
                    if (!borderColor.startsWith('#') && !borderColor.startsWith('rgb') && !borderColor.startsWith('hsl') && borderColor !== 'none' && borderColor !== 'hidden') {
                        borderColor = "#" + borderColor;
                    }
                }
            }
            cssText += " " + borderColor + " ";

            if (isSvgMode) {
                let result = { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
                // DEBUG: Log for TextBox 5
                if (shapeName === "TextBox 5") {
                    console.log("=== TextBox 5 getBorder Result ===");
                    console.log("borderColor:", borderColor);
                    console.log("borderWidth:", borderWidth);
                    console.log("borderType:", borderType);
                    console.log("strokeDasharray:", strokeDasharray);
                    console.log("Final result:", result);
                }
                return result;
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
        function getSlideBackgroundFill(warpObj, index) {
            let slideContent = warpObj["slideContent"];
            let slideLayoutContent = warpObj["slideLayoutContent"];
            let slideMasterContent = warpObj["slideMasterContent"];

            //console.log("slideContent: ", slideContent)
            //console.log("slideLayoutContent: ", slideLayoutContent)
            //console.log("slideMasterContent: ", slideMasterContent)
            //PPTXShapeUtils.getFillType(node)
            let bgPr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
            let bgRef = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
            //console.log("slideContent >> bgPr: ", bgPr, ", bgRef: ", bgRef)
            let bgcolor;
            if (bgPr !== undefined) {
                //bgcolor = "background-color: blue;";
                let bgFillTyp = getFillType(bgPr);

                if (bgFillTyp === "SOLID_FILL") {
                    let sldFill = bgPr["a:solidFill"];
                    let clrMapOvr;
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                        if (sldClrMapOvr !== undefined) {
                            clrMapOvr = sldClrMapOvr;
                        } else {
                            clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                        }

                    }
                    let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                    //var sldTint = getColorOpacity(sldFill);
                    //console.log("bgColor: ", bgColor)
                    //bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                    bgcolor = `background: #${sldBgClr};`;

                } else if (bgFillTyp === "GRADIENT_FILL") {
                    bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                } else if (bgFillTyp === "PIC_FILL") {
                    //console.log("PIC_FILL - ", bgFillTyp, bgPr, warpObj);
                    bgcolor = getBgPicFill(bgPr, "slideBg", warpObj, undefined, index);

                }
                //console.log(slideContent,slideMasterContent,color_ary,tint_ary,rot,bgcolor)
            } else if (bgRef !== undefined) {
                //console.log("slideContent",bgRef)
                let clrMapOvr;
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        clrMapOvr = sldClrMapOvr;
                    } else {
                        clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
                let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);

                // if (bgRef["a:srgbClr"] !== undefined) {
                //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                //     let schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);
                //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                // }
                let idx = Number(bgRef["attrs"]["idx"]);


                if (idx == 0 || idx == 1000) {
                    //no background
                } else if (idx > 0 && idx < 1000) {
                    //fillStyleLst in themeContent
                    //themeContent["a:fmtScheme"]["a:fillStyleLst"]
                    //bgcolor = "background: red;";
                } else if (idx > 1000) {
                    //bgFillStyleLst  in themeContent
                    //themeContent["a:fmtScheme"]["a:bgFillStyleLst"]
                    let trueIdx = idx - 1000;
                    // themeContent["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                    let sortblAry = [];
                    Object.keys(bgFillLst).forEach(key => {
                        let bgFillLstTyp = bgFillLst[key];
                        if (key != "attrs") {
                            if (bgFillLstTyp.constructor === Array) {
                                for (let i = 0; i < bgFillLstTyp.length; i++) {
                                    let obj = {};
                                    obj[key] = bgFillLstTyp[i];
                                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp[i]["attrs"]["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            } else {
                                let obj = {};
                                obj[key] = bgFillLstTyp;
                                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                obj["attrs"] = {
                                    "order": bgFillLstTyp["attrs"]["order"]
                                }
                                sortblAry.push(obj)
                            }
                        }
                    });
                    let sortByOrder = sortblAry.slice(0);
                    sortByOrder.sort((a, b) => {
                        return a.idex - b.idex;
                    });
                    let bgFillLstIdx = sortByOrder[trueIdx - 1];
                    let bgFillTyp = getFillType(bgFillLstIdx);
                    if (bgFillTyp === "SOLID_FILL") {
                        let sldFill = bgFillLstIdx["a:solidFill"];
                        let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                        bgcolor = `background: #${sldBgClr};`;
                        //console.log("slideMasterContent - sldFill",sldFill)
                    } else if (bgFillTyp === "GRADIENT_FILL") {
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
                let clrMapOvr;
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    clrMapOvr = sldClrMapOvr;
                } else {
                    clrMapOvr = PPTXXmlUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
                }
                if (bgPr !== undefined) {
                    let bgFillTyp = getFillType(bgPr);
                    if (bgFillTyp === "SOLID_FILL") {
                        let sldFill = bgPr["a:solidFill"];

                        let sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
                        //var sldTint = getColorOpacity(sldFill);
                        // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                        bgcolor = `background: #${sldBgClr};`;
                    } else if (bgFillTyp === "GRADIENT_FILL") {
                        bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                    } else if (bgFillTyp === "PIC_FILL") {
                        bgcolor = getBgPicFill(bgPr, "slideLayoutBg", warpObj, undefined, index);

                    }
                    //console.log("slideLayoutContent",bgcolor)
                } else if (bgRef !== undefined) {
                    console.log("slideLayoutContent: bgRef", bgRef)
                    //bgcolor = "background: white;";
                    let phClr = getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
                    let idx = Number(bgRef["attrs"]["idx"]);
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
                        let trueIdx = idx - 1000;
                        let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                        let sortblAry = [];
                        Object.keys(bgFillLst).forEach(key => {
                            //console.log("cubicBezTo[" + key + "]:");
                            let bgFillLstTyp = bgFillLst[key];
                            if (key != "attrs") {
                                if (bgFillLstTyp.constructor === Array) {
                                    for (let i = 0; i < bgFillLstTyp.length; i++) {
                                        let obj = {};
                                        obj[key] = bgFillLstTyp[i];
                                        obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp[i]["attrs"]["order"]
                                        }
                                        sortblAry.push(obj)
                                    }
                                } else {
                                    let obj = {};
                                    obj[key] = bgFillLstTyp;
                                    obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                    obj["attrs"] = {
                                        "order": bgFillLstTyp["attrs"]["order"]
                                    }
                                    sortblAry.push(obj)
                                }
                            }
                        });
                        let sortByOrder = sortblAry.slice(0);
                        sortByOrder.sort((a, b) => {
                            return a.idex - b.idex;
                        });
                        let bgFillLstIdx = sortByOrder[trueIdx - 1];
                        let bgFillTyp = getFillType(bgFillLstIdx);
                        if (bgFillTyp === "SOLID_FILL") {
                            let sldFill = bgFillLstIdx["a:solidFill"];
                            //console.log("sldFill: ", sldFill)
                            //var sldTint = getColorOpacity(sldFill);
                            //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                            let sldBgClr = getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
                            //console.log("bgcolor: ", bgcolor)
                            bgcolor = `background: #${sldBgClr};`;
                        } else if (bgFillTyp === "GRADIENT_FILL") {
                            //console.log("GRADIENT_FILL: ", bgFillLstIdx, phClr)
                            bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
                        } else if (bgFillTyp === "PIC_FILL") {
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
                        let bgFillTyp = getFillType(bgPr);
                        if (bgFillTyp === "SOLID_FILL") {
                            let sldFill = bgPr["a:solidFill"];
                            let sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj);
                            // var sldTint = getColorOpacity(sldFill);
                            // bgcolor = "background: rgba(" + hexToRgbNew(bgColor) + "," + sldTint + ");";
                            bgcolor = `background: #${sldBgClr};`;
                        } else if (bgFillTyp === "GRADIENT_FILL") {
                            bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
                        } else if (bgFillTyp === "PIC_FILL") {
                            bgcolor = getBgPicFill(bgPr, "slideMasterBg", warpObj, undefined, index);
                        }
                    } else if (bgRef !== undefined) {
                        //let obj={
                        //    "a:solidFill": bgRef
                        //}
                        let phClr = getSolidFill(bgRef, clrMap, undefined, warpObj);
                        // let phClr;
                        // if (bgRef["a:srgbClr"] !== undefined) {
                        //     phClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:srgbClr", "attrs", "val"]); //#...
                        // } else if (bgRef["a:schemeClr"] !== undefined) { //a:schemeClr
                        //     let schemeClr = PPTXXmlUtils.getTextByPathList(bgRef, ["a:schemeClr", "attrs", "val"]);

                        //     phClr = getSchemeColorFromTheme("a:" + schemeClr, slideMasterContent, undefined); //#...
                        // }
                        let idx = Number(bgRef["attrs"]["idx"]);
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
                            let trueIdx = idx - 1000;
                            let bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
                            let sortblAry = [];
                            Object.keys(bgFillLst).forEach(key => {
                                //console.log("cubicBezTo[" + key + "]:");
                                let bgFillLstTyp = bgFillLst[key];
                                if (key != "attrs") {
                                    if (bgFillLstTyp.constructor === Array) {
                                        for (let i = 0; i < bgFillLstTyp.length; i++) {
                                            let obj = {};
                                            obj[key] = bgFillLstTyp[i];
                                            obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                                            obj["attrs"] = {
                                                "order": bgFillLstTyp[i]["attrs"]["order"]
                                            }
                                            sortblAry.push(obj)
                                        }
                                    } else {
                                        let obj = {};
                                        obj[key] = bgFillLstTyp;
                                        obj["idex"] = bgFillLstTyp["attrs"]["order"];
                                        obj["attrs"] = {
                                            "order": bgFillLstTyp["attrs"]["order"]
                                        }
                                        sortblAry.push(obj)
                                    }
                                }
                            });
                            let sortByOrder = sortblAry.slice(0);
                            sortByOrder.sort((a, b) => {
                                return a.idex - b.idex;
                            });
                            let bgFillLstIdx = sortByOrder[trueIdx - 1];
                            let bgFillTyp = getFillType(bgFillLstIdx);
                            //console.log("bgFillLstIdx: ", bgFillLstIdx, ", bgFillTyp: ", bgFillTyp, ", phClr: ", phClr);
                            if (bgFillTyp == "SOLID_FILL") {
                                let sldFill = bgFillLstIdx["a:solidFill"];
                                //console.log("sldFill: ", sldFill)
                                //var sldTint = getColorOpacity(sldFill);
                                //bgcolor = "background: rgba(" + hexToRgbNew(phClr) + "," + sldTint + ");";
                                let sldBgClr = getSolidFill(sldFill, clrMap, phClr, warpObj);
                                //console.log("bgcolor: ", bgcolor)
                                bgcolor = `background: #${sldBgClr};`;
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
            let bgcolor = "";
            if (bgPr !== undefined) {
                let grdFill = bgPr["a:gradFill"];
                let gsLst = grdFill["a:gsLst"]["a:gs"];
                //var startColorNode, endColorNode;
                let color_ary = [];
                var pos_ary = [];
                //let tint_ary = [];
                for (let i = 0; i < gsLst.length; i++) {
                    let lo_tint;
                    let lo_color = getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
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
                let lin = grdFill["a:lin"];
                let rot = 90;
                if (lin !== undefined) {
                    rot = PPTXXmlUtils.angleToDegrees(lin["attrs"]["ang"]);// + 270;
                    //console.log("rot: ", rot)
                    rot = rot + 90;
                }
                bgcolor = `background: linear-gradient(${rot}deg,`;
                for (let i = 0; i < gsLst.length; i++) {
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
            let bgcolor;
            let picFillResult = getPicFill(sorce, bgPr["a:blipFill"], warpObj);
            let picFillBase64 = picFillResult;
            if (typeof picFillResult === 'object' && picFillResult.img) {
                picFillBase64 = picFillResult.img;
            }
            let ordr = bgPr["attrs"]["order"];
            let aBlipNode = bgPr["a:blipFill"]["a:blip"];
            //a:duotone
            let duotone = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
            if (duotone !== undefined) {
                //console.log("pic duotone: ", duotone)
                let clr_ary = [];
                // duotone.forEach(clr => {
                //     console.log("pic duotone clr: ", clr)
                // }) 
                Object.keys(duotone).forEach(clr_type => {
                    //console.log("pic duotone clr: clr_type: ", clr_type, duotone[clr_type])
                    if (clr_type != "attrs") {
                        let obj = {};
                        obj[clr_type] = duotone[clr_type];
                        clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    }
                    // Object.keys(duotone[clr_type]).forEach(clr => {
                    //     if (clr != "order") {
                    //         let obj = {};
                    //         obj[clr_type] = duotone[clr_type][clr];
                    //         clr_ary.push(getSolidFill(obj, undefined, phClr, warpObj));
                    //     }
                    // })
                })
                //console.log("pic duotone clr_ary: ", clr_ary);
                //filter: url(file.svg#filter-element-id)
                //https://codepen.io/bhenbe/pen/QEZOvd
                //https://www.w3schools.com/cssref/css3_pr_filter.asp

                // let color1 = clr_ary[0];
                // let color2 = clr_ary[1];
                // let cssName = "";

                // let styleText_before_after = "content: '';" +
                //     "display: block;" +
                //     "width: 100%;" +
                //     "height: 100%;" +
                //     // "z-index: 1;" +
                //     "position: absolute;" +
                //     "top: 0;" +
                //     "left: 0;";

                // let cssName = "slide-background-" + index + "::before," + " .slide-background-" + index + "::after";
                // styleTable[styleText_before_after] = {
                //     "name": cssName,
                //     "text": styleText_before_after
                // };


                // let styleText_after = "background-color: #" + clr_ary[1] + ";" +
                //     "mix-blend-mode: darken;";

                // cssName = "slide-background-" + index + "::after";
                // styleTable[styleText_after] = {
                //     "name": cssName,
                //     "text": styleText_after
                // };

                // let styleText_before = "background-color: #" + clr_ary[0] + ";" +
                //     "mix-blend-mode: lighten;";

                // cssName = "slide-background-" + index + "::before";
                // styleTable[styleText_before] = {
                //     "name": cssName,
                //     "text": styleText_before
                // };

            }
            //a:alphaModFix
            let aphaModFixNode = PPTXXmlUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"])
            let imgOpacity = "";
            if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
                var amt = parseInt(aphaModFixNode["amt"]) / 100000;
                //let opacity = amt;
                imgOpacity = "opacity:" + amt + ";";

            }
            // 使用getPicFill函数返回的填充模式信息
            let prop_style = "";
            if (typeof picFillResult === 'object') {
                if (picFillResult.backgroundSize) {
                    prop_style += "background-size: " + picFillResult.backgroundSize + ";";
                }
                if (picFillResult.backgroundPosition) {
                    prop_style += "background-position: " + picFillResult.backgroundPosition + ";";
                }
                if (picFillResult.backgroundRepeat) {
                    prop_style += "background-repeat: " + picFillResult.backgroundRepeat + ";";
                }
            }
            bgcolor = "background: url(" + picFillBase64 + ");  z-index: " + ordr + ";" + prop_style + imgOpacity;

            return bgcolor;
        }
      
        
        function getGradientFill(node, warpObj) {
            //console.log("getGradientFill: node", node)
            let gsLst = node["a:gsLst"]["a:gs"];
            //get start color
            let color_ary = [];
            let tint_ary = [];
            for (let i = 0; i < gsLst.length; i++) {
                let lo_tint;
                let lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
                //console.log("lo_color",lo_color)
                color_ary[i] = lo_color;
            }
            //get rot
            let lin = node["a:lin"];
            let rot = 0;
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
            let img;
            let rId = node["a:blip"]["attrs"]["r:embed"];
            let imgPath;
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

                let imgExt = imgPath.split(".").pop();
                if (imgExt == "xml") {
                    return undefined;
                }
                let imgFile = warpObj["zip"].file(imgPath);
                if (imgFile === null || imgFile === undefined) {
                    console.warn("Image file not found:", imgPath);
                    return undefined;
                }
                let imgArrayBuffer = imgFile.asArrayBuffer();
                let imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                img = "data:" + imgMimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
                //warpObj["loaded-images"][imgPath] = img; //"defaultTextStyle": defaultTextStyle,
                setTextByPathList(warpObj, ["loaded-images", imgPath], img); //, type, rId
            }
            // 处理图像属性 - Tile, Stretch, or Display Portion of Image
            let tileNode = node["a:tile"];
            let stretchNode = node["a:stretch"];
            let fillMode = "stretch";
            let backgroundSize = "cover";
            let backgroundPosition = "center";
            let backgroundRepeat = "no-repeat";
            
            if (tileNode) {
                // 平铺模式
                fillMode = "tile";
                backgroundRepeat = "repeat";
                
                // 处理平铺大小
                let sx = tileNode["attrs"]["sx"];
                let sy = tileNode["attrs"]["sy"];
                if (sx && sy) {
                    let widthPercent = parseInt(sx) / 100000 * 100;
                    let heightPercent = parseInt(sy) / 100000 * 100;
                    backgroundSize = widthPercent + "% " + heightPercent + "%";
                }
                
                // 处理平铺偏移
                let tx = tileNode["attrs"]["tx"];
                let ty = tileNode["attrs"]["ty"];
                if (tx && ty) {
                    let xPercent = parseInt(tx) / 100000 * 100;
                    let yPercent = parseInt(ty) / 100000 * 100;
                    backgroundPosition = xPercent + "% " + yPercent + "%";
                }
            } else if (stretchNode) {
                // 拉伸模式
                fillMode = "stretch";
                let fillRect = stretchNode["a:fillRect"];
                if (fillRect) {
                    // 处理填充矩形
                    backgroundSize = "cover";
                }
            }
            
            // 返回包含图像和填充模式的对象
            return {
                "img": img,
                "fillMode": fillMode,
                "backgroundSize": backgroundSize,
                "backgroundPosition": backgroundPosition,
                "backgroundRepeat": backgroundRepeat
            };
        }
        function getPatternFill(node, warpObj) {
            //https://developer.mozilla.org/en-US/docs/Web/CSS/CSS_Images/Using_CSS_gradients
            //https://cssgradient.io/blog/css-gradient-text/
            //https://css-tricks.com/background-patterns-simplified-by-conic-gradients/
            //https://stackoverflow.com/questions/6705250/how-to-get-a-pattern-into-a-written-text-via-css
            //https://stackoverflow.com/questions/14072142/striped-text-in-css
            //https://css-tricks.com/stripes-css/
            //https://yuanchuan.dev/gradient-shapes/
            let fgColor = "", bgColor = "", prst = "";
            let bgClr = node["a:bgClr"];
            let fgClr = node["a:fgClr"];
            prst = node["attrs"]["prst"];
            fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
            bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);
            //var angl_ary = getAnglefromParst(prst);
            //let ptrClr = "repeating-linear-gradient(" + angl + "deg,  #" + bgColor + ",#" + fgColor + " 2px);"
            //linear-gradient(0deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black), 
            //linear-gradient(90deg, black 10 %, transparent 10 %, transparent 90 %, black 90 %, black);
            let linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
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
            let color = "";
            let clrNode;
            if (node["a:srgbClr"] !== undefined) {
                clrNode = node["a:srgbClr"];
                color = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //#...
            } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
                clrNode = node["a:schemeClr"];
                let schemeClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]);
                color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
                //console.log("schemeClr: ", schemeClr, "color: ", color)
            } else if (node["a:scrgbClr"] !== undefined) {
                clrNode = node["a:scrgbClr"];
                //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
                let defBultColorVals = clrNode["attrs"];
                let red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
                let green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
                let blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
                //let scrgbClr = red + "," + green + "," + blue;
                color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
                //console.log("scrgbClr: " + scrgbClr);

            } else if (node["a:prstClr"] !== undefined) {
                clrNode = node["a:prstClr"];
                //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
                let prstClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]["val"];
                color = getColorName2Hex(prstClr);
                //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
            } else if (node["a:hslClr"] !== undefined) {
                clrNode = node["a:hslClr"];
                //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
                let defBultColorVals = clrNode["attrs"];
                let hue = Number(defBultColorVals["hue"]) / 100000;
                let sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
                let lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
                //let hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
                let hsl2rgb = hslToRgb(hue, sat, lum);
                color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
                //defBultColor = cnvrtHslColor2Hex(hslClr); //TODO
                // console.log("hslClr: " + hslClr);
            } else if (node["a:sysClr"] !== undefined) {
                clrNode = node["a:sysClr"];
                //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
                let sysClr = PPTXXmlUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
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
            let isAlpha = false;
            let alpha = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
            //console.log("alpha: ", alpha)
            if (!isNaN(alpha)) {
                // var al_color = new colz.Color(color);
                // al_color.setAlpha(alpha);
                // let ne_color = al_color.rgba.toString();
                // color = (rgba2hex(ne_color))
                let al_color = tinycolor(color);
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

            let hueMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
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
            //let hueOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
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
            let lumMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
            //console.log("lumMod: ", lumMod)
            if (!isNaN(lumMod)) {
                color = applyLumMod(color, lumMod, isAlpha);
            }
            //let lumMod_color = applyLumMod(color, 0.5);
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
            let lumOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
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
            let satMod = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
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
            // let satOff = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
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
            let shade = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
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
            let tint = parseInt (PPTXXmlUtils.getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
            if (!isNaN(tint)) {
                color = applyTint(color, tint, isAlpha);
            }
            //console.log("color [%cfinal]: ", "color: #" + color, tinycolor(color).toHslString(), color)

            return color;
        }
        function toHex(n) {
            let hex = n.toString(16);
            while (hex.length < 2) { hex = "0" + hex; }
            return hex;
        }
        function hslToRgb(hue, sat, light) {
            let t1, t2, r, g, b;
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
            let hex;
            let colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
            let colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
            let findIndx = colorName.indexOf(name);
            if (findIndx != -1) {
                hex = colorHex[findIndx];
            }
            return hex;
        }
        function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
            //<p:clrMap ...> in slide master
            // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
            //console.log("getSchemeColorFromTheme: schemeClr: ", schemeClr, ",clrMap: ", clrMap)
            let color = '';
            var slideLayoutClrOvride;
            if (clrMap !== undefined) {
                slideLayoutClrOvride = clrMap;//getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
            } else if (warpObj !== undefined) {
                let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                if (sldClrMapOvr !== undefined) {
                    slideLayoutClrOvride = sldClrMapOvr;
                } else {
                    let sldClrMapOvr = PPTXXmlUtils.getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
                    if (sldClrMapOvr !== undefined) {
                        slideLayoutClrOvride = sldClrMapOvr;
                    } else {
                        slideLayoutClrOvride = PPTXXmlUtils.getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
                    }

                }
            }
            //console.log("getSchemeColorFromTheme slideLayoutClrOvride: ", slideLayoutClrOvride);
            let schmClrName = schemeClr.substr(2);
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
                let refNode = PPTXXmlUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
                color = PPTXXmlUtils.getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
                //console.log("themeContent: color", color);
                if (color === undefined && refNode !== undefined) {
                    color = PPTXXmlUtils.getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
                }
            }
            //console.log(color)
            return color;
        }

        function extractChartData(serNode, warpObj) {

            let dataMat = new Array();

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
                    let colName = PPTXXmlUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

                    // Category (string or number)
                    let rowNames = {};
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

                    // Extract series style information
                    let seriesStyle = {};
                    
                    // Extract fill color if available
                    let fillType = getFillType(PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr"]));
                    if (fillType === "SOLID_FILL" && warpObj !== undefined) {
                        let fillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:solidFill"]);
                        if (fillNode !== undefined) {
                            let fillColor = getSolidFill(fillNode, undefined, undefined, warpObj);
                            if (fillColor !== undefined) {
                                if (fillColor && !fillColor.startsWith('#')) {
                                    fillColor = '#' + fillColor;
                                }
                                seriesStyle.fillColor = fillColor;
                            }
                        }
                    } else if (fillType === "GRADIENT_FILL" && warpObj !== undefined) {
                        let gradFillNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:gradFill"]);
                        if (gradFillNode !== undefined) {
                            let gradientFill = getGradientFill(gradFillNode, warpObj);
                            if (gradientFill !== undefined) {
                                seriesStyle.gradientFill = gradientFill;
                            }
                        }
                    }
                    
                    // Extract line color if available
                    let lineNode = PPTXXmlUtils.getTextByPathList(innerNode, ["c:spPr", "a:ln"]);
                    if (lineNode !== undefined && warpObj !== undefined) {
                        let lineFillType = getFillType(lineNode);
                        if (lineFillType === "SOLID_FILL") {
                            let lineColor = getSolidFill(lineNode["a:solidFill"], undefined, undefined, warpObj);
                            if (lineColor !== undefined) {
                                if (lineColor && !lineColor.startsWith('#')) {
                                    lineColor = '#' + lineColor;
                                }
                                seriesStyle.lineColor = lineColor;
                            }
                        } else if (lineFillType === "GRADIENT_FILL") {
                            let lineGradFillNode = lineNode["a:gradFill"];
                            if (lineGradFillNode !== undefined) {
                                let lineGradientFill = getGradientFill(lineGradFillNode, warpObj);
                                if (lineGradientFill !== undefined) {
                                    seriesStyle.lineGradientFill = lineGradientFill;
                                }
                            }
                        }
                    }

                    dataMat.push({ key: colName, values: dataRow, xlabels: rowNames, style: seriesStyle });
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
                let obj = this;
                let lent = parts.length;
                for (let i = 0; i < lent; i++) {
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
            let result = "";
            if (node.constructor === Array) {
                let l = node.length;
                for (let i = 0; i < l; i++) {
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
            let color = tinycolor(rgbStr).toHsl();
            //console.log("applyShade  color: ", color, ", shadeValue: ", shadeValue)
            // 确保shadeValue在0-1之间
            shadeValue = Math.max(0, Math.min(1, shadeValue));
            // PPTX标准：Shade = L * shadeValue
            let cacl_l = Math.max(0, Math.min(1, color.l * shadeValue));
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
            let color = tinycolor(rgbStr).toHsl();
            //console.log("applyTint  color: ", color, ", tintValue: ", tintValue)
            // 确保tintValue在0-1之间
            tintValue = Math.max(0, Math.min(1, tintValue));
            // PPTX标准：Tint = L * tintValue + (1 - tintValue)
            let cacl_l = Math.max(0, Math.min(1, color.l * tintValue + (1 - tintValue)));
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
            let color = tinycolor(rgbStr).toHsl();
            //console.log("applyLumOff  color.l: ", color.l, ", offset: ", offset, ", color.l + offset : ", color.l + offset)
            let lum = offset + color.l;
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
            let color = tinycolor(rgbStr).toHsl();
            //console.log("applyLumMod  color.l: ", color.l, ", multiplier: ", multiplier, ", color.l * multiplier : ", color.l * multiplier)
            let cacl_l = color.l * multiplier;
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
            let color = tinycolor(rgbStr).toHsl();
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
        //     let color = tinycolor(rgbStr).toHsl();
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
            let color = tinycolor(rgbStr).toHsl();
            //console.log("applySatMod  color.s: ", color.s, ", multiplier: ", multiplier, ", color.s * multiplier : ", color.s * multiplier)
            let cacl_s = color.s * multiplier;
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

            let svgAngle = '',
                svgHeight = h,
                svgWidth = w,
                svg = '',
                xy_ary = SVGangle(angl, svgHeight, svgWidth),
                x1 = xy_ary[0],
                y1 = xy_ary[1],
                x2 = xy_ary[2],
                y2 = xy_ary[3];

            let sal = stopsArray.length,
                sr = sal < 20 ? 100 : 1000;
            svgAngle = ' gradientUnits="userSpaceOnUse" x1="' + x1 + '%" y1="' + y1 + '%" x2="' + x2 + '%" y2="' + y2 + '%"';
            svgAngle = '<linearGradient id="linGrd_' + shpId + '"' + svgAngle + '>\n';
            svg += svgAngle;

            for (let i = 0; i < sal; i++) {
                var tinClr = tinycolor("#" + color_arry[i]);
                let alpha = tinClr.getAlpha();
                //console.log("color: ", color_arry[i], ", rgba: ", tinClr.toHexString(), ", alpha: ", alpha)
                svg += '<stop offset="' + Math.round(parseFloat(stopsArray[i]) / 100 * sr) / sr + '" style="stop-color:' + tinClr.toHexString() + '; stop-opacity:' + (alpha) + ';"';
                svg += '/>\n'
            }

            svg += '</linearGradient>\n' + '';

            return svg
        }
        function getMiddleStops(s) {
            let sArry = ['0%', '100%'];
            if (s == 0) {
                return sArry;
            } else {
                let i = s;
                while (i--) {
                    let middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                        middleStopString = middleStop + "%";
                    sArry.splice(-1, 0, middleStopString);
                } // AM: add into stopsArray before 100%
            }
            return sArry
        }
        function SVGangle(deg, svgHeight, svgWidth) {
            let w = parseFloat(svgWidth),
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
            let m = o + (n / i);
                tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1,
                ty1 = ty1 == 2 ? i * tx1 + l : ty1,
                tx2 = tx2 == 2 ? w - tx1 : tx2,
                ty2 = ty2 == 2 ? h - ty1 : ty2;
            let x1 = Math.round(tx2 / w * 100 * 100) / 100,
                y1 = Math.round(ty2 / h * 100 * 100) / 100,
                x2 = Math.round(tx1 / w * 100 * 100) / 100,
                y2 = Math.round(ty1 / h * 100 * 100) / 100;
            return [x1, y1, x2, y2];
        }
        function getSvgImagePattern(node, fill, shpId, warpObj) {
            // 处理 fill 参数是对象的情况
            let fillUrl = fill;
            if (typeof fill === 'object' && fill.img) {
                fillUrl = fill.img;
            }
            
            let pic_dim = getBase64ImageDimensions(fillUrl);
            let width = pic_dim[0];
            let height = pic_dim[1];
            //console.log("getSvgImagePattern node:", node);
            let blipFillNode = node["p:spPr"]["a:blipFill"];
            let sx = 0, sy = 0;
            let tileNode = PPTXXmlUtils.getTextByPathList(blipFillNode, ["a:tile", "attrs"])
            if (tileNode !== undefined && tileNode["sx"] !== undefined) {
                sx = (parseInt(tileNode["sx"]) / 100000) * width;
                sy = (parseInt(tileNode["sy"]) / 100000) * height;
            }

            let blipNode = node["p:spPr"]["a:blipFill"]["a:blip"];
            let tialphaModFixNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:alphaModFix", "attrs"])
            let imgOpacity = "";
            if (tialphaModFixNode !== undefined && tialphaModFixNode["amt"] !== undefined && tialphaModFixNode["amt"] != "") {
                var amt = parseInt(tialphaModFixNode["amt"]) / 100000;
                let opacity = amt;
                let imgOpacity = "opacity='" + opacity + "'";

            }
            let ptrn = '';
            if (sx !== undefined && sx != 0) {
                ptrn = '<pattern id="imgPtrn_' + shpId + '" x="0" y="0"  width="' + sx + '" height="' + sy + '" patternUnits="userSpaceOnUse">';
            } else {
                ptrn = '<pattern id="imgPtrn_' + shpId + '"  patternContentUnits="objectBoundingBox"  width="1" height="1">';
            }
            let duotoneNode = PPTXXmlUtils.getTextByPathList(blipNode, ["a:duotone"])
            let fillterNode = "";
            let filterUrl = "";
            if (duotoneNode !== undefined) {
                //console.log("pic duotoneNode: ", duotoneNode)
                var clr_ary = [];
                Object.keys(duotoneNode).forEach(clr_type => {
                    //Object.keys(duotoneNode[clr_type]).forEach(clr => {
                    //console.log("blip pic duotone clr: ", duotoneNode[clr_type][clr], clr)
                    if (clr_type != "attrs") {
                        let obj = {};
                        obj[clr_type] = duotoneNode[clr_type];
                        //console.log("blip pic duotone obj: ", obj)
                        let hexClr = getSolidFill(obj, undefined, undefined, warpObj)
                        //clr_ary.push();

                        let color = tinycolor("#" + hexClr);
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

            fillUrl = PPTXXmlUtils.escapeHtml(fillUrl);
            if (sx !== undefined && sx != 0) {
                ptrn += '<image  xlink:href="' + fillUrl + '" x="0" y="0" width="' + sx + '" height="' + sy + '" ' + imgOpacity + ' ' + filterUrl + '></image>';
            } else {
                ptrn += '<image  xlink:href="' + fillUrl + '" preserveAspectRatio="none" width="1" height="1" ' + imgOpacity + ' ' + filterUrl + '></image>';
            }
            ptrn += '</pattern>';

            //console.log("getSvgImagePattern(...) pic_dim:", pic_dim, ", fillColor: ", fill, ", blipNode: ", blipNode, ",sx: ", sx, ", sy: ", sy, ", clr_ary: ", clr_ary, ", ptrn: ", ptrn)

            return ptrn;
        }

        function getBase64ImageDimensions(imgSrc) {
            let image = new Image();
            let w, h;
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
            let defRtl = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "a:defPPr", "attrs", "rtl"]);
            if (defRtl !== undefined) {
                if (defRtl == "1"){
                    return "content-rtl";
                } else if (defRtl == "0") {
                    return "content";
                }
            }
            //let lvl1Rtl = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:lstStyle", "lvl1pPr", "attrs", "rtl"]);
            // if (lvl1Rtl !== undefined) {
            //     if (lvl1Rtl == "1") {
            //         return "content-rtl";
            //     } else if (lvl1Rtl == "0") {
            //         return "content";
            //     }
            // }
            let rtlCol = PPTXXmlUtils.getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "rtlCol"]);
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
            let slideMasterTextStyles = warpObj["slideMasterTextStyles"];
            let dirLoc = "";

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
                let dirVal = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles[dirLoc], ["a:lvl1pPr", "attrs", "rtl"]);
                if (dirVal == "1") {
                    return "content-rtl";
                }
            } 
            // else {
            //     if (type == "textBox") {
            //         let dirVal = PPTXXmlUtils.getTextByPathList(warpObj, ["defaultTextStyle", "a:lvl1pPr", "attrs", "rtl"]);
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
            //let lstStyle = textBodyNode["a:lstStyle"];
            let lvl = 1
            var spcBefNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcBef", "a:spcPts", "attrs", "val"]);
            var spcAftNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:spcAft", "a:spcPts", "attrs", "val"]);
            let lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPct", "attrs", "val"]);
            let lnSpcNodeType = "Pct";
            if (lnSpcNode === undefined) {
                lnSpcNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "a:lnSpc", "a:spcPts", "attrs", "val"]);
                if (lnSpcNode !== undefined) {
                    lnSpcNodeType = "Pts";
                }
            }
            let lvlNode = PPTXXmlUtils.getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            let fontSize;
            if  (PPTXXmlUtils.getTextByPathList(pNode, ["a:r"]) !== undefined) {
                let fontSizeStr = getFontSize(pNode["a:r"], textBodyNode,undefined, lvl, type, warpObj);
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
            let isInLayoutOrMaster = true;
            if(type == "shape" || type == "textBox"){
                isInLayoutOrMaster = false;
            }
            if (isInLayoutOrMaster && (spcBefNode === undefined || spcAftNode === undefined || lnSpcNode === undefined)) {
                //check in layout
                if (idx !== undefined) {
                    let laypPrNode = PPTXXmlUtils.getTextByPathList(warpObj, ["slideLayoutTables", "idxTable", idx, "p:txBody", "a:p", (lvl - 1), "a:pPr"]);

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
                let dirLoc = "";
                lvl = "a:lvl" + lvl + "pPr";
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
                let inLvlNode = PPTXXmlUtils.getTextByPathList(slideMasterTextStyles, [dirLoc, lvl]);
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
            let spcBefor = 0, spcAfter = 0, spcLines = 0;
            let marginTopBottomStr = "";
            if (spcBefNode !== undefined) {
                spcBefor = parseInt(spcBefNode) / 100;
            }
            if (spcAftNode !== undefined) {
                spcAfter = parseInt(spcAftNode) / 100;
            }
            
            // 移除line-height设置，因为PPTX的行间距与CSS的line-height含义不同
            // 直接使用会导致行高过大
            // 只保留段落间距设置

            // 段落前间距
            if (spcBefNode !== undefined) {
                // 转换为像素（假设1pt = 1.33px）
                let marginTop = spcBefor * 1.33;
                // 减小段落前间距，避免行距过大
                marginTop = Math.max(0, marginTop * 0.8);
                marginTopBottomStr += "margin-top: " + marginTop + "px;";
            }
            
            // 段落后间距
            if (spcAftNode !== undefined) {
                // 转换为像素（假设1pt = 1.33px）
                let marginBottom = spcAfter * 1.33;
                // 减小段落后间距，避免行距过大
                marginBottom = Math.max(0, marginBottom * 0.8);
                marginTopBottomStr += "margin-bottom: " + marginBottom + "px;";
            }

            //console.log("getVerticalMargins 2 fontSize:", fontSize, "lnSpcNode:", lnSpcNode, "spcLines:", spcLines, "spcBefor:", spcBefor, "spcAfter:", spcAfter)
            //console.log("getVerticalMargins 3 ", marginTopBottomStr, pNode, warpObj)

            //return spcAft + spcBef;
            return marginTopBottomStr;
        }
        function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj) {
            let algn = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
            if (algn === undefined) {
                let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;
                let lvlIdx = 1;
                let lvlNode = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
                if (lvlNode !== undefined) {
                    lvlIdx = parseInt(lvlNode) + 1;
                }
                let lvlStr = "a:lvl" + lvlIdx + "pPr";

                let lstStyle = textBodyNode["a:lstStyle"];
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
                // 尝试从布局和母版节点中直接获取对齐属性
                if (algn === undefined && pPrNodeLaout) {
                    algn = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                }
                if (algn === undefined && pPrNodeMaster) {
                    algn = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }

            if (algn === undefined) {
                // 对于特定位置的文本元素，尝试根据位置推断对齐方式
                // 例如，位于幻灯片右侧的文本元素可能是右对齐的
                if (type == "title" || type == "subTitle" || type == "ctrTitle") {
                    return "h-mid";
                } else if (type == "sldNum") {
                    return "h-right";
                } else {
                    // 默认返回左对齐
                    return "h-left";
                }
            }
            if (algn !== undefined) {
                switch (algn) {
                    case "l":
                        if (prg_dir == "pregraph-rtl"){
                            return "h-left-rtl";
                        }else{
                            return "h-left";
                        }
                        break;
                    case "r":
                        if (prg_dir == "pregraph-rtl") {
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
            let pPrNodeLaout, pPrNodeMaster;
            var pPrNode = node["a:pPr"];
            //lvl
            let lvl = 1;
            let lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
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
                let lvlStr = "a:lvl" + lvl + "pPr";
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
            let rtl = PPTXXmlUtils.getTextByPathList(node, ["a:pPr", "attrs", "rtl"]);
            //console.log("getPregraphDir node:", node, "textBodyNode", textBodyNode, "rtl:", rtl, "idx", idx, "type", type, "warpObj", warpObj)
          

            if (rtl === undefined) {
                let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
                let pPrNodeLaout = layoutMasterNode.nodeLaout;
                let pPrNodeMaster = layoutMasterNode.nodeMaster;
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
    function getPregraphMargn(pNode, idx, type, isBullate, warpObj, fontSize){
            if (!isBullate){
                return ["",0];
            }
            let marLStr = "", marRStr = "" , maginVal = 0;
            let pPrNode = pNode["a:pPr"];
            let layoutMasterNode = getLayoutAndMasterNode(pNode, idx, type, warpObj);
            let pPrNodeLaout = layoutMasterNode.nodeLaout;
            let pPrNodeMaster = layoutMasterNode.nodeMaster;

            // 在 RTL 模式下，margin 和 indent 的语义保持不变
            // - marL (margin-left): 左边距（在 RTL 中是文本结束边的距离）
            // - indent: 缩进，负值表示从起始边向内缩进
            // 这些属性值直接转换为 CSS 的 padding，不需要根据 RTL 进行方向翻转
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
            //
            //marL
            let marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }
            let marginLeft = 0;
            if (marLNode !== undefined) {
                marginLeft = parseInt(marLNode) * SLIDE_FACTOR;
            }
            if ((indentNode !== undefined || marLNode !== undefined)) {
                // let lvlIndent = defTabSz * lvl;
                // 无论 RTL 还是 LTR，marL 和 indent 都转换为 padding-left
                // 在 RTL 模式下，文本从右向左排列，但 padding-left 仍然是左边距
                // align="ctr" + marL > 0: 文本居中，但有左边距，导致整体偏右
                // align="ctr" + indent < 0: 文本居中，向左（文本起始方向）缩进，导致整体偏右
                marLStr = "padding-left: ";
                if (isBullate) {
                    maginVal = Math.abs(0 - indent);
                    // 减去项目符号数字的长度/大小，根据字体大小估算
                    let bulletSizeAdjustment = 0;
                    if (fontSize !== undefined) {
                        // 对于数字项目符号，根据字体大小估算宽度
                        bulletSizeAdjustment = fontSize * 0.8;
                    }
                    maginVal = Math.max(0, maginVal - bulletSizeAdjustment);
                    marLStr += maginVal + "px;";
                } else {
                    maginVal = Math.abs(marginLeft + indent);
                    marLStr += maginVal + "px;";
                }
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
            if (marRNode !== undefined && isBullate) {
                let marginRight = parseInt(marRNode) * SLIDE_FACTOR;
                // 无论 RTL 还是 LTR，marR 都转换为 padding-right（右边距）
                marRStr = "padding-right: ";
                marRStr += Math.abs(0 - indent) + "px;";
            }


            return [marLStr, maginVal];
        }
// 提取图表标题样式
function extractChartTitleStyle(chartNode, warpObj) {
    const titleNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:title"]);
    if (!titleNode) return {};
    
    const style = {};
    
    // 提取标题文本
    // 尝试从富文本中提取
    const txPr = PPTXXmlUtils.getTextByPathList(titleNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            // 从段落中提取文本
            const r = PPTXXmlUtils.getTextByPathList(p, ["a:r"]);
            if (r) {
                // 可能有多个 run
                if (Array.isArray(r)) {
                    const textArray = r.map(run => PPTXXmlUtils.getTextByPathList(run, ["a:t"]));
                    style.text = textArray.filter(t => t).join('');
                } else {
                    style.text = PPTXXmlUtils.getTextByPathList(r, ["a:t"]);
                }
            }
            
            // 提取标题文本属性
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体粗细
                    if (defRPr["attrs"] && defRPr["attrs"]["b"] === "1") {
                        style.fontWeight = "bold";
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getColor(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    // 如果没有找到文本，尝试从字符串引用中提取
    if (!style.text) {
        const tx = PPTXXmlUtils.getTextByPathList(titleNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]);
        if (tx) {
            style.text = tx;
        }
    }
    
    return style;
}

// 提取图表区域样式
function extractChartAreaStyle(chartSpaceNode, warpObj) {
    const style = {};
    
    // 提取图表区域填充
    const spPr = PPTXXmlUtils.getTextByPathList(chartSpaceNode, ["c:spPr"]);
    if (spPr) {
        // 提取填充样式
        const fillType = getFillType(spPr);
        if (fillType === "SOLID_FILL") {
            const solidFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:solidFill"]);
            if (solidFill) {
                let fillColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (fillColor && !fillColor.startsWith('#')) {
                    fillColor = '#' + fillColor;
                }
                style.fillColor = fillColor;
            }
        } else if (fillType === "GRADIENT_FILL") {
            const gradFill = PPTXXmlUtils.getTextByPathList(spPr, ["a:gradFill"]);
            if (gradFill) {
                style.gradientFill = getGradientFill(gradFill, warpObj);
            }
        }
        
        // 提取边框样式
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let borderColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (borderColor && !borderColor.startsWith('#')) {
                    borderColor = '#' + borderColor;
                }
                style.borderColor = borderColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.borderWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
            }
        }
    }
    
    return style;
}

// 提取图表图例样式
function extractChartLegendStyle(chartNode, warpObj) {
    const legendNode = PPTXXmlUtils.getTextByPathList(chartNode, ["c:legend"]);
    if (!legendNode) return {};
    
    const style = {};
    
    // 提取图例位置
    if (legendNode["c:legendPos"]) {
        style.position = legendNode["c:legendPos"]["attrs"]["val"];
    }
    
    // 提取图例文本属性
    const txPr = PPTXXmlUtils.getTextByPathList(legendNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    return style;
}

// 提取图表轴样式
function extractChartAxisStyle(plotAreaNode, axisType, warpObj) {
    const axisNode = PPTXXmlUtils.getTextByPathList(plotAreaNode, [axisType]);
    if (!axisNode) return {};
    
    const style = {};
    
    // 提取轴文本属性
    const txPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:txPr"]);
    if (txPr) {
        const p = PPTXXmlUtils.getTextByPathList(txPr, ["a:p"]);
        if (p) {
            const pPr = PPTXXmlUtils.getTextByPathList(p, ["a:pPr"]);
            if (pPr) {
                const defRPr = PPTXXmlUtils.getTextByPathList(pPr, ["a:defRPr"]);
                if (defRPr) {
                    // 提取字体大小
                    if (defRPr["attrs"] && defRPr["attrs"]["sz"]) {
                        style.fontSize = parseFloat(defRPr["attrs"]["sz"]) / 100;
                    }
                    
                    // 提取字体颜色
                    const solidFill = PPTXXmlUtils.getTextByPathList(defRPr, ["a:solidFill"]);
                    if (solidFill) {
                        let color = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (color && !color.startsWith('#')) {
                            color = '#' + color;
                        }
                        style.color = color;
                    }
                }
            }
        }
    }
    
    // 提取轴线条样式
    const spPr = PPTXXmlUtils.getTextByPathList(axisNode, ["c:spPr"]);
    if (spPr) {
        const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
        if (ln) {
            const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
            if (solidFill) {
                let lineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                if (lineColor && !lineColor.startsWith('#')) {
                    lineColor = '#' + lineColor;
                }
                style.lineColor = lineColor;
            }
            if (ln["attrs"] && ln["attrs"]["w"]) {
                style.lineWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
            }
        }
    }
    
    // 提取轴网格线样式
    if (axisType === "c:valAx") {
        const majorGridlines = PPTXXmlUtils.getTextByPathList(axisNode, ["c:majorGridlines"]);
        if (majorGridlines) {
            const spPr = PPTXXmlUtils.getTextByPathList(majorGridlines, ["c:spPr"]);
            if (spPr) {
                const ln = PPTXXmlUtils.getTextByPathList(spPr, ["a:ln"]);
                if (ln) {
                    const solidFill = PPTXXmlUtils.getTextByPathList(ln, ["a:solidFill"]);
                    if (solidFill) {
                        let gridlineColor = getSolidFill(solidFill, undefined, undefined, warpObj);
                        if (gridlineColor && !gridlineColor.startsWith('#')) {
                            gridlineColor = '#' + gridlineColor;
                        }
                        style.gridlineColor = gridlineColor;
                    }
                    if (ln["attrs"] && ln["attrs"]["w"]) {
                        style.gridlineWidth = parseFloat(ln["attrs"]["w"]) / 9525; // 转换为像素
                    }
                }
            }
        }
    }
    
    return style;
}

// 辅助函数：获取颜色
function getColor(node, clrMap, phClr, warpObj) {
    if (node["a:solidFill"]) {
        return getSolidFill(node["a:solidFill"], clrMap, phClr, warpObj);
    }
    return "";
}

const PPTXStyleUtils = {
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
        getGradientFill,
        extractChartData,
        extractChartTitleStyle,
        extractChartAreaStyle,
        extractChartLegendStyle,
        extractChartAxisStyle,
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
        getColor
    };

export { PPTXStyleUtils };
export default PPTXStyleUtils;
