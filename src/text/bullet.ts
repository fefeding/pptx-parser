import { PPTXUtils } from '../core/utils';
import { PPTXColorUtils } from '../core/color';
import { PPTXLayoutUtils } from '../core/layout';
import { PPTXConstants } from '../core/constants';
import { PPTXTextStyleUtils } from './style.js';
import { TextUtils } from './text.js';

const slideFactor = PPTXConstants.SLIDE_FACTOR;

class PPTXBulletUtils {
    /**
     * 生成项目符号字符
     * @param {Object} node - 节点
     * @param {Number} i - 索引
     * @param {Object} spNode - 形状节点
     * @param {Object} textBodyNode - 文本主体节点
     * @param {Object} pFontStyle - 段落字体样式
     * @param {Number} idx - 索引
     * @param {String} type - 类型
     * @param {Object} warpObj - 包装对象
     * @returns {Array} 返回项目符号HTML、边距值、字体值
     */
    static genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
    //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
    ///////////////////////////////////////Amir///////////////////////////////
    var sldMstrTxtStyles: any = warpObj["slideMasterTextStyles"];
    var lstStyle: any = textBodyNode["a:lstStyle"];

    var rNode: any = PPTXUtils.getTextByPathList(node, ["a:r"]);
    if (rNode !== undefined && rNode.constructor === Array) {
        rNode = rNode[0]; //bullet only to first "a:r"
    }
    var lvl: any = parseInt(PPTXUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
    if (isNaN(lvl)) {
        lvl = 1;
    }
    var lvlStr: any = "a:lvl" + lvl + "pPr";
    var dfltBultColor: any;
    var dfltBultSize: any;
    var bultColor: any;
    var bultSize: any;
    var color_type: any;  // 修正拼写

    if (rNode !== undefined) {
        dfltBultColor = PPTXTextStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
        color_type = dfltBultColor[2];
        dfltBultSize = PPTXTextStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
    } else {
        return "";
    }
    //console.log("Bullet Size: " + bultSize);

    var bullet: any = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
    /////////////////////////////////////////////////////////////////

    var pPrNode: any = node["a:pPr"];
    var BullNONE: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buNone"]);
    if (BullNONE !== undefined) {
        return "";
    }

    var buType: any = "TYPE_NONE";

    var layoutMasterNode: any = PPTXLayoutUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
    var pPrNodeLaout: any = layoutMasterNode.nodeLaout;
    var pPrNodeMaster: any = layoutMasterNode.nodeMaster;

    var buChar: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
    var buNum: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
    var buPic: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
    if (buChar !== undefined) {
        buType = "TYPE_BULLET";
    }
    if (buNum !== undefined) {
        buType = "TYPE_NUMERIC";
    }
    if (buPic !== undefined) {
        buType = "TYPE_BULPIC";
    }

    var buFontSize: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
    if (buFontSize === undefined) {
        buFontSize = PPTXUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
        if (buFontSize !== undefined) {
            var prcnt: any = parseInt(buFontSize) / 100000;
            //dfltBultSize = XXpt
            //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
            var dfltBultSizeNoPt: any = parseInt(dfltBultSize, 10);
            bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
        }
    } else {
        bultSize = (parseInt(buFontSize) / 100) * PPTXConstants.FONT_SIZE_FACTOR + "px";
    }

    //get definde bullet COLOR
    var buClrNode: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buClr"]);

    if (buChar === undefined && buNum === undefined && buPic === undefined) {

        if (lstStyle !== undefined) {
            BullNONE = PPTXUtils.getTextByPathList(lstStyle, [lvlStr,"a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = PPTXUtils.getTextByPathList(lstStyle, [lvlStr,"a:buChar", "attrs", "char"]);
            buNum = PPTXUtils.getTextByPathList(lstStyle, [lvlStr,"a:buAutoNum", "attrs", "type"]);
            buPic = PPTXUtils.getTextByPathList(lstStyle, [lvlStr,"a:buBlip"]);
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

        if (pPrNodeLaout !== undefined) {
            BullNONE = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }
            buType = "TYPE_NONE";
            buChar = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
            buNum = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
            buPic = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
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
                BullNONE = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                if (BullNONE !== undefined) {
                    return "";
                }
                buType = "TYPE_NONE";
                buChar = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                buNum = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                buPic = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
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
    var getRtlVal: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
    if (getRtlVal === undefined) {
        getRtlVal = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
        if (getRtlVal === undefined && type != "shape") {
            getRtlVal = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
        }
    }
    var isRTL: any = false;
    if (getRtlVal !== undefined && getRtlVal == "1") {
        isRTL = true;
    }
    //align
    var alignNode: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
    if (alignNode === undefined) {
        alignNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
        if (alignNode === undefined) {
            alignNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
        }
    }
    //indent?
    var indentNode: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
    if (indentNode === undefined) {
        indentNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
        if (indentNode === undefined) {
            indentNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
        }
    }
    var indent: any = 0;
    if (indentNode !== undefined) {
        indent = parseInt(indentNode) * PPTXConstants.SLIDE_FACTOR;
    }
    //marL
    var marLNode: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
    if (marLNode === undefined) {
        marLNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
        if (marLNode === undefined) {
            marLNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
        }
    }
    //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
    if (marLNode !== undefined) {
        var marginLeft: any = parseInt(marLNode) * PPTXConstants.SLIDE_FACTOR;
        if (isRTL) {// && alignNode == "r") {
            marLStr = "padding-right:";// "margin-right: ";
        } else {
            marLStr = "padding-left:";//"margin-left: ";
        }
        margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
        marLStr += margin_val + "px;";
    }
    
    //marR?
    var marRNode: any = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
    if (marRNode === undefined && marLNode === undefined) {

        marRNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
        if (marRNode === undefined) {
            marRNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
        }
    }
    if (marRNode !== undefined) {
        var marginRight: any = parseInt(marRNode) * PPTXConstants.SLIDE_FACTOR;
        if (isRTL) {// && alignNode == "r") {
            marLStr = "padding-right:";// "margin-right: ";
        } else {
            marLStr = "padding-left:";//"margin-left: ";
        }
        marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
    }

    if (buType != "TYPE_NONE") {
        //var buFontAttrs = PPTXUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
    }
    //console.log("Bullet Type: " + buType);
    //console.log("NumericTypr: " + buNum);
    //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
    //get definde bullet COLOR
    if (buClrNode === undefined){
        //lstStyle
        buClrNode = PPTXUtils.getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
    }
    if (buClrNode === undefined) {
        buClrNode = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buClr"]);
        if (buClrNode === undefined) {
            buClrNode = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buClr"]);
        }
    }
    var defBultColor: any;
    if (buClrNode !== undefined) {
        defBultColor = PPTXColorUtils.getSolidFill(buClrNode, undefined, undefined, warpObj);
    } else {
        if (pFontStyle !== undefined) {
            //console.log("genBuChar pFontStyle: ", pFontStyle)
            defBultColor = PPTXColorUtils.getSolidFill(pFontStyle, undefined, undefined, warpObj);
        }
    }
    if (defBultColor === undefined || defBultColor == "NONE") {
        bultColor = dfltBultColor;
    } else {
        bultColor = [defBultColor, "", "solid"];
        color_type = "solid";
    }
    //console.log("genBuChar node:", node, "pPrNode", pPrNode, " buClrNode: ", buClrNode, "defBultColor:", defBultColor,"dfltBultColor:" , dfltBultColor , "bultColor:", bultColor)

    //console.log("genBuChar: buClrNode: ", buClrNode, "bultColor", bultColor)
    //get definde bullet SIZE
    if (buFontSize === undefined) {
        buFontSize = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = PPTXUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                const prcnt = parseInt(buFontSize) / 100000;
                //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                const dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                bultSize = prcnt * dfltBultSizeNoPt + "px";// + "pt";
            }
        }else{
            bultSize = (parseInt(buFontSize) / 100) * PPTXConstants.FONT_SIZE_FACTOR + "px";
        }
    }
    if (buFontSize === undefined) {
        buFontSize = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
        if (buFontSize === undefined) {
            buFontSize = PPTXUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
            if (buFontSize !== undefined) {
                const prcnt = parseInt(buFontSize) / 100000;
                //dfltBultSize = XXpt
                //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                const dfltBultSizeNoPt = parseInt(dfltBultSize, 10);
                bultSize = prcnt * dfltBultSizeNoPt + "px";// + "pt";
            }
        } else {
            bultSize = (parseInt(buFontSize) / 100) * PPTXConstants.FONT_SIZE_FACTOR + "px";
        }
    }
    if (buFontSize === undefined) {
        bultSize = dfltBultSize;
    }
    font_val = parseInt(bultSize, 10);
    ////////////////////////////////////////////////////////////////////////
    if (buType == "TYPE_BULLET") {
        var typefaceNode: any = PPTXUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
        var typeface: any = "";
        if (typefaceNode !== undefined) {
            typeface = "font-family: " + typefaceNode;
        }
        // var marginLeft = parseInt(PPTXUtils.getTextByPathList(marLNode)) * slideFactor;
        // var marginRight = parseInt(PPTXUtils.getTextByPathList(marRNode)) * slideFactor;
        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // }
        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }

        bullet = "<div style='height: 100%;" + typeface + ";" +
            marLStr + marRStr +
            "font-size:" + bultSize + ";" ;
        
        //bullet += "display: table-cell;";
        //"line-height: 0px;";
        if (color_type == "solid") {
            if (bultColor[0] !== undefined && bultColor[0] != "") {
                bullet += "color:#" + bultColor[0] + "; ";
            }
            if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                bullet += "text-shadow:" + bultColor[1] + ";";
            }
            //no highlight/background-color to bullet
            // if (bultColor[3] !== undefined && bultColor[3] != "") {
            //     styleText += "background-color: #" + bultColor[3] + ";";
            // }
        } else if (color_type == "pattern" || color_type == "pic" || color_type == "gradient") {
            if (color_type == "pattern") {
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
            } else if (color_type == "pic") {
                bullet += bultColor[0] + ";";
                // bullet += "-webkit-background-clip: text;" +
                //     "background-clip: text;" +
                //     "color: transparent;" +
                //     "-webkit-text-stroke: " + bultColor[1].border + ";";

            } else if (color_type == "gradient") {

                var colorAry: any = bultColor[0].color;
                var rot: any = bultColor[0].rot;

                bullet += "background: linear-gradient(" + rot + "deg,";
                for (let j = 0; j < colorAry.length; j++) {
                    if (j == colorAry.length - 1) {
                        bullet += "#" + colorAry[j] + ");";
                    } else {
                        bullet += "#" + colorAry[j] + ", ";
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
        var isIE11: any = !!(window as any).MSInputMethodContext && !!(document as any).documentMode;
        var htmlBu: any = buChar;

        if (!isIE11) {
            //ie11 does not support unicode ?
            htmlBu = TextUtils.getHtmlBullet(typefaceNode, buChar);
        }
        bullet += "'><div style='line-height: " + (font_val/2) + "px;'>" + htmlBu + "</div></div>"; //font_val
        //} 
        // else {
        //     marginLeft = 328600 * slideFactor * lvl;
        //
        //     bullet = "<div style='" + marLStr + "'>" + buChar + "</div>";
        // }
    } else if (buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
        //if (buFontAttrs !== undefined) {
        // var marginLeft = parseInt(PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
        // var marginRight = parseInt(buFontAttrs["pitchFamily"]);

        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // }
        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }
        //var typeface = buFontAttrs["typeface"];

        bullet = "<div style='height: 100%;" + marLStr + marRStr +
            "color:#" + bultColor[0] + ";" +
            "font-size:" + bultSize + ";";// +
        //"line-height: 0px;";
        if (isRTL) {
            bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; // float: right;
        } else {
            bullet += "display: inline-block;white-space: nowrap ;direction:ltr;"; //float: left;
        }
        bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
        // } else {
        //     marginLeft = 328600 * slideFactor * lvl;
        //     bullet = "<div style='margin-left: " + marginLeft + "px;";
        //     if (isRTL) {
        //         bullet += " float: right; direction:rtl;";
        //     } else {
        //         bullet += " float: left; direction:ltr;";
        //     }
        //     bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
        // }

    } else if (buType == "TYPE_BULPIC") { //PIC BULLET
        // var marginLeft = parseInt(PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
        // var marginRight = parseInt(PPTXUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * slideFactor;

        // if (isNaN(marginRight)) {
        //     marginRight = 0;
        // }
        // //console.log("marginRight: "+marginRight)
        // //buPic
        // if (isNaN(marginLeft)) {
        //     marginLeft = 328600 * slideFactor;
        // } else {
        //     marginLeft = 0;
        // }
        //var buPicId = PPTXUtils.getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
        var buPicId: any = PPTXUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
        var svgPicPath: any = "";
        var buImg: any;
        if (buPicId !== undefined) {
            //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
            //buImg = warpObj["zip"].file(svgPicPath).asText();
            //}else{
            //buPicId = PPTXUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
            var imgPath: any = warpObj["slideResObj"][buPicId]["target"];
            //console.log("imgPath: ", imgPath);
            // 尝试解析图片路径，处理相对路径问题
            var imgFile: any = warpObj["zip"].file(imgPath);
            if (!imgFile && !imgPath.startsWith("ppt/")) {
                // 尝试添加 ppt/ 前缀
                imgFile = warpObj["zip"].file("ppt/" + imgPath);
            }
            if (!imgFile) {
                buImg = "&#8227;";
            } else {
                var imgArrayBuffer: any = imgFile.asArrayBuffer();
                var imgExt: any = imgPath.split(".").pop();
                var imgMimeType: any = PPTXUtils.getMimeType(imgExt);
                buImg = "<img src='" + PPTXUtils.arrayBufferToBlobUrl(imgArrayBuffer, imgMimeType) + "' style='width: 100%;'/>";// height: 100%
                //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
            }
        } else {
            buImg = "&#8227;";
        }
        bullet = "<div style='height: 100%;" + marLStr + marRStr +
            "width:" + bultSize + ";display: inline-block; ";// +
        //"line-height: 0px;";
        if (isRTL) {
            bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
        }
        bullet += "'>" + buImg + "  </div>";
        //////////////////////////////////////////////////////////////////////////////////////
    }
    // else {
    //     bullet = "<div style='margin-left: " + 328600 * slideFactor * lvl + "px" +
    //         "; margin-right: " + 0 + "px;'></div>";
    // }
    //console.log("genBuChar: width: ", $(bullet).outerWidth())
    return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
}
};

export { PPTXBulletUtils };

// Also export to global scope for backward compatibility
// if (typeof window !== 'undefined') {
//     window.PPTXBulletUtils = PPTXBulletUtils;
// } // Removed for ES modules
