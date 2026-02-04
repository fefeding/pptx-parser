

var PPTXTextUtils = (function() {
    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;
    
    var rtl_langs_array = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR","ur-PK"]
    
    var is_first_br = false;

    function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
            var text = "";
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            if (textBodyNode === undefined) {
                return text;
            }
            //rtl : <p:txBody>
            //          <a:bodyPr wrap="square" rtlCol="1">

            var pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
            //console.log("genTextBody spNode: ", PPTXXmlUtils.getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

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
                    is_first_br = true;
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
                var styleText = "";
                var marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
                if (marginsVer != "") {
                    styleText = marginsVer;
                }
                if (type == "body" || type == "obj" || type == "shape") {
                    styleText += "font-size: 0px;";
                    //styleText += "line-height: 0;";
                    styleText += "font-weight: 100;";
                    styleText += "font-style: normal;";
                }
                var cssName = "";

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
                var prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                var prg_height_node;// = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
                var sld_prg_width = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * slideFactor) + "px;") : "width:inherit;");
                var sld_prg_height = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * slideFactor) + "px;") : "");
                var prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
                text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + " " +
                    prg_dir + " " + cssName + "' >";
                var buText_ary = genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
                var isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
                var bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
                text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
                //get text margin 
                var margin_ary = PPTXStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj);
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
                    var prgr_text = genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
                    if (isBullate) {
                        var txt_obj = $(prgr_text)
                            .css({ 'position': 'absolute', 'float': 'left', 'white-space': 'nowrap', 'visibility': 'hidden' })
                            .appendTo($('body'));
                        total_text_len += txt_obj.outerWidth();
                        txt_obj.remove();
                    }
                    prgrph_text += prgr_text;
                } else if (rNode !== undefined) {
                    // with multi r
                    for (var j = 0; j < rNode.length; j++) {
                        var prgr_text = genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
                        if (isBullate) {
                            var txt_obj = $(prgr_text)
                                .css({ 'position': 'absolute', 'float': 'left', 'white-space': 'nowrap', 'visibility': 'hidden'})
                                .appendTo($('body'));
                            total_text_len += txt_obj.outerWidth();
                            txt_obj.remove();
                        }
                        prgrph_text += prgr_text;
                    }
                }

                prg_width_node = parseInt(prg_width_node) * slideFactor - bu_width - mrgin_val;
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
        }
        
        function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
            //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
            ///////////////////////////////////////Amir///////////////////////////////
            var sldMstrTxtStyles = warpObj["slideMasterTextStyles"];
            var lstStyle = textBodyNode["a:lstStyle"];

            var rNode = PPTXXmlUtils.getTextByPathList(node, ["a:r"]);
            if (rNode !== undefined && rNode.constructor === Array) {
                rNode = rNode[0]; //bullet only to first "a:r"
            }
            var lvl = parseInt (PPTXXmlUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
            if (isNaN(lvl)) {
                lvl = 1;
            }
            var lvlStr = "a:lvl" + lvl + "pPr";
            var dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;

            if (rNode !== undefined) {
                dfltBultColor = PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
                color_tye = dfltBultColor[2];
                dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
            } else {
                return "";
            }
            //console.log("Bullet Size: " + bultSize);

            var bullet = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
            /////////////////////////////////////////////////////////////////


            var pPrNode = node["a:pPr"];
            var BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }

            var buType = "TYPE_NONE";

            var layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            var buChar = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
            var buNum = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
            var buPic = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }

            var buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    var prcnt = parseInt(buFontSize) / 100000;
                    //dfltBultSize = XXpt
                    //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                    bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
            }

            //get definde bullet COLOR
            var buClrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buClr"]);


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
            var getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            var isRTL = false;
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
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
            //marL
            var marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }
            //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
            if (marLNode !== undefined) {
                var marginLeft = parseInt(marLNode) * slideFactor;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
                marLStr += margin_val + "px;";
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
            if (marRNode !== undefined) {
                var marginRight = parseInt(marRNode) * slideFactor;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
            }

            if (buType != "TYPE_NONE") {
                //var buFontAttrs = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
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
            var defBultColor;
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
                        var prcnt = parseInt(buFontSize) / 100000;
                        //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                }else{
                    bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
                }
            }
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        var prcnt = parseInt(buFontSize) / 100000;
                        //dfltBultSize = XXpt
                        //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                } else {
                    bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
                }
            }
            if (buFontSize === undefined) {
                bultSize = dfltBultSize;
            }
            font_val = parseInt(bultSize, "px");
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                var typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
                var typeface = "";
                if (typefaceNode !== undefined) {
                    typeface = "font-family: " + typefaceNode;
                }
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(marLNode)) * slideFactor;
                // var marginRight = parseInt (PPTXXmlUtils.getTextByPathList(marRNode)) * slideFactor;
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
                if (color_tye == "solid") {
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

                        var colorAry = bultColor[0].color;
                        var rot = bultColor[0].rot;

                        bullet += "background: linear-gradient(" + rot + "deg,";
                        for (var i = 0; i < colorAry.length; i++) {
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
                var isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
                var htmlBu = buChar;

                if (!isIE11) {
                    //ie11 does not support unicode ?
                    htmlBu = getHtmlBullet(typefaceNode, buChar);
                }
                bullet += "'><div style='line-height: " + (font_val/2) + "px;'>" + htmlBu + "</div></div>"; //font_val
                //} 
                // else {
                //     marginLeft = 328600 * slideFactor * lvl;

                //     bullet = "<div style='" + marLStr + "'>" + buChar + "</div>";
                // }
            } else if (buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
                //if (buFontAttrs !== undefined) {
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
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
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
                // var marginRight = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * slideFactor;

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
                //var buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
                var buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                var svgPicPath = "";
                var buImg;
                if (buPicId !== undefined) {
                    //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                    //buImg = warpObj["zip"].file(svgPicPath).asText();
                    //}else{
                    //buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                    var imgPath = (warpObj["slideResObj"][buPicId] !== undefined) ? warpObj["slideResObj"][buPicId]["target"] : undefined;
                    //console.log("imgPath: ", imgPath);
                    if (imgPath === undefined) {
                        console.warn("Bullet image reference not found for buPicId:", buPicId);
                        buImg = "";
                    } else {
                        var imgFile = warpObj["zip"].file(imgPath);
                        if (imgFile === null) {
                            console.warn("Bullet image file not found:", imgPath);
                            buImg = "";
                        } else {
                            var imgArrayBuffer = imgFile.asArrayBuffer();
                            var imgExt = imgPath.split(".").pop();
                            var imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                            buImg = "<img src='data:" + imgMimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>"// height: 100%
                            //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
                        }
                    }
                }
                if (buPicId === undefined) {
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
                default:
                    if (/*typefaceNode == "Wingdings" ||*/ typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3"){
                        var wingCharCode =  getDingbatToUnicode(typefaceNode, buChar);
                        if (wingCharCode !== null){
                            return "&#" + wingCharCode + ";";
                        }
                    }
                    return "&#" + (buChar.charCodeAt(0)) + ";";
            }
        }
        function getDingbatToUnicode(typefaceNode, buChar){
            if (dingbat_unicode){
                var dingbat_code = buChar.codePointAt(0) & 0xFFF;
                var char_unicode = null;
                var len = dingbat_unicode.length;
                var i = 0;
                while (len--) {
                    // blah blah
                    var item = dingbat_unicode[i];
                    if (item.f == typefaceNode && item.code == dingbat_code) {
                        char_unicode = item.unicode;
                        break;
                    }
                    i++;
                }
                return char_unicode
            }
        }

        function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
            //https://codepen.io/imdunn/pen/GRgwaye ?
            var text_style = "";
            var lstStyle = textBodyNode["a:lstStyle"];
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            var text = node["a:t"];
            //var text_count = text.length;

            var openElemnt = "<span";//"<bdi";
            var closeElemnt = "</span>";// "</bdi>";
            var styleText = "";
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

            var pPrNode = pNode["a:pPr"];
            //lvl
            var lvl = 1;
            var lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, warpObj);
            var layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            //Language
            var lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
            var isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1)?true:false;
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

            var linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
            var linkTooltip = "";
            var defLinkClr;
            if (linkID !== undefined) {
                linkTooltip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
                if (linkTooltip !== undefined) {
                    linkTooltip = "title='" + linkTooltip + "'";
                }
                defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

                var linkClrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                var rPrlinkClr = PPTXStyleUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);


                //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
                if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                    defLinkClr = rPrlinkClr;
                }

            }
            /////////////////////////////////////////////////////////////////////////////////////
            //getFontColor
            var fontClrPr = PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
            var fontClrType = fontClrPr[2];
            //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
            if (fontClrType == "solid") {
                if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                    styleText += "color: #" + fontClrPr[0] + ";";
                }
                else if (linkID !== undefined && defLinkClr !== undefined) {
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
            var font_size = PPTXStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
            //text_style += "font-size:" + font_size + ";"
            
            text_style += "font-size:" + font_size + ";" +
                // marLStr +
                "font-family:" + PPTXStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
                "font-weight:" + PPTXStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
                "font-style:" + PPTXStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
                "text-decoration:" + PPTXStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
                "text-align:" + PPTXStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
                "vertical-align:" + PPTXStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
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

            //     //"direction:" + dirStr + ";";
            //if (rNodeLength == 1 || rIndex == 0 ){
            //styleText += "display: table-cell;white-space: nowrap;";
            //}
            var highlight = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
            if (highlight !== undefined) {
                styleText += "background-color:#" + PPTXStyleUtils.getSolidFill(highlight, undefined, undefined, warpObj) + ";";
                //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
            }

            //letter-spacing:
            var spcNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
                if (spcNode === undefined) {
                    spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
                }
            }
            if (spcNode !== undefined) {
                var ltrSpc = parseInt(spcNode) / 100; //pt
                styleText += "letter-spacing: " + ltrSpc + "px;";// + "pt;";
            }

            //Text Cap Types
            var capNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
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

            var cssName = "";

            if (styleText in warpObj.styleTable) {
                cssName = warpObj.styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                warpObj.styleTable[styleText] = {
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
                linkURL = PPTXXmlUtils.escapeHtml(linkURL);
                return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + " target='_blank'>" +
                        text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a>" + closeElemnt;
            } else {
                return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + closeElemnt;//"</bdi>";
            }

        }

    
        function genChart(node, warpObj) {

            var order = node["attrs"]["order"];
            var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            var result = "<div id='chart" + warpObj.chartID + "' class='block content' style='" +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) + PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                " z-index: " + order + ";'></div>";

            var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
            var refName = warpObj["slideResObj"][rid]["target"];
            var content = PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
            var plotArea = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

            var chartData = null;
            for (var key in plotArea) {
                switch (key) {
                    case "c:lineChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "lineChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:barChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "barChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pieChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "pieChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pie3DChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "pie3DChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:areaChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "areaChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:scatterChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "scatterChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:catAx":
                        break;
                    case "c:valAx":
                        break;
                    default:
                }
            }

            if (chartData !== null) {
                warpObj.MsgQueue.push(chartData);
            }

            warpObj.chartID++;
            return result;
        }

        function genDiagram(node, warpObj, source, sType) {
            //console.log(warpObj)
            //PPTXXmlUtils.readXmlFile(zip, sldFileName)
            /**files define the diagram:
             * 1-colors#.xml,
             * 2-data#.xml, 
             * 3-layout#.xml,
             * 4-quickStyle#.xml.
             * 5-drawing#.xml, which Microsoft added as an extension for persisting diagram layout information.
             */
            ///get colors#.xml, data#.xml , layout#.xml , quickStyle#.xml
            var order = node["attrs"]["order"];
            var zip = warpObj["zip"];
            var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            var dgmRelIds = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);
            //console.log(dgmRelIds)
            var dgmClrFileId = dgmRelIds["r:cs"];
            var dgmDataFileId = dgmRelIds["r:dm"];
            var dgmLayoutFileId = dgmRelIds["r:lo"];
            var dgmQuickStyleFileId = dgmRelIds["r:qs"];
            var dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target,
                dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target,
                dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
            dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;
            //console.log("dgmClrFileName: " , dgmClrFileName,", dgmDataFileName: ",dgmDataFileName,", dgmLayoutFileName: ",dgmLayoutFileName,", dgmQuickStyleFileName: ",dgmQuickStyleFileName);
            var dgmClr = PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
            var dgmData = PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
            var dgmLayout = PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
            var dgmQuickStyle = PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);
            //console.log(dgmClr,dgmData,dgmLayout,dgmQuickStyle)
            ///get drawing#.xml
            // var dgmDrwFileName = "";
            // var dataModelExt = PPTXXmlUtils.getTextByPathList(dgmData, ["dgm:dataModel", "dgm:extLst", "a:ext", "dsp:dataModelExt", "attrs"]);
            // if (dataModelExt !== undefined) {
            //     var dgmDrwFileId = dataModelExt["relId"];
            //     dgmDrwFileName = warpObj["slideResObj"][dgmDrwFileId]["target"];
            // }
            // var dgmDrwFile = "";
            // if (dgmDrwFileName != "") {
            //     dgmDrwFile = PPTXXmlUtils.readXmlFile(zip, dgmDrwFileName);
            // }
            // var dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(dgmDrwFile, ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
            //var dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(warpObj["digramFileContent"], ["dsp:drawing", "dsp:spTree", "dsp:sp"]);
            var dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
            var rslt = "";
            if (dgmDrwSpArray !== undefined) {
                var dgmDrwSpArrayLen = dgmDrwSpArray.length;
                for (var i = 0; i < dgmDrwSpArrayLen; i++) {
                    var dspSp = dgmDrwSpArray[i];
                    // var dspSpObjToStr = JSON.stringify(dspSp);
                    // var pSpStr = dspSpObjToStr.replace(/dsp:/g, "p:");
                    // var pSpStrToObj = JSON.parse(pSpStr);
                    //console.log("pSpStrToObj[" + i + "]: ", pSpStrToObj);
                    //rslt += PPTXNodeUtils.processSpNode(pSpStrToObj, node, warpObj, "diagramBg", sType)
                    rslt += PPTXNodeUtils.processSpNode(dspSp, node, warpObj, "diagramBg", sType)
                }
                // dgmDrwFile: "dsp:"-> "p:"
            }

            return "<div class='block diagram-content' style='" +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined, sType) +
                PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                "'>" + rslt + "</div>";
        }
        
        function genTable(node, warpObj) {
            var order = node["attrs"]["order"];
            var tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            var getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
            var getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
            var tblDir = "";
            if (getTblPr !== undefined) {
                var isRTL = getTblPr["attrs"]["rtl"];
                tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
            }
            var firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
            var firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
            var lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
            var lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
            var bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
            var bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
            //console.log("getTblPr: ", getTblPr);
            var tblStylAttrObj = {
                isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
                isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
                isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
                isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
                isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
                isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
            }

            var thisTblStyle;
            var tbleStyleId = getTblPr["a:tableStyleId"];
            if (tbleStyleId !== undefined) {
                var tbleStylList = tableStyles["a:tblStyleLst"]["a:tblStyle"];
                if (tbleStylList !== undefined) {
                    if (tbleStylList.constructor === Array) {
                        for (var k = 0; k < tbleStylList.length; k++) {
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
            var tblStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
            var tblBorderStyl = PPTXXmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
            var tbl_borders = "";
            if (tblBorderStyl !== undefined) {
                tbl_borders = PPTXStyleUtils.getTableBorders(tblBorderStyl, warpObj);
            }
            var tbl_bgcolor = "";
            var tbl_opacity = 1;
            var tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
            //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
            if (tbl_bgFillschemeClr !== undefined) {
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgFillschemeClr === undefined) {
                tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgcolor !== "") {
                tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
            }
            ////////////////////////////////////////////////////////////////////////////////////////////
            var tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) +
                PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                " z-index: " + order + ";" +
                tbl_borders + ";" +
                tbl_bgcolor + "'>";

            var trNodes = tableNode["a:tr"];
            if (trNodes.constructor !== Array) {
                trNodes = [trNodes];
            }
            //if (trNodes.constructor === Array) {
                //multi rows
                var totalrowSpan = 0;
                var rowSpanAry = [];
                for (var i = 0; i < trNodes.length; i++) {
                    //////////////rows Style ////////////Amir
                    var rowHeightParam = trNodes[i]["attrs"]["h"];
                    var rowHeight = 0;
                    var rowsStyl = "";
                    if (rowHeightParam !== undefined) {
                        rowHeight = parseInt(rowHeightParam) * slideFactor;
                        rowsStyl += "height:" + rowHeight + "px;";
                    }
                    var fillColor = "";
                    var row_borders = "";
                    var fontClrPr = "";
                    var fontWeight = "";
                    var band_1H_fillColor;
                    var band_2H_fillColor;

                    if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontColor = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontColor !== undefined) {
                                fontClrPr = local_fontColor;
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight
                            }
                        }
                    }

                    if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
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
                            var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== "") {
                                    fillColor = local_fillColor;
                                    band_2H_fillColor = local_fillColor;
                                }
                            }


                            var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                        if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                            var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== undefined) {
                                    fillColor = local_fillColor;
                                    band_1H_fillColor = local_fillColor;
                                }
                            }
                            var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                                var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                                if (local_fontWeight != "") {
                                    fontWeight = local_fontWeight;
                                }
                            }
                        }

                    }
                    //last row
                    if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                            // var local_colorOpacity = getColorOpacity(bgFillschemeClr);
                            // if(local_colorOpacity !== undefined){
                            //     colorOpacity = local_colorOpacity;
                            // }
                        }
                        var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }
                    rowsStyl += ((row_borders !== undefined) ? row_borders : "");
                    rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
                    rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
                    if (fillColor !== undefined && fillColor != "") {
                        //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + "," + colorOpacity + ");";
                        rowsStyl += "background-color: #" + fillColor + ";";
                    }
                    tableHtml += "<tr style='" + rowsStyl + "'>";
                    ////////////////////////////////////////////////

                    var tcNodes = trNodes[i]["a:tc"];
                    if (tcNodes !== undefined) {
                        if (tcNodes.constructor === Array) {
                            //multi columns
                            var j = 0;
                            if (rowSpanAry.length == 0) {
                                rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(function () { return 0 });
                            }
                            var totalColSpan = 0;
                            while (j < tcNodes.length) {
                                if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                                    var a_sorce;
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

                                            var aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
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

                                    var cellParmAry = getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj)
                                    var text = cellParmAry[0];
                                    var colStyl = cellParmAry[1];
                                    var cssName = cellParmAry[2];
                                    var rowSpan = cellParmAry[3];
                                    var colSpan = cellParmAry[4];



                                    if (rowSpan !== undefined) {
                                        totalrowSpan++;
                                        rowSpanAry[j] = parseInt(rowSpan) - 1;
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                            parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                    } else if (colSpan !== undefined) {
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                            parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                        totalColSpan = parseInt(colSpan) - 1;
                                    } else {
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
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

                            var a_sorce;
                            if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:firstCol";

                            } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                                var aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
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


                            var cellParmAry = getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj)
                            var text = cellParmAry[0];
                            var colStyl = cellParmAry[1];
                            var cssName = cellParmAry[2];
                            var rowSpan = cellParmAry[3];

                            if (rowSpan !== undefined) {
                                tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                            } else {
                                tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
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
            //var text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
            var rowSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
            var colSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
            var vMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
            var hMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
            var colStyl = "word-wrap: break-word;";
            var colWidth;
            var celFillColor = "";
            var col_borders = "";
            var colFontClrPr = "";
            var colFontWeight = "";
            var lin_bottm = "",
                lin_top = "",
                lin_left = "",
                lin_right = "",
                lin_bottom_left_to_top_right = "",
                lin_top_left_to_bottom_right = "";
            
            var colSapnInt = parseInt(colSpan);
            var total_col_width = 0;
            if (!isNaN(colSapnInt) && colSapnInt > 1){
                for (var k = 0; k < colSapnInt ; k++) {
                    total_col_width += parseInt (PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
                }
            }else{
                total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
            }
            

            var text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width);//tableStyles

            if (total_col_width != 0 /*&& row_idx == 0*/) {
                colWidth = parseInt(total_col_width) * slideFactor;
                colStyl += "width:" + colWidth + "px;";
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
                var bottom_line_border = PPTXStyleUtils.getBorder(lin_bottm, undefined, false, "", warpObj)
                if (bottom_line_border != "") {
                    colStyl += "border-bottom:" + bottom_line_border + ";";
                }
            }
            if (lin_top !== undefined && lin_top != "") {
                var top_line_border = PPTXStyleUtils.getBorder(lin_top, undefined, false, "", warpObj);
                if (top_line_border != "") {
                    colStyl += "border-top: " + top_line_border + ";";
                }
            }
            if (lin_left !== undefined && lin_left != "") {
                var left_line_border = PPTXStyleUtils.getBorder(lin_left, undefined, false, "", warpObj)
                if (left_line_border != "") {
                    colStyl += "border-left: " + left_line_border + ";";
                }
            }
            if (lin_right !== undefined && lin_right != "") {
                var right_line_border = PPTXStyleUtils.getBorder(lin_right, undefined, false, "", warpObj)
                if (right_line_border != "") {
                    colStyl += "border-right:" + right_line_border + ";";
                }
            }

            //cell fill color custom
            var getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
            if (getCelFill !== undefined && getCelFill != "") {
                var cellObj = {
                    "p:spPr": getCelFill
                };
                celFillColor = PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide")
            }

            //cell fill color theme
            if (celFillColor == "" || celFillColor == "background-color: inherit;") {
                var bgFillschemeClr;
                if (cellSource !== undefined)
                    bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        celFillColor = " background-color: #" + local_fillColor + ";";
                    }
                }
            }
            var cssName = "";
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
            // var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
            // if (borderStyl !== undefined) {
            //     var local_col_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
            //     if (local_col_borders != "") {
            //         col_borders = local_col_borders;
            //     }
            // }
            // if (col_borders != "") {
            //     colStyl += col_borders;
            // }

            //Text style
            var rowTxtStyl;
            if (cellSource !== undefined) {
                rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
            }
            // if (rowTxtStyl === undefined) {
            //     rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            // }
            if (rowTxtStyl !== undefined) {
                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    colFontClrPr = local_fontClrPr;
                }
                var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    colFontWeight = local_fontWeight;
                }
            }
            colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
            colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

            return [text, colStyl, cssName, rowSpan, colSpan];
        }
    return {
        genTextBody,
        genBuChar,
        getHtmlBullet,
        getDingbatToUnicode,
        genSpanElement,
        genChart,
        genDiagram,
        genTable,
        getTableCellParams,
    };
})();

window.PPTXTextUtils = PPTXTextUtils;
