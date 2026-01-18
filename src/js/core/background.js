import { PPTXUtils } from './utils.js';
import { PPTXColorUtils } from './color.js';
var PPTXBackgroundUtils = {};
    /**
 * 获取幻灯片背景
 * @param {Object} warpObj - 包装对象,包含幻灯片内容
 * @param {Object} slideSize - 幻灯片尺寸
 * @param {number} index - 幻灯片索引
 * @param {Function} processNodesInSlide - 处理幻灯片中节点的回调函数
 * @returns {string} 背景HTML字符串
 */
PPTXBackgroundUtils.getBackground = function(warpObj, slideSize, index, processNodesInSlide) {
    var slideContent = warpObj["slideContent"];
    var slideLayoutContent = warpObj["slideLayoutContent"];
    var slideMasterContent = warpObj["slideMasterContent"];

    var nodesSldLayout = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
    var nodesSldMaster = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);

    var showMasterSp = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
    var bgColor = this.getSlideBackgroundFill(warpObj, index);
    var result = "<div class='slide-background-" + index + "' style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;" + bgColor + "'>";
    var node_ph_type_ary = [];
    if (nodesSldLayout !== undefined) {
        for (var nodeKey in nodesSldLayout) {
            if (nodesSldLayout[nodeKey].constructor === Array) {
                for (var i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                    var ph_type = PPTXUtils.getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    if (ph_type != "pic") {
                        result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg");
                    }
                }
            } else {
                var ph_type = PPTXUtils.getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (ph_type != "pic") {
                    result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg");
                }
            }
        }
    }
    if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
        for (var nodeKey in nodesSldMaster) {
            if (nodesSldMaster[nodeKey].constructor === Array) {
                for (var i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                    var ph_type = PPTXUtils.getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg");
                }
            } else {
                var ph_type = PPTXUtils.getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg");
            }
        }
    }
    return result + "</div>";
};

    /**
 * 获取幻灯片背景填充样式
 * @param {Object} warpObj - 包装对象
 * @param {number} index - 幻灯片索引
 * @returns {string} 背景CSS样式字符串
 */
PPTXBackgroundUtils.getSlideBackgroundFill = function(warpObj, index) {
    var slideContent = warpObj["slideContent"];
    var slideLayoutContent = warpObj["slideLayoutContent"];
    var slideMasterContent = warpObj["slideMasterContent"];

    var bgPr = PPTXUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
    var bgRef = PPTXUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
    var bgcolor;

    // 检查幻灯片级别的背景
    if (bgPr !== undefined) {
        bgcolor = this._getBgFillFromPr(bgPr, slideContent, slideLayoutContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, undefined, index);
    } else if (bgRef !== undefined) {
        bgcolor = this._getBgFillFromRef(bgRef, slideContent, slideLayoutContent, slideMasterContent, warpObj);
    }
    else {
        // 检查幻灯片布局级别的背景
        bgPr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
        bgRef = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);

        if (bgPr !== undefined) {
            var clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (clrMapOvr === undefined) {
                clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
            bgcolor = this._getBgFillFromPr(bgPr, slideLayoutContent, slideLayoutContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, clrMapOvr, index);
        } else if (bgRef !== undefined) {
            var clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (clrMapOvr === undefined) {
                clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
            bgcolor = this._getBgFillFromRef(bgRef, slideLayoutContent, slideLayoutContent, slideMasterContent, warpObj, clrMapOvr);
        } else {
            // 检查幻灯片母版级别的背景
            bgPr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
            bgRef = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);
            var clrMap = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);

            if (bgPr !== undefined) {
                bgcolor = this._getBgFillFromPr(bgPr, slideMasterContent, slideMasterContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, clrMap, index);
            } else if (bgRef !== undefined) {
                bgcolor = this._getBgFillFromRef(bgRef, slideMasterContent, slideMasterContent, slideMasterContent, warpObj, clrMap);
            }
        }
    }

    return bgcolor || "";
};

    /**
 * 从背景属性节点获取背景填充
 * @private
 */
PPTXBackgroundUtils._getBgFillFromPr = function(bgPr, currentContent, slideLayoutContent, slideMasterContent, warpObj, slideContent, slideLayoutContentRef, slideMasterContentRef, clrMapOvr, index) {
    var bgFillTyp = PPTXColorUtils.getFillType(bgPr);
    var bgcolor = "";

    if (bgFillTyp == "SOLID_FILL") {
        var sldFill = bgPr["a:solidFill"];
        if (clrMapOvr === undefined) {
            clrMapOvr = this._getClrMapOverride(currentContent, slideLayoutContentRef, slideMasterContentRef);
        }
        var sldBgClr = PPTXColorUtils.getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
        bgcolor = "background: #" + sldBgClr + ";";
    } else if (bgFillTyp == "GRADIENT_FILL") {
        bgcolor = this.getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
    } else if (bgFillTyp == "PIC_FILL") {
        var source = currentContent === slideContent ? "slideBg" : (currentContent === slideLayoutContentRef ? "slideLayoutBg" : "slideMasterBg");
        bgcolor = this.getBgPicFill(bgPr, source, warpObj, undefined, index);
    }

    return bgcolor;
};

    /**
 * 从背景引用节点获取背景填充
 * @private
 */
PPTXBackgroundUtils._getBgFillFromRef = function(bgRef, currentContent, slideLayoutContent, slideMasterContent, warpObj, clrMapOvr) {
    if (clrMapOvr === undefined) {
        clrMapOvr = this._getClrMapOverride(currentContent, slideLayoutContent, slideMasterContent);
    }

    var phClr = PPTXColorUtils.getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
    var idx = Number(bgRef["attrs"]["idx"]);
    var bgcolor = "";

    if (idx == 0 || idx == 1000) {
        // 无背景
    } else if (idx > 0 && idx < 1000) {
        // fillStyleLst in themeContent - 暂不实现
    } else if (idx > 1000) {
        // bgFillStyleLst in themeContent
        var trueIdx = idx - 1000;
        var bgFillLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:bgFillStyleLst"];
        var bgFillLstIdx = this._getBgFillLstIndex(bgFillLst, trueIdx);
        var bgFillTyp = PPTXColorUtils.getFillType(bgFillLstIdx);

        if (bgFillTyp == "SOLID_FILL") {
            var sldFill = bgFillLstIdx["a:solidFill"];
            var sldBgClr = PPTXColorUtils.getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
            bgcolor = "background: #" + sldBgClr + ";";
        } else if (bgFillTyp == "GRADIENT_FILL") {
            bgcolor = this.getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
        } else if (bgFillTyp == "PIC_FILL") {
            bgcolor = this.getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, undefined);
        }
    }

    return bgcolor;
};

    /**
 * 获取颜色映射覆盖
 * @private
 */
PPTXBackgroundUtils._getClrMapOverride = function(currentContent, slideLayoutContent, slideMasterContent) {
    var clrMapOvr = PPTXUtils.getTextByPathList(currentContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
    if (clrMapOvr === undefined && currentContent !== slideLayoutContent) {
        clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
    }
    if (clrMapOvr === undefined) {
        clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
    }
    return clrMapOvr;
};

    /**
 * 获取背景填充列表中指定索引的项
 * @private
 */
PPTXBackgroundUtils._getBgFillLstIndex = function(bgFillLst, trueIdx) {
    var sortblAry = [];
    Object.keys(bgFillLst).forEach(function (key) {
        var bgFillLstTyp = bgFillLst[key];
        if (key != "attrs") {
            if (bgFillLstTyp.constructor === Array) {
                for (var i = 0; i < bgFillLstTyp.length; i++) {
                    var obj = {};
                    obj[key] = bgFillLstTyp[i];
                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                    obj["attrs"] = { "order": bgFillLstTyp[i]["attrs"]["order"] };
                    sortblAry.push(obj);
                }
            } else {
                var obj = {};
                obj[key] = bgFillLstTyp;
                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                obj["attrs"] = { "order": bgFillLstTyp["attrs"]["order"] };
                sortblAry.push(obj);
            }
        }
    });
    var sortByOrder = sortblAry.slice(0);
    sortByOrder.sort(function (a, b) {
        return a.idex - b.idex;
    });
    return sortByOrder[trueIdx - 1];
};

    /**
 * 获取渐变背景填充
 * @param {Object} bgPr - 背景属性节点
 * @param {string} phClr - 占位符颜色
 * @param {Object} slideMasterContent - 幻灯片母版内容
 * @param {Object} warpObj - 包装对象
 * @returns {string} 渐变背景CSS样式字符串
 */
PPTXBackgroundUtils.getBgGradientFill = function(bgPr, phClr, slideMasterContent, warpObj) {
    var bgcolor = "";
    if (bgPr !== undefined) {
        var grdFill = bgPr["a:gradFill"];
        var gsLst = grdFill["a:gsLst"]["a:gs"];
        var color_ary = [];
        var pos_ary = [];

        for (var i = 0; i < gsLst.length; i++) {
            var lo_color = PPTXColorUtils.getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
            var pos = PPTXUtils.getTextByPathList(gsLst[i], ["attrs", "pos"]);
            if (pos !== undefined) {
                pos_ary[i] = pos / 1000 + "%";
            } else {
                pos_ary[i] = "";
            }
            color_ary[i] = "#" + lo_color;
        }

        // 获取旋转角度
        var lin = grdFill["a:lin"];
        var rot = 90;
        if (lin !== undefined) {
            rot = PPTXUtils.angleToDegrees(lin["attrs"]["ang"]);
            rot = rot + 90;
        }

        bgcolor = "background: linear-gradient(" + rot + "deg,";
        for (var i = 0; i < gsLst.length; i++) {
            if (i == gsLst.length - 1) {
                bgcolor += color_ary[i] + " " + pos_ary[i] + ");";
            } else {
                bgcolor += color_ary[i] + " " + pos_ary[i] + ", ";
            }
        }
    } else {
        if (phClr !== undefined) {
            bgcolor = "background: #" + phClr + ";";
        }
    }
    return bgcolor;
};

    /**
 * 获取图片背景填充
 * @param {Object} bgPr - 背景属性节点
 * @param {string} source - 来源标识
 * @param {Object} warpObj - 包装对象
 * @param {string} phClr - 占位符颜色
 * @param {number} index - 幻灯片索引
 * @returns {string} 图片背景CSS样式字符串
 */
PPTXBackgroundUtils.getBgPicFill = function(bgPr, source, warpObj, phClr, index) {
    var picFillResult = PPTXColorUtils.getPicFill(source, bgPr["a:blipFill"], warpObj);
    // 提取图片 URL（picFillResult 可能是对象或字符串）
    var picFillBase64 = typeof picFillResult === 'object' && picFillResult.img ? picFillResult.img : picFillResult;
    var ordr = bgPr["attrs"]["order"];
    var aBlipNode = bgPr["a:blipFill"]["a:blip"];

    // 处理双色调效果
    var duotone = PPTXUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
    // duotone效果暂未实现

    // 处理透明度
    var aphaModFixNode = PPTXUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"]);
    var imgOpacity = "";
    if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
        var amt = parseInt(aphaModFixNode["amt"]) / 100000;
        imgOpacity = "opacity:" + amt + ";";
    }

    // 处理平铺
    var tileNode = PPTXUtils.getTextByPathList(bgPr, ["a:blipFill", "a:tile", "attrs"]);
    var prop_style = "";
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        prop_style += "background-repeat: round;";
    }

    // 处理拉伸
    var stretch = PPTXUtils.getTextByPathList(bgPr, ["a:blipFill", "a:stretch"]);
    if (stretch !== undefined) {
        var fillRect = PPTXUtils.getTextByPathList(stretch, ["a:fillRect", "attrs"]);
        prop_style += "background-repeat: no-repeat;";
        prop_style += "background-position: center;";
        if (fillRect !== undefined) {
            prop_style += "background-size: 100% 100%;";
        }
    }

    var bgcolor = "background: url(" + picFillBase64 + "); z-index: " + ordr + ";" + prop_style + imgOpacity;
    return bgcolor;
};

    // Export to global scope
// window.PPTXBackgroundUtils = PPTXBackgroundUtils; // Removed for ES modules


export { PPTXBackgroundUtils };

// Also export to global scope for backward compatibility
// window.PPTXBackgroundUtils = PPTXBackgroundUtils; // Removed for ES modules