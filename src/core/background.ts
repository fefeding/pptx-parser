/**
 * PPTX 背景工具函数
 * 用于处理幻灯片背景的生成和渲染
 */

import { PPTXUtils } from './utils';
import { PPTXColorUtils } from './color';

interface WarpObj {
    slideContent?: any;
    slideLayoutContent?: any;
    slideMasterContent?: any;
    themeContent?: any;
}

interface SlideSize {
    width: number;
    height: number;
}

/**
 * 获取幻灯片背景
 * @param warpObj - 包装对象,包含幻灯片内容
 * @param slideSize - 幻灯片尺寸
 * @param index - 幻灯片索引
 * @param processNodesInSlide - 处理幻灯片中节点的回调函数
 * @returns 背景HTML字符串
 */
export function getBackground(
    warpObj: WarpObj,
    slideSize: SlideSize,
    index: number,
    processNodesInSlide: (nodeKey: string, node: any, nodes: any, warpObj: WarpObj, type: string) => string
): string {
    const slideContent = warpObj.slideContent;
    const slideLayoutContent = warpObj.slideLayoutContent;
    const slideMasterContent = warpObj.slideMasterContent;

    const nodesSldLayout = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
    const nodesSldMaster = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);

    const showMasterSp = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
    const bgColor = getSlideBackgroundFill(warpObj, index);
    let result = `<div class='slide-background-${index}' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;
    const node_ph_type_ary: string[] = [];

    if (nodesSldLayout !== undefined) {
        for (const nodeKey in nodesSldLayout) {
            if (Array.isArray(nodesSldLayout[nodeKey])) {
                for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
                    const ph_type = PPTXUtils.getTextByPathList(nodesSldLayout[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    if (ph_type != "pic") {
                        result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], nodesSldLayout, warpObj, "slideLayoutBg");
                    }
                }
            } else {
                const ph_type = PPTXUtils.getTextByPathList(nodesSldLayout[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                if (ph_type != "pic") {
                    result += processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, warpObj, "slideLayoutBg");
                }
            }
        }
    }

    if (nodesSldMaster !== undefined && (showMasterSp == "1" || showMasterSp === undefined)) {
        for (const nodeKey in nodesSldMaster) {
            if (Array.isArray(nodesSldMaster[nodeKey])) {
                for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
                    const ph_type = PPTXUtils.getTextByPathList(nodesSldMaster[nodeKey][i], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                    result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], nodesSldMaster, warpObj, "slideMasterBg");
                }
            } else {
                const ph_type = PPTXUtils.getTextByPathList(nodesSldMaster[nodeKey], ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
                result += processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, warpObj, "slideMasterBg");
            }
        }
    }

    return result + "</div>";
}

/**
 * 获取幻灯片背景填充样式
 * @param warpObj - 包装对象
 * @param index - 幻灯片索引
 * @returns 背景CSS样式字符串
 */
export function getSlideBackgroundFill(warpObj: WarpObj, index: number): string {
    const slideContent = warpObj.slideContent;
    const slideLayoutContent = warpObj.slideLayoutContent;
    const slideMasterContent = warpObj.slideMasterContent;

    let bgPr = PPTXUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr"]);
    let bgRef = PPTXUtils.getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgRef"]);
    let bgcolor: string | undefined;

    // 检查幻灯片级别的背景
    if (bgPr !== undefined) {
        bgcolor = _getBgFillFromPr(bgPr, slideContent, slideLayoutContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, undefined, index);
    } else if (bgRef !== undefined) {
        bgcolor = _getBgFillFromRef(bgRef, slideContent, slideLayoutContent, slideMasterContent, warpObj);
    } else {
        // 检查幻灯片布局级别的背景
        bgPr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr"]);
        bgRef = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgRef"]);

        if (bgPr !== undefined) {
            let clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (clrMapOvr === undefined) {
                clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
            bgcolor = _getBgFillFromPr(bgPr, slideLayoutContent, slideLayoutContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, clrMapOvr, index);
        } else if (bgRef !== undefined) {
            let clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (clrMapOvr === undefined) {
                clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
            }
            bgcolor = _getBgFillFromRef(bgRef, slideLayoutContent, slideLayoutContent, slideMasterContent, warpObj, clrMapOvr);
        } else {
            // 检查幻灯片母版级别的背景
            bgPr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr"]);
            bgRef = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgRef"]);
            const clrMap = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);

            if (bgPr !== undefined) {
                bgcolor = _getBgFillFromPr(bgPr, slideMasterContent, slideMasterContent, slideMasterContent, warpObj, slideContent, slideLayoutContent, slideMasterContent, clrMap, index);
            } else if (bgRef !== undefined) {
                bgcolor = _getBgFillFromRef(bgRef, slideMasterContent, slideMasterContent, slideMasterContent, warpObj, clrMap);
            }
        }
    }

    return bgcolor || "";
}

/**
 * 从背景属性节点获取背景填充
 * @private
 */
function _getBgFillFromPr(
    bgPr: any,
    currentContent: any,
    slideLayoutContent: any,
    slideMasterContent: any,
    warpObj: WarpObj,
    slideContent: any,
    slideLayoutContentRef: any,
    slideMasterContentRef: any,
    clrMapOvr: any,
    index: number
): string {
    const bgFillTyp = PPTXColorUtils.getFillType(bgPr);
    let bgcolor = "";

    if (bgFillTyp == "SOLID_FILL") {
        const sldFill = bgPr["a:solidFill"];
        if (clrMapOvr === undefined) {
            clrMapOvr = _getClrMapOverride(currentContent, slideLayoutContentRef, slideMasterContentRef);
        }
        const sldBgClr = PPTXColorUtils.getSolidFill(sldFill, clrMapOvr, undefined, warpObj);
        bgcolor = "background: #" + sldBgClr + ";";
    } else if (bgFillTyp == "GRADIENT_FILL") {
        bgcolor = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj);
    } else if (bgFillTyp == "PIC_FILL") {
        const source = currentContent === slideContent ? "slideBg" : (currentContent === slideLayoutContentRef ? "slideLayoutBg" : "slideMasterBg");
        bgcolor = getBgPicFill(bgPr, source, warpObj, undefined, index);
    }

    return bgcolor;
}

/**
 * 从背景引用节点获取背景填充
 * @private
 */
function _getBgFillFromRef(
    bgRef: any,
    currentContent: any,
    slideLayoutContent: any,
    slideMasterContent: any,
    warpObj: WarpObj,
    clrMapOvr?: any
): string {
    if (clrMapOvr === undefined) {
        clrMapOvr = _getClrMapOverride(currentContent, slideLayoutContent, slideMasterContent);
    }

    const phClr = PPTXColorUtils.getSolidFill(bgRef, clrMapOvr, undefined, warpObj);
    const idx = Number(bgRef["attrs"]["idx"]);
    let bgcolor = "";

    if (idx == 0 || idx == 1000) {
        // 无背景
    } else if (idx > 0 && idx < 1000) {
        // fillStyleLst in themeContent - 暂不实现
    } else if (idx > 1000) {
        // bgFillStyleLst in themeContent
        const trueIdx = idx - 1000;
        const bgFillLst = warpObj.themeContent?.["a:theme"]?.["a:themeElements"]?.["a:fmtScheme"]?.["a:bgFillStyleLst"];
        const bgFillLstIdx = _getBgFillLstIndex(bgFillLst, trueIdx);
        const bgFillTyp = PPTXColorUtils.getFillType(bgFillLstIdx);

        if (bgFillTyp == "SOLID_FILL") {
            const sldFill = bgFillLstIdx["a:solidFill"];
            const sldBgClr = PPTXColorUtils.getSolidFill(sldFill, clrMapOvr, phClr, warpObj);
            bgcolor = "background: #" + sldBgClr + ";";
        } else if (bgFillTyp == "GRADIENT_FILL") {
            bgcolor = getBgGradientFill(bgFillLstIdx, phClr, slideMasterContent, warpObj);
        } else if (bgFillTyp == "PIC_FILL") {
            bgcolor = getBgPicFill(bgFillLstIdx, "themeBg", warpObj, phClr, undefined);
        }
    }

    return bgcolor;
}

/**
 * 获取颜色映射覆盖
 * @private
 */
function _getClrMapOverride(currentContent: any, slideLayoutContent: any, slideMasterContent: any): any {
    let clrMapOvr = PPTXUtils.getTextByPathList(currentContent, ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
    if (clrMapOvr === undefined && currentContent !== slideLayoutContent) {
        clrMapOvr = PPTXUtils.getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
    }
    if (clrMapOvr === undefined) {
        clrMapOvr = PPTXUtils.getTextByPathList(slideMasterContent, ["p:sldMaster", "p:clrMap", "attrs"]);
    }
    return clrMapOvr;
}

/**
 * 获取背景填充列表中指定索引的项
 * @private
 */
function _getBgFillLstIndex(bgFillLst: any, trueIdx: number): any {
    const sortblAry: any[] = [];
    Object.keys(bgFillLst).forEach(function (key) {
        const bgFillLstTyp = bgFillLst[key];
        if (key != "attrs") {
            if (Array.isArray(bgFillLstTyp)) {
                for (let i = 0; i < bgFillLstTyp.length; i++) {
                    const obj: any = {};
                    obj[key] = bgFillLstTyp[i];
                    obj["idex"] = bgFillLstTyp[i]["attrs"]["order"];
                    obj["attrs"] = { "order": bgFillLstTyp[i]["attrs"]["order"] };
                    sortblAry.push(obj);
                }
            } else {
                const obj: any = {};
                obj[key] = bgFillLstTyp;
                obj["idex"] = bgFillLstTyp["attrs"]["order"];
                obj["attrs"] = { "order": bgFillLstTyp["attrs"]["order"] };
                sortblAry.push(obj);
            }
        }
    });
    const sortByOrder = sortblAry.slice(0);
    sortByOrder.sort(function (a, b) {
        return a.idex - b.idex;
    });
    return sortByOrder[trueIdx - 1];
}

/**
 * 获取渐变背景填充
 * @param bgPr - 背景属性节点
 * @param phClr - 占位符颜色
 * @param slideMasterContent - 幻灯片母版内容
 * @param warpObj - 包装对象
 * @returns 渐变背景CSS样式字符串
 */
export function getBgGradientFill(bgPr: any, phClr: string | undefined, slideMasterContent: any, warpObj: WarpObj): string {
    let bgcolor = "";
    if (bgPr !== undefined) {
        const grdFill = bgPr["a:gradFill"];
        const gsLst = grdFill["a:gsLst"]["a:gs"];
        const color_ary: string[] = [];
        const pos_ary: string[] = [];

        for (let i = 0; i < gsLst.length; i++) {
            const lo_color = PPTXColorUtils.getSolidFill(gsLst[i], slideMasterContent["p:sldMaster"]["p:clrMap"]["attrs"], phClr, warpObj);
            const pos = PPTXUtils.getTextByPathList(gsLst[i], ["attrs", "pos"]);
            if (pos !== undefined) {
                pos_ary[i] = pos / 1000 + "%";
            } else {
                pos_ary[i] = "";
            }
            color_ary[i] = "#" + lo_color;
        }

        // 获取旋转角度
        const lin = grdFill["a:lin"];
        let rot = 90;
        if (lin !== undefined) {
            rot = PPTXUtils.angleToDegrees(lin["attrs"]["ang"]);
            rot = rot + 90;
        }

        bgcolor = "background: linear-gradient(" + rot + "deg,";
        for (let i = 0; i < gsLst.length; i++) {
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
}

/**
 * 获取图片背景填充
 * @param bgPr - 背景属性节点
 * @param source - 来源标识
 * @param warpObj - 包装对象
 * @param phClr - 占位符颜色
 * @param index - 幻灯片索引
 * @returns 图片背景CSS样式字符串
 */
export function getBgPicFill(bgPr: any, source: string, warpObj: WarpObj, phClr?: string, index?: number): string {
    const picFillResult = PPTXColorUtils.getPicFill(source, bgPr["a:blipFill"], warpObj);
    // 提取图片 URL（picFillResult 可能是对象或字符串）
    const picFillBase64 = typeof picFillResult === 'object' && picFillResult.img ? picFillResult.img : picFillResult;
    const ordr = bgPr["attrs"]["order"];
    const aBlipNode = bgPr["a:blipFill"]["a:blip"];

    // 处理双色调效果
    const duotone = PPTXUtils.getTextByPathList(aBlipNode, ["a:duotone"]);
    // duotone效果暂未实现

    // 处理透明度
    const aphaModFixNode = PPTXUtils.getTextByPathList(aBlipNode, ["a:alphaModFix", "attrs"]);
    let imgOpacity = "";
    if (aphaModFixNode !== undefined && aphaModFixNode["amt"] !== undefined && aphaModFixNode["amt"] != "") {
        const amt = parseInt(aphaModFixNode["amt"]) / 100000;
        imgOpacity = "opacity:" + amt + ";";
    }

    // 处理平铺
    const tileNode = PPTXUtils.getTextByPathList(bgPr, ["a:blipFill", "a:tile", "attrs"]);
    let prop_style = "";
    if (tileNode !== undefined && tileNode["sx"] !== undefined) {
        prop_style += "background-repeat: round;";
    }

    // 处理拉伸
    const stretch = PPTXUtils.getTextByPathList(bgPr, ["a:blipFill", "a:stretch"]);
    if (stretch !== undefined) {
        const fillRect = PPTXUtils.getTextByPathList(stretch, ["a:fillRect", "attrs"]);
        prop_style += "background-repeat: no-repeat;";
        prop_style += "background-position: center;";
        if (fillRect !== undefined) {
            prop_style += "background-size: 100% 100%;";
        }
    }

    const bgcolor = "background: url(" + picFillBase64 + "); z-index: " + ordr + ";" + prop_style + imgOpacity;
    return bgcolor;
}

// 向后兼容的导出对象
const PPTXBackgroundUtils = {
    getBackground,
    getSlideBackgroundFill,
    getBgGradientFill,
    getBgPicFill,
    // 私有方法（不导出）
    _getBgFillFromPr,
    _getBgFillFromRef,
    _getClrMapOverride,
    _getBgFillLstIndex
};

export { PPTXBackgroundUtils };