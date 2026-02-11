/**
 * 箭头形状渲染模块
 * 处理各种箭头形状的 SVG 生成
 */

import { PPTXXmlUtils } from '../utils/xml.js';

/**
 * 判断形状是否为箭头
 */
export function isArrow(shapType) {
    return [
        "rightArrow", "leftArrow", "upArrow", "downArrow",
        "leftRightArrow", "upDownArrow",
        "quadArrow", "leftRightUpArrow", "leftUpArrow",
        "bentUpArrow", "bentArrow", "uturnArrow",
        "stripedRightArrow", "notchedRightArrow",
        "homePlate", "chevron",
        "curvedDownArrow", "curvedLeftArrow", "curvedRightArrow", "curvedUpArrow",
        "swooshArrow", "circularArrow", "leftCircularArrow",
        "rightArrowCallout", "downArrowCallout", "leftArrowCallout",
        "upArrowCallout", "leftRightArrowCallout", "quadArrowCallout"
    ].includes(shapType);
}

/**
 * 渲染箭头形状
 */
export function renderArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    // 基础箭头形状（rightArrow, leftArrow, upArrow, downArrow）
    if (shapType === "rightArrow" || shapType === "leftArrow" ||
        shapType === "upArrow" || shapType === "downArrow") {
        return renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    // 双向箭头
    if (shapType === "leftRightArrow" || shapType === "upDownArrow") {
        return renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    
    // 复杂箭头形状暂未实现，返回空字符串由 shape.js 处理
    return "";
}

/**
 * 渲染基础箭头形状
 */
function renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    var sAdj1, sAdj1_val = 0.25;
    var sAdj2, sAdj2_val = 0.5;
    var max_sAdj2_const = w / h;

    if (shapAdjst_ary !== undefined) {
        for (var i = 0; i < shapAdjst_ary.length; i++) {
            var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdj_name == "adj2") {
                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
            }
        }
    }

    var points;
    if (shapType === "rightArrow") {
        points = w + " " + h / 2 + "," + sAdj2_val * w + " 0," + sAdj2_val * w + " " + sAdj1_val * h + ",0 " + sAdj1_val * h + ",0 " + (1 - sAdj1_val) * h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + ", " + sAdj2_val * w + " " + h;
    } else if (shapType === "leftArrow") {
        points = "0 " + h / 2 + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + "," + w + " " + (1 - sAdj1_val) * h + "," + w + " " + sAdj1_val * h + "," + sAdj2_val * w + " " + sAdj1_val * h + ", " + sAdj2_val * w + " 0";
    } else if (shapType === "upArrow") {
        var sAdj1_val = 0.25;
        var sAdj2_val = 0.5;
        var max_sAdj2_const = h / w;
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }
        points = (w / 2) + " 0,0 " + sAdj2_val * h + "," + (0.5 - sAdj1_val) * w + " " + sAdj2_val * h + "," + (0.5 - sAdj1_val) * w + " " + h + "," + (0.5 + sAdj1_val) * w + " " + h + "," + (0.5 + sAdj1_val) * w + " " + sAdj2_val * h + ", " + w + " " + sAdj2_val * h;
    } else { // downArrow
        var sAdj1_val = 0.25;
        var sAdj2_val = 0.5;
        var max_sAdj2_const = h / w;
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }
        points = (0.5 - sAdj1_val) * w + " 0," + (0.5 - sAdj1_val) * w + " " + (1 - sAdj2_val) * h + ",0 " + (1 - sAdj2_val) * h + "," + (w / 2) + " " + h + "," + w + " " + (1 - sAdj2_val) * h + "," + (0.5 + sAdj1_val) * w + " " + (1 - sAdj2_val) * h + ", " + (0.5 + sAdj1_val) * w + " 0";
    }

    return " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}

/**
 * 渲染双向箭头
 */
function renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    var sAdj1_val = 0.25;
    var sAdj2_val = 0.5;
    var max_sAdj2_const = w / h;
    var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst_ary !== undefined) {
        for (var i = 0; i < shapAdjst_ary.length; i++) {
            var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
            if (sAdj_name == "adj1") {
                var sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdj_name == "adj2") {
                var sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
            }
        }
    }

    var points;
    if (shapType === "leftRightArrow") {
        points = "0 " + h / 2 + "," + sAdj2_val * w + " 0," + sAdj2_val * w + " " + h + ",0 " + h + "," + w + " " + h / 2 + "," + sAdj2_val * w + " " + w + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " 0";
    } else { // upDownArrow
        var sAdj1_val = 0.25;
        var sAdj2_val = 0.5;
        var max_sAdj2_const = h / w;
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    var sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdj_name == "adj2") {
                    var sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }
        points = w / 2 + " 0," + w + " " + sAdj2_val * h + "," + w + " " + h + ", " + sAdj2_val * w + " " + h + "," + w / 2 + " " + h + ",0 " + sAdj2_val * h + ",0 " + sAdj2_val * h + "," + sAdj1_val * w + " 0, " + sAdj1_val * w + " 0";
    }

    return " <polygon points='" + points + "' fill='" + (!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")") + "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
}
