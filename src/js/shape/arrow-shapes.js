/**
 * 箭头形状渲染模块
 * 提供各种箭头形状的生成和渲染功能
 */

import { PPTXXmlUtils } from '../utils/xml.js';
import { shapeArc } from './path-generators.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * 检查形状是否为箭头形状
 */
export function isArrow(shapType) {
    const arrowShapes = [
        'bentArrow', 'uturnArrow', 'leftArrow', 'rightArrow', 'upArrow', 'downArrow',
        'leftRightArrow', 'upDownArrow', 'quadArrow', 'leftRightUpArrow', 'leftUpArrow',
        'bentUpArrow', 'stripedRightArrow', 'notchedRightArrow', 'curvedDownArrow',
        'curvedLeftArrow', 'curvedRightArrow', 'curvedUpArrow'
    ];
    return arrowShapes.includes(shapType);
}

/**
 * 渲染箭头形状
 */
export function renderArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let result = "";

    // 基础箭头形状（rightArrow, leftArrow, upArrow, downArrow）
    if (shapType === "rightArrow" || shapType === "leftArrow" || 
        shapType === "upArrow" || shapType === "downArrow") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, sAdj1_val = 0.25;
        var sAdj2, sAdj2_val = 0.5;
        var max_sAdj2_const;
        
        if (shapType === "rightArrow" || shapType === "leftArrow") {
            max_sAdj2_const = w / h;
        } else {
            max_sAdj2_const = h / w;
        }
        
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = 1 - ((sAdj2_val2) / max_sAdj2_const);
                }
            }
        }

        var points;
        if (shapType === "rightArrow") {
            points = `${w} ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${sAdj1_val * h},0 ${sAdj1_val * h},0 ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${(1 - sAdj1_val) * h}, ${sAdj2_val * w} ${h}`;
        } else if (shapType === "leftArrow") {
            points = `0 ${h / 2},${(1 - sAdj2_val) * w} 0,${(1 - sAdj2_val) * w} ${sAdj1_val * h},${w} ${sAdj1_val * h},${w} ${(1 - sAdj1_val) * h},${(1 - sAdj2_val) * w} ${(1 - sAdj1_val) * h}, ${(1 - sAdj2_val) * w} ${h}`;
        } else if (shapType === "upArrow") {
            points = `${w / 2} 0,${w} ${(1 - sAdj2_val) * h},${(1 - sAdj1_val) * w} ${(1 - sAdj2_val) * h},${(1 - sAdj1_val) * w} ${h},${sAdj1_val * w} ${h},${sAdj1_val * w} ${(1 - sAdj2_val) * h},0 ${(1 - sAdj2_val) * h}`;
        } else { // downArrow
            points = `${w / 2} ${h},${w} ${sAdj2_val * h},${(1 - sAdj1_val) * w} ${sAdj2_val * h},${(1 - sAdj1_val) * w} 0,${sAdj1_val * w} 0,${sAdj1_val * w} ${sAdj2_val * h},0 ${sAdj2_val * h}`;
        }
        
        result += ` <polygon points='${points}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
    }
    // 双向箭头
    else if (shapType === "leftRightArrow" || shapType === "upDownArrow") {
        var shapAdjst_ary = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, sAdj1_val = 0.25;
        var sAdj2, sAdj2_val = 0.5;
        var max_sAdj2_const = (shapType === "leftRightArrow") ? w / h : h / w;
        
        if (shapAdjst_ary !== undefined) {
            for (var i = 0; i < shapAdjst_ary.length; i++) {
                var sAdj_name = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name === "adj1") {
                    sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name === "adj2") {
                    sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    var sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = 1 - ((sAdj2_val2) / max_sAdj2_const);
                }
            }
        }

        var points;
        if (shapType === "leftRightArrow") {
            points = `0 ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${sAdj1_val * h},${(1 - sAdj2_val) * w} ${sAdj1_val * h},${(1 - sAdj2_val) * w} 0,${w} ${h / 2},${(1 - sAdj2_val) * w} ${h},${(1 - sAdj2_val) * w} ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${h}`;
        } else {
            points = `${w / 2} 0,${w} ${sAdj2_val * h},${(1 - sAdj1_val) * w} ${sAdj2_val * h},${(1 - sAdj1_val) * w} ${(1 - sAdj2_val) * h},${w} ${(1 - sAdj2_val) * h},${w / 2} ${h},0 ${(1 - sAdj2_val) * h},${sAdj1_val * w} ${(1 - sAdj2_val) * h},${sAdj1_val * w} ${sAdj2_val * h},0 ${sAdj2_val * h}`;
        }
        
        result += ` <polygon points='${points}' fill='${(!imgFillFlg ? (grndFillFlg ? "url(#linGrd_" + shpId + ")" : fillColor) : "url(#imgPtrn_" + shpId + ")")}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
    }
    // TODO: 添加其他复杂箭头形状（bentArrow, quadArrow, curvedArrow等）

    return result;
}
