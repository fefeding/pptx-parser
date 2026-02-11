/**
 * 箭头形状渲染模块
 * 处理各种箭头形状的 SVG 生成
 * 
 * 箭头分类:
 * - 基础箭头: rightArrow, leftArrow, upArrow, downArrow
 * - 双向箭头: leftRightArrow, upDownArrow  
 * - 复杂箭头: quadArrow, bentArrow, curvedArrow, circularArrow 等
 * - 标注箭头: xxxArrowCallout
 */

import { PPTXXmlUtils } from '../utils/xml.js';

// ==================== 导出函数 ====================

/**
 * 判断形状是否为箭头
 * @param {string} shapType - 形状类型
 * @returns {boolean} 是否为箭头形状
 */
export function isArrow(shapType) {
    const arrowShapes = [
        // 基础箭头
        "rightArrow", "leftArrow", "upArrow", "downArrow",
        // 双向箭头
        "leftRightArrow", "upDownArrow",
        // 多向箭头
        "quadArrow", "leftRightUpArrow", "leftUpArrow",
        // 弯曲箭头
        "bentUpArrow", "bentArrow", "uturnArrow",
        // 条纹/缺口箭头
        "stripedRightArrow", "notchedRightArrow",
        // 其他箭头
        "homePlate", "chevron",
        // 曲线箭头
        "curvedDownArrow", "curvedLeftArrow", "curvedRightArrow", "curvedUpArrow",
        "swooshArrow", "circularArrow", "leftCircularArrow",
        // 标注箭头
        "rightArrowCallout", "downArrowCallout", "leftArrowCallout",
        "upArrowCallout", "leftRightArrowCallout", "quadArrowCallout"
    ];
    return arrowShapes.includes(shapType);
}

/**
 * 渲染箭头形状
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置 {color, width, strokeDasharray}
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
export function renderArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    // 基础箭头形状（rightArrow, leftArrow, upArrow, downArrow）
    if (["rightArrow", "leftArrow", "upArrow", "downArrow"].includes(shapType)) {
        return renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    
    // 双向箭头
    if (["leftRightArrow", "upDownArrow"].includes(shapType)) {
        return renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node);
    }
    
    // 复杂箭头形状暂未实现，返回空字符串由 shape.js 处理
    return "";
}

// ==================== 内部函数 ====================

/**
 * 读取形状调整参数
 * @param {Object} node - 形状节点
 * @returns {Object} 包含 adj1, adj2 值的对象
 */
function readAdjustmentParams(node, w, h) {
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    let sAdj1, sAdj1_val = 0.25;
    let sAdj2, sAdj2_val = 0.5;
    
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj1") {
                sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdjName === "adj2") {
                sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                const maxConst = w / h;  // 会在调用处重新计算
                sAdj2_val = sAdj2Val2 / maxConst;
            }
        }
    }
    
    return { sAdj1_val, sAdj2_val };
}

/**
 * 渲染基础箭头形状
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
function renderBasicArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let { sAdj1_val, sAdj2_val } = readAdjustmentParams(node, w, h);
    const max_sAdj2_const = w / h;
    
    // 重新读取并计算 sAdj2_val（因为需要正确的 max_sAdj2_const）
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    
    let points;
    if (shapType === "rightArrow") {
        points = `${w} ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${sAdj1_val * h},0 ${sAdj1_val * h},0 ${(1 - sAdj1_val) * h},${sAdj2_val * w} ${(1 - sAdj1_val) * h}, ${sAdj2_val * w} ${h}`;
    } else if (shapType === "leftArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} ${h},${sAdj2_val * w} ${(1 - sAdj1_val) * h},${w} ${(1 - sAdj1_val) * h},${w} ${sAdj1_val * h},${sAdj2_val * w} ${sAdj1_val * h}, ${sAdj2_val * w} 0`;
    } else if (shapType === "upArrow") {
        // upArrow 使用不同的宽高比计算
        const max_sAdj2_const_up = h / w;
        const { sAdj1_val: sAdj1_up, sAdj2_val: sAdj2_up } = readAdjustmentParams(node, w, h);
        const shapAdjst_up = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_up = 0.5;
        
        if (shapAdjst_up) {
            for (let i = 0; i < shapAdjst_up.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_up[i], ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_up[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_up = sAdj2Val2 / max_sAdj2_const_up;
                }
            }
        }
        
        points = `${w / 2} 0,0 ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${sAdj2_val_up * h},${(0.5 - sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${h},${(0.5 + sAdj1_up) * w} ${sAdj2_val_up * h}, ${w} ${sAdj2_val_up * h}`;
    } else { // downArrow
        const max_sAdj2_const_down = h / w;
        const { sAdj1_val: sAdj1_down, sAdj2_val: sAdj2_down } = readAdjustmentParams(node, w, h);
        const shapAdjst_down = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj2_val_down = 0.5;
        
        if (shapAdjst_down) {
            for (let i = 0; i < shapAdjst_down.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_down[i], ["attrs", "name"]);
                if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_down[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val_down = sAdj2Val2 / max_sAdj2_const_down;
                }
            }
        }
        
        points = `${(0.5 - sAdj1_down) * w} 0,${(0.5 - sAdj1_down) * w} ${(1 - sAdj2_val_down) * h},0 ${(1 - sAdj2_val_down) * h},${w / 2} ${h},${w} ${(1 - sAdj2_val_down) * h},${(0.5 + sAdj1_down) * w} ${(1 - sAdj2_val_down) * h}, ${(0.5 + sAdj1_down) * w} 0`;
    }
    
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 渲染双向箭头
 * @param {string} shapType - 箭头类型
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @param {Object} node - 形状节点
 * @returns {string} SVG 字符串
 */
function renderDoubleArrow(shapType, w, h, imgFillFlg, grndFillFlg, fillColor, border, shpId, node) {
    let sAdj1_val = 0.25;
    let sAdj2_val = 0.5;
    const max_sAdj2_const = w / h;
    
    const shapAdjst = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    if (shapAdjst) {
        for (let i = 0; i < shapAdjst.length; i++) {
            const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "name"]);
            if (sAdjName === "adj1") {
                const sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
            } else if (sAdjName === "adj2") {
                const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst[i], ["attrs", "fmla"]);
                const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                sAdj2_val = sAdj2Val2 / max_sAdj2_const;
            }
        }
    }
    
    let points;
    if (shapType === "leftRightArrow") {
        points = `0 ${h / 2},${sAdj2_val * w} 0,${sAdj2_val * w} ${h},0 ${h},${w} ${h / 2},${sAdj2_val * w} ${w},${sAdj2_val * w} ${h},${sAdj2_val * w} 0`;
    } else { // upDownArrow
        // upDownArrow 使用不同的宽高比计算
        sAdj1_val = 0.25;
        sAdj2_val = 0.5;
        const max_sAdj2_const_ud = h / w;
        
        const shapAdjst_ud = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        if (shapAdjst_ud) {
            for (let i = 0; i < shapAdjst_ud.length; i++) {
                const sAdjName = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "name"]);
                if (sAdjName === "adj1") {
                    const sAdj1 = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdjName === "adj2") {
                    const sAdj2 = PPTXXmlUtils.getTextByPathList(shapAdjst_ud[i], ["attrs", "fmla"]);
                    const sAdj2Val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = sAdj2Val2 / max_sAdj2_const_ud;
                }
            }
        }
        
        points = `${w / 2} 0,${w} ${sAdj2_val * h},${w} ${h}, ${sAdj2_val * w} ${h},${w / 2} ${h},0 ${sAdj2_val * h},0 ${sAdj2_val * h},${sAdj1_val * w} 0, ${sAdj1_val * w} 0`;
    }
    
    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 构建 polygon 元素字符串
 * @param {string} points - 点坐标字符串
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {Object} border - 边框配置
 * @param {string} shpId - 形状ID
 * @returns {string} SVG 字符串
 */
function buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId) {
    const fillUrl = !imgFillFlg 
        ? (grndFillFlg ? `url(#linGrd_${shpId})` : fillColor) 
        : `url(#imgPtrn_${shpId})`;
    
    return ` <polygon points='${points}' fill='${fillUrl}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
