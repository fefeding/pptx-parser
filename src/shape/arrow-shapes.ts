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

import { PPTXXmlUtils } from '../utils/xml';
import type { XmlNode } from '../core/types';
import type { ShapeBorder } from './pie-shapes';

// ==================== 导出函数 ====================

/**
 * 判断形状是否为箭头
 * @param {string} shapType - 形状类型
 * @returns {boolean} 是否为箭头形状
 */
export function isArrow(shapType: string): boolean {
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
 * @param {ShapeBorder} border - 边框配置 {color, width, strokeDasharray}
 * @param {string} shpId - 形状ID
 * @param {XmlNode} node - 形状节点
 * @returns {string} SVG 字符串
 */
export function renderArrow(shapType: string, w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, fillColor: string, border: ShapeBorder, shpId: string, node: XmlNode): string {
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
 * @param {XmlNode} node - 形状节点
 * @returns {Object} 包含 adj1, adj2 值的对象
 */
/**
 * 读取 a:avLst 中的 adj1/adj2 原始值（100000 = 100%），未提供时回退到预设默认值 50000
 * —— 与 PowerPoint/WPS 一致（非法/缺失的 gd 会被忽略并用默认值）。
 *
 * 这里保持 OOXML 的原始语义，不做提前换算：
 *   rightArrow / leftArrow：dx1 = ss*a2/100000（箭头长度），dy1 = h*a1/200000（半箭身厚度）
 *   upArrow / downArrow   ：dy1 = ss*a2/100000，dx1 = w*a1/200000
 */
function readArrowAdj(node: XmlNode): { adj1: number; adj2: number } {
    let adj1 = 50000;
    let adj2 = 50000;
    const gdList = PPTXXmlUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
    // 单个 gd 时解析结果是对象而非数组，这里统一成数组处理
    const items: any[] = Array.isArray(gdList) ? gdList : (gdList ? [gdList] : []);
    for (const item of items) {
        const name = PPTXXmlUtils.getTextByPathList(item, ["attrs", "name"]);
        const fmla = String(PPTXXmlUtils.getTextByPathList(item, ["attrs", "fmla"]) || "");
        const val = parseInt(fmla.replace(/^\s*val\s+/, ""));
        if (isNaN(val)) continue;
        if (name === "adj1") {
            adj1 = val;
        } else if (name === "adj2") {
            adj2 = val;
        }
    }
    return { adj1, adj2 };
}

/** 收敛到 [min, max] */
function clampValue(v: number, min: number, max: number): number {
    return Math.min(max, Math.max(min, v));
}

/**
 * 渲染基础箭头形状
 */
function renderBasicArrow(shapType: string, w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, fillColor: string, border: ShapeBorder, shpId: string, node: XmlNode): string {
    const { adj1, adj2 } = readArrowAdj(node);
    const ss = Math.min(w, h);
    const a1 = clampValue(adj1, 0, 100000);
    let points: string;

    if (shapType === "rightArrow" || shapType === "leftArrow") {
        // a2 上限 = 100000*w/ss（预设公式 maxAdj2），箭头长度 dx1 = ss*a2/100000，
        // 半箭身厚度 dy1 = h*a1/200000；箭头肩部位于距箭尖 dx1 处。
        const a2 = clampValue(adj2, 0, 100000 * w / ss);
        const dx1 = ss * a2 / 100000;
        const dy1 = h * a1 / 200000;
        const y1 = h / 2 - dy1;
        const y2 = h / 2 + dy1;
        if (shapType === "rightArrow") {
            const x1 = w - dx1;
            points = `0 ${y1},${x1} ${y1},${x1} 0,${w} ${h / 2},${x1} ${h},${x1} ${y2},0 ${y2}`;
        } else {
            const x1 = dx1;
            points = `${w} ${y1},${x1} ${y1},${x1} 0,0 ${h / 2},${x1} ${h},${x1} ${y2},${w} ${y2}`;
        }
    } else {
        // upArrow / downArrow：主轴为高度，箭头长度 dy1 = ss*a2/100000，半箭身厚度 dx1 = w*a1/200000
        const a2 = clampValue(adj2, 0, 100000 * h / ss);
        const dy1 = ss * a2 / 100000;
        const dx1 = w * a1 / 200000;
        const x1 = w / 2 - dx1;
        const x2 = w / 2 + dx1;
        if (shapType === "upArrow") {
            const y1 = dy1;
            points = `${x1} ${h},${x1} ${y1},0 ${y1},${w / 2} 0,${w} ${y1},${x2} ${y1},${x2} ${h}`;
        } else {
            const y1 = h - dy1;
            points = `${x1} 0,${x1} ${y1},0 ${y1},${w / 2} ${h},${w} ${y1},${x2} ${y1},${x2} 0`;
        }
    }

    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 渲染双向箭头
 */
function renderDoubleArrow(shapType: string, w: number, h: number, imgFillFlg: boolean, grndFillFlg: boolean, fillColor: string, border: ShapeBorder, shpId: string, node: XmlNode): string {
    const { adj1, adj2 } = readArrowAdj(node);
    const ss = Math.min(w, h);
    const a1 = clampValue(adj1, 0, 100000);
    let points: string;

    if (shapType === "leftRightArrow") {
        // 左右各一个箭头：箭头长度 dx = ss*a2/100000，半箭身厚度 dy1 = h*a1/200000
        const a2 = clampValue(adj2, 0, 100000 * w / ss);
        const dx = ss * a2 / 100000;
        const x1 = dx;
        const x2 = w - dx;
        const dy1 = h * a1 / 200000;
        const y1 = h / 2 - dy1;
        const y2 = h / 2 + dy1;
        points = `0 ${y1},${x1} ${y1},${x1} 0,0 ${h / 2},${x1} ${h},${x1} ${y2},` +
            `${x2} ${y2},${x2} ${h},${w} ${h / 2},${x2} 0,${x2} ${y1}`;
    } else {
        // upDownArrow：上下各一个箭头
        const a2 = clampValue(adj2, 0, 100000 * h / ss);
        const dy = ss * a2 / 100000;
        const y1 = dy;
        const y2 = h - dy;
        const dx1 = w * a1 / 200000;
        const x1 = w / 2 - dx1;
        const x2 = w / 2 + dx1;
        points = `${x1} ${h},${x1} ${y2},0 ${y2},${w / 2} ${h},${w} ${y2},${x2} ${y2},` +
            `${x2} ${y1},${w} ${y1},${w / 2} 0,0 ${y1},${x1} ${y1}`;
    }

    return buildPolygon(points, imgFillFlg, grndFillFlg, fillColor, border, shpId);
}

/**
 * 构建 polygon 元素字符串
 * @param {string} points - 点坐标字符串
 * @param {boolean} imgFillFlg - 是否使用图片填充
 * @param {boolean} grndFillFlg - 是否使用渐变填充
 * @param {string} fillColor - 填充颜色
 * @param {ShapeBorder} border - 边框配置
 * @param {string} shpId - 形状ID
 * @returns {string} SVG 字符串
 */
function buildPolygon(points: string, imgFillFlg: boolean, grndFillFlg: boolean, fillColor: string, border: ShapeBorder, shpId: string): string {
    const fillUrl = !imgFillFlg
        ? (grndFillFlg ? `url(#linGrd_${shpId})` : fillColor)
        : `url(#imgPtrn_${shpId})`;

    return ` <polygon points='${points}' fill='${fillUrl}' stroke='${border.color}' stroke-width='${border.width}' stroke-dasharray='${border.strokeDasharray}' />`;
}
