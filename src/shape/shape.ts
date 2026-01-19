/**
 * PPTX 形状工具模块
 * 提供绘制各种形状的函数
 */

import { PPTXUtils } from '../core/utils.js';

/**
 * 绘制饼图形状
 */
function shapePie(H: number, w: number, adj1: number, adj2: number, isClose: boolean): [string, string] {
    const pieVal: number = parseInt(adj2.toString());
    const piAngle: number = parseInt(adj1.toString());
    const size: number = parseInt(H.toString());
    const radius: number = size / 2;
    let value: number = pieVal - piAngle;

    if (value < 0) {
        value = 360 + value;
    }

    value = Math.min(Math.max(value, 0), 360);

    // 计算 x,y 坐标
    const x: number = Math.cos((2 * Math.PI) / (360 / value));
    const y: number = Math.sin((2 * Math.PI) / (360 / value));

    let longArc: number, d: string, rot: string;
    if (isClose) {
        longArc = (value <= 180) ? 0 : 1;
        d = `M${radius},${radius} L${radius},0 A${radius},${radius} 0 ${longArc},1 ${radius + y * radius},${radius - x * radius} z`;
        rot = `rotate(${piAngle - 270}, ${radius}, ${radius})`;
    } else {
        longArc = (value <= 180) ? 0 : 1;
        const radius1: number = radius;
        const radius2: number = w / 2;
        d = `M${radius1},0 A${radius2},${radius1} 0 ${longArc},1 ${radius2 + y * radius2},${radius1 - x * radius1}`;
        rot = `rotate(${piAngle + 90}, ${radius}, ${radius})`;
    }

    return [d, rot];
}

/**
 * 绘制齿轮形状
 */
function shapeGear(w: number, h: number, points: number): string {
    const innerRadius: number = h;
    const outerRadius: number = 1.5 * innerRadius;
    const cx: number = outerRadius;
    const cy: number = outerRadius;
    const notches: number = points;
    const radiusO: number = outerRadius;
    const radiusI: number = innerRadius;
    const taperO: number = 50;
    const taperI: number = 35;

    const pi2: number = 2 * Math.PI;
    const angle: number = pi2 / (notches * 2);
    const taperAI: number = angle * taperI * 0.005;
    const taperAO: number = angle * taperO * 0.005;
    let a: number = angle;
    let toggle: boolean = false;

    let d: string = ` M${cx + radiusO * Math.cos(taperAO)} ${cy + radiusO * Math.sin(taperAO)}`;

    for (; a <= pi2 + angle; a += angle) {
        if (toggle) {
            d += ` L${cx + radiusI * Math.cos(a - taperAI)},${cy + radiusI * Math.sin(a - taperAI)}`;
            d += ` L${cx + radiusO * Math.cos(a + taperAO)},${cy + radiusO * Math.sin(a + taperAO)}`;
        } else {
            d += ` L${cx + radiusO * Math.cos(a - taperAO)},${cy + radiusO * Math.sin(a - taperAO)}`;
            d += ` L${cx + radiusI * Math.cos(a + taperAI)},${cy + radiusI * Math.sin(a + taperAI)}`;
        }
        toggle = !toggle;
    }
    d += " ";
    return d;
}

/**
 * 绘制弧形形状
 */
function shapeArc(cX: number, cY: number, rX: number, rY: number, stAng: number, endAng: number, isClose: boolean): string {
    let dData: string;
    let angle: number = stAng;
    let x: number, y: number;
    if (endAng >= stAng) {
        while (angle <= endAng) {
            let radians: number = angle * (Math.PI / 180);
            x = cX + Math.cos(radians) * rX;
            y = cY + Math.sin(radians) * rY;
            if (angle === stAng) {
                dData = ` M${x} ${y}`;
            }
            dData += ` L${x} ${y}`;
            angle++;
        }
    } else {
        while (angle > endAng) {
            let radians: number = angle * (Math.PI / 180);
            x = cX + Math.cos(radians) * rX;
            y = cY + Math.sin(radians) * rY;
            if (angle === stAng) {
                dData = ` M ${x} ${y}`;
            }
            dData += ` L ${x} ${y}`;
            angle--;
        }
    }
    dData += (isClose ? " z" : "");
    return dData;
}

/**
 * 绘制圆角或切角矩形
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} adj1 - 调整参数1
 * @param {number} adj2 - 调整参数2
 * @param {string} shapeType - 形状类型: "snip" 或 "round"
 * @param {string} adjType - 调整类型: "cornr1", "cornr2", "cornrAll", "diag"
 * @returns {string} SVG 路径字符串
 */
function shapeSnipRoundRect(w: number, h: number, adj1: number, adj2: number, shapeType: string, adjType: string): string {
    let adjA: number, adjB: number, adjC: number, adjD: number;
    if (adjType === "cornr1") {
        adjA = 0;
        adjB = 0;
        adjC = 0;
        adjD = adj1;
    } else if (adjType === "cornr2") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj2;
        adjD = adj1;
    } else if (adjType === "cornrAll") {
        adjA = adj1;
        adjB = adj1;
        adjC = adj1;
        adjD = adj1;
    } else if (adjType === "diag") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj1;
        adjD = adj2;
    }

    let d: string;
    if (shapeType === "round") {
        d = `M0,${h / 2 + (1 - adjB) * (h / 2)} Q0,${h} ${adjB * (w / 2)},${h} L${w / 2 + (1 - adjC) * (w / 2)},${h} Q${w},${h} ${w},${h / 2 + (h / 2) * (1 - adjC)}L${w},${h / 2 * adjD} Q${w},0 ${w / 2 + (w / 2) * (1 - adjD)},0 L${w / 2 * adjA},0 Q0,0 0,${h / 2 * adjA} z`;
    } else if (shapeType === "snip") {
        d = `M0,${adjA * (h / 2)} L0,${h / 2 + (h / 2) * (1 - adjB)}L${adjB * (w / 2)},${h} L${w / 2 + (w / 2) * (1 - adjC)},${h}L${w},${h / 2 + (h / 2) * (1 - adjC)} L${w},${adjD * (h / 2)}L${w / 2 + (w / 2) * (1 - adjD)},0 L${(w / 2) * adjA},0 z`;
    }
    return d;
}

/**
 * PPTX 形状工具模块
 */
const PPTXShapeUtils = {
    shapePie,
    shapeGear,
    shapeArc,
    shapeSnipRoundRect
};

export { PPTXShapeUtils };