/**
 * 路径生成器模块
 * 纯数学计算函数,无外部依赖,无副作用
 * 从 shape.js 提取的独立工具函数
 */

/**
 * polarToCartesian - 将极坐标转换为笛卡尔坐标
 * @param {number} cx - 圆心X坐标
 * @param {number} cy - 圆心Y坐标
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} angleInDegrees - 角度
 * @returns {Object} 笛卡尔坐标对象 {x, y}
 */
export function polarToCartesian(cx, cy, w, h, angleInDegrees) {
    var angleInRadians = (angleInDegrees - 90) * Math.PI / 180.0;
    return {
        x: cx + (w / 2) * Math.cos(angleInRadians),
        y: cy + (h / 2) * Math.sin(angleInRadians)
    };
}

/**
 * shapeArc - 生成圆弧路径
 * @param {number} cx - 圆心X坐标
 * @param {number} cy - 圆心Y坐标
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} startAngle - 起始角度
 * @param {number} endAngle - 结束角度
 * @param {boolean} clockwise - 是否顺时针
 * @returns {string} SVG路径字符串
 */
export function shapeArc(cx, cy, w, h, startAngle, endAngle, clockwise) {
    var start = polarToCartesian(cx, cy, w, h, endAngle);
    var end = polarToCartesian(cx, cy, w, h, startAngle);
    var largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
    var d = [
        "M", start.x, start.y,
        "A", w, h, 0, largeArcFlag, clockwise ? "0" : "1", end.x, end.y
    ].join(" ");
    return d;
}

/**
 * shapeArcAlt - 生成圆弧路径(逐点计算实现)
 * @param {number} cX - 圆心X坐标
 * @param {number} cY - 圆心Y坐标
 * @param {number} rX - X半径
 * @param {number} rY - Y半径
 * @param {number} stAng - 起始角度
 * @param {number} endAng - 结束角度
 * @param {boolean} isClose - 是否闭合
 * @returns {string} SVG路径字符串
 */
export function shapeArcAlt(cX, cY, rX, rY, stAng, endAng, isClose) {
    var dData;
    var angle = stAng;
    if (endAng >= stAng) {
        while (angle <= endAng) {
            var radians = angle * (Math.PI / 180);
            var x = cX + Math.cos(radians) * rX;
            var y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M" + x + " " + y;
            }
            dData += " L" + x + " " + y;
            angle++;
        }
    } else {
        while (angle > endAng) {
            var radians = angle * (Math.PI / 180);
            var x = cX + Math.cos(radians) * rX;
            var y = cY + Math.sin(radians) * rY;
            if (angle == stAng) {
                dData = " M " + x + " " + y;
            }
            dData += " L " + x + " " + y;
            angle--;
        }
    }
    dData += (isClose ? " z" : "");
    return dData;
}

/**
 * shapeSnipRoundRect - 生成圆角或裁剪矩形路径
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} sAdj1_val - 调整值1
 * @param {number} sAdj2_val - 调整值2
 * @param {string} shpTyp - 形状类型 ("round" 或 "snip")
 * @param {string} adjTyp - 调整类型 ("cornr1", "cornr2", "cornrAll", "diag")
 * @returns {string} SVG路径字符串
 */
export function shapeSnipRoundRect(w, h, sAdj1_val, sAdj2_val, shpTyp, adjTyp) {
    var d = "";
    var sAdj1 = 0;
    var sAdj2 = 0;

    if (shpTyp == "round") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornrAll") {
            d = "M0," + sAdj1 + " Q0,0 " + sAdj1 + ",0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj1) + " Q" + w + "," + h + " " + (w - sAdj1) + "," + h + " L" + sAdj1 + "," + h + " Q0," + h + " 0," + (h - sAdj1) + " z";
        } else if (adjTyp == "cornr1") {
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + h + " L0," + h + " z";
        } else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj2) + " Q" + w + "," + h + " " + (w - sAdj2) + "," + h + " L" + sAdj1 + "," + h + " Q0," + h + " 0," + (h - sAdj1) + " L0," + sAdj2 + " Q0,0 " + sAdj2 + ",0 z";
        } else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = "M0,0 L" + (w - sAdj1) + ",0 Q" + w + ",0 " + w + "," + sAdj1 + " L" + w + "," + (h - sAdj2) + " Q" + w + "," + h + " " + (w - sAdj2) + "," + h + " L0," + h + " z";
        }
    } else if (shpTyp == "snip") {
        sAdj1 = w * sAdj1_val;
        if (adjTyp == "cornr1") {
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + h + " L0," + h + " L0," + sAdj1 + " z";
        } else if (adjTyp == "diag") {
            sAdj2 = h * sAdj2_val;
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + (h - sAdj2) + " L" + sAdj2 + "," + h + " L0," + h + " L0," + sAdj1 + " z";
        } else if (adjTyp == "cornr2") {
            sAdj2 = w * sAdj2_val;
            d = "M" + sAdj1 + ",0 L" + w + ",0 L" + w + "," + (h - sAdj2) + " L" + (w - sAdj2) + "," + h + " L0," + h + " z";
        }
    }

    return d;
}

/**
 * shapeSnipRoundRectAlt - 生成圆角或裁剪矩形路径(备选实现)
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} adj1 - 调整值1
 * @param {number} adj2 - 调整值2
 * @param {string} shapeType - 形状类型 ("snip" 或 "round")
 * @param {string} adjType - 调整类型 ("cornr1", "cornr2", "cornrAll", "diag")
 * @returns {string} SVG路径字符串
 */
export function shapeSnipRoundRectAlt(w, h, adj1, adj2, shapeType, adjType) {
    var adjA, adjB, adjC, adjD;
    if (adjType == "cornr1") {
        adjA = 0;
        adjB = 0;
        adjC = 0;
        adjD = adj1;
    } else if (adjType == "cornr2") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj2;
        adjD = adj1;
    } else if (adjType == "cornrAll") {
        adjA = adj1;
        adjB = adj1;
        adjC = adj1;
        adjD = adj1;
    } else if (adjType == "diag") {
        adjA = adj1;
        adjB = adj2;
        adjC = adj1;
        adjD = adj2;
    }

    var d;
    if (shapeType == "round") {
        d = "M0" + "," + (h / 2 + (1 - adjB) * (h / 2)) + " Q" + 0 + "," + h + " " + adjB * (w / 2) + "," + h + " L" + (w / 2 + (1 - adjC) * (w / 2)) + "," + h +
            " Q" + w + "," + h + " " + w + "," + (h / 2 + (h / 2) * (1 - adjC)) + "L" + w + "," + (h / 2) * adjD +
            " Q" + w + "," + 0 + " " + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + (w / 2) * adjA + ",0" +
            " Q" + 0 + "," + 0 + " 0," + (h / 2) * (adjA) + " z";
    } else if (shapeType == "snip") {
        d = "M0" + "," + adjA * (h / 2) + " L0" + "," + (h / 2 + (h / 2) * (1 - adjB)) + "L" + adjB * (w / 2) + "," + h +
            " L" + (w / 2 + (w / 2) * (1 - adjC)) + "," + h + "L" + w + "," + (h / 2 + (h / 2) * (1 - adjC)) +
            " L" + w + "," + adjD * (h / 2) + "L" + (w / 2 + (w / 2) * (1 - adjD)) + ",0 L" + ((w / 2) * adjA) + ",0 z";
    }
    return d;
}

/**
 * shapePie - 生成饼图路径
 * @param {number} H - 高度
 * @param {number} w - 宽度
 * @param {number} adj1 - 调整值1(起始角度)
 * @param {number} adj2 - 调整值2(结束角度)
 * @param {boolean} isClose - 是否闭合
 * @returns {Array} [路径字符串, 旋转字符串]
 */
export function shapePie(H, w, adj1, adj2, isClose) {
    var pieVal = parseInt(adj2);
    var piAngle = parseInt(adj1);
    var size = parseInt(H),
        radius = (size / 2),
        value = pieVal - piAngle;
    if (value < 0) {
        value = 360 + value;
    }
    value = Math.min(Math.max(value, 0), 360);

    var x = Math.cos((2 * Math.PI) / (360 / value));
    var y = Math.sin((2 * Math.PI) / (360 / value));

    var longArc, d, rot;
    if (isClose) {
        longArc = (value <= 180) ? 0 : 1;
        d = "M" + radius + "," + radius + " L" + radius + "," + 0 + " A" + radius + "," + radius + " 0 " + longArc + ",1 " + (radius + y * radius) + "," + (radius - x * radius) + " z";
        rot = "rotate(" + (piAngle - 270) + ", " + radius + ", " + radius + ")";
    } else {
        longArc = (value <= 180) ? 0 : 1;
        var radius1 = radius;
        var radius2 = w / 2;
        d = "M" + radius1 + "," + 0 + " A" + radius2 + "," + radius1 + " 0 " + longArc + ",1 " + (radius2 + y * radius2) + "," + (radius1 - x * radius1);
        rot = "rotate(" + (piAngle + 90) + ", " + radius + ", " + radius + ")";
    }

    return [d, rot];
}

/**
 * shapeGear - 生成齿轮形状路径
 * @param {number} w - 宽度
 * @param {number} h - 高度
 * @param {number} points - 点数(齿轮齿数)
 * @returns {string} SVG路径字符串
 */
export function shapeGear(w, h, points) {
    var innerRadius = h;
    var outerRadius = 1.5 * innerRadius;
    var cx = outerRadius;
    var cy = outerRadius;
    var notches = points;
    var radiusO = outerRadius;
    var radiusI = innerRadius;
    var taperO = 50;
    var taperI = 35;
    var pi2 = 2 * Math.PI;
    var angle = pi2 / (notches * 2);
    var taperAI = angle * taperI * 0.005;
    var taperAO = angle * taperO * 0.005;
    var a = angle;
    var toggle = false;

    var d = " M" + (cx + radiusO * Math.cos(taperAO)) + " " + (cy + radiusO * Math.sin(taperAO));

    for (; a <= pi2 + angle; a += angle) {
        if (toggle) {
            d += " L" + (cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
            d += " L" + (cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
        } else {
            d += " L" + (cx + radiusO * Math.cos(a - taperAO)) + "," + (cy + radiusO * Math.sin(a - taperAO));
            d += " L" + (cx + radiusI * Math.cos(a + taperAI)) + "," + (cy + radiusI * Math.sin(a + taperAI));
        }
        toggle = !toggle;
    }
    d += " ";
    return d;
}
