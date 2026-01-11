/**
 * PPTX Parser - Shape Utilities Module
 * 形状处理工具模块 - 负责形状路径的生成和计算
 */

(function () {
    'use strict';

    // Initialize PPTXShapeUtils object if it doesn't exist
    if (!window.PPTXShapeUtils) {
        window.PPTXShapeUtils = {};
    }

    var PPTXUtils = window.PPTXUtils;

    function shapePie(H, w, adj1, adj2, isClose) {
        var pieVal = parseInt(adj2);
        var piAngle = parseInt(adj1);
        var size = parseInt(H),
            radius = (size / 2),
            value = pieVal - piAngle;
        if (value < 0) {
            value = 360 + value;
        }
        //console.log("value: ",value)      
        value = Math.min(Math.max(value, 0), 360);

        //calculate x,y coordinates of the point on the circle to draw the arc to. 
        var x = Math.cos((2 * Math.PI) / (360 / value));
        var y = Math.sin((2 * Math.PI) / (360 / value));


        //d is a string that describes the path of the slice.
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

    function shapeGear(w, h, points) {
        var innerRadius = h;//gear.innerRadius;
        var outerRadius = 1.5 * innerRadius;
        var cx = outerRadius;//Math.max(innerRadius, outerRadius),                   // center x
            cy = outerRadius;//Math.max(innerRadius, outerRadius),                    // center y
            notches = points,//gear.points,                      // num. of notches
            radiusO = outerRadius,                    // outer radius
            radiusI = innerRadius,                    // inner radius
            taperO = 50,                     // outer taper %
            taperI = 35,                     // inner taper %

            // pre-calculate values for loop

            pi2 = 2 * Math.PI,            // cache 2xPI (360deg)
            angle = pi2 / (notches * 2),    // angle between notches
            taperAI = angle * taperI * 0.005, // inner taper offset (100% = half notch)
            taperAO = angle * taperO * 0.005, // outer taper offset
            a = angle,                  // iterator (angle)
            toggle = false;
        // move to starting point
        var d = " M" + (cx + radiusO * Math.cos(taperAO)) + " " + (cy + radiusO * Math.sin(taperAO));

        // loop
        for (; a <= pi2 + angle; a += angle) {
            // draw inner to outer line
            if (toggle) {
                d += " L" + (cx + radiusI * Math.cos(a - taperAI)) + "," + (cy + radiusI * Math.sin(a - taperAI));
                d += " L" + (cx + radiusO * Math.cos(a + taperAO)) + "," + (cy + radiusO * Math.sin(a + taperAO));
            } else { // draw outer to inner line
                d += " L" + (cx + radiusO * Math.cos(a - taperAO)) + "," + (cy + radiusO * Math.sin(a - taperAO)); // outer line
                d += " L" + (cx + radiusI * Math.cos(a + taperAI)) + "," + (cy + radiusI * Math.sin(a + taperAI));// inner line

            }
            // switch level
            toggle = !toggle;
        }
        // close the final line
        d += " ";
        return d;
    }

    function shapeArc(cX, cY, rX, rY, stAng, endAng, isClose) {
        var dData;
        var angle = stAng;
        if (endAng >= stAng) {
            while (angle <= endAng) {
                var radians = angle * (Math.PI / 180);  // convert degree to radians
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
                var radians = angle * (Math.PI / 180);  // convert degree to radians
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

    function shapeSnipRoundRect(w, h, adj1, adj2, shapeType, adjType) {
        /* 
        shapeType: snip,round
        adjType: cornr1,cornr2,cornrAll,diag
        */
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
        //d is a string that describes the path of the slice.
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

    // 导出函数到全局对象
    window.PPTXShapeUtils.shapePie = shapePie;
    window.PPTXShapeUtils.shapeGear = shapeGear;
    window.PPTXShapeUtils.shapeArc = shapeArc;
    window.PPTXShapeUtils.shapeSnipRoundRect = shapeSnipRoundRect;

})();