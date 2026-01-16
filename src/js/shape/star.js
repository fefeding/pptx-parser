    import { PPTXUtils } from '../core/utils.js';
    const PPTXStarShapes = {};

    // 4角星
    PPTXStarShapes.genStar4 = function(w, h, node, slideFactor) {
        var a, iwd2, ihd2, sdx, sdy, sx1, sx2, sy1, sy2, yAdj;
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 19098 * slideFactor;
        var cnstVal1 = 50000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            var name = shapAdjst["attrs"]["name"];
            if (name == "adj") {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        a = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
        iwd2 = wd2 * a / cnstVal1;
        ihd2 = hd2 * a / cnstVal1;
        sdx = iwd2 * Math.cos(0.7853981634);
        sdy = ihd2 * Math.sin(0.7853981634);
        sx1 = hc - sdx;
        sx2 = hc + sdx;
        sy1 = vc - sdy;
        sy2 = vc + sdy;
        yAdj = vc - ihd2;

        return "M0," + vc +
            " L" + sx1 + "," + sy1 +
            " L" + hc + ",0" +
            " L" + sx2 + "," + sy1 +
            " L" + w + "," + vc +
            " L" + sx2 + "," + sy2 +
            " L" + hc + "," + h +
            " L" + sx1 + "," + sy2 +
            " z";
    };

    // 5角星
    PPTXStarShapes.genStar5 = function(w, h, node, slideFactor) {
        var a, swd2, shd2, svc, dx1, dx2, dy1, dy2, x1, x2, x3, x4, y1, y2, iwd2, ihd2, sdx1, sdx2, sdy1, sdy2, sx1, sx2, sx3, sx4, sy1, sy2, sy3, yAdj;
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 19098 * slideFactor;
        var hf = 105146 * slideFactor;
        var vf = 110557 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }
        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        swd2 = wd2 * hf / cnstVal1;
        shd2 = hd2 * vf / cnstVal1;
        svc = vc * vf / cnstVal1;
        dx1 = swd2 * Math.cos(0.31415926536);
        dx2 = swd2 * Math.cos(5.3407075111);
        dy1 = shd2 * Math.sin(0.31415926536);
        dy2 = shd2 * Math.sin(5.3407075111);
        x1 = hc - dx1;
        x2 = hc - dx2;
        x3 = hc + dx2;
        x4 = hc + dx1;
        y1 = svc - dy1;
        y2 = svc - dy2;
        iwd2 = swd2 * a / maxAdj;
        ihd2 = shd2 * a / maxAdj;
        sdx1 = iwd2 * Math.cos(5.9690260418);
        sdx2 = iwd2 * Math.cos(0.94247779608);
        sdy1 = ihd2 * Math.sin(0.94247779608);
        sdy2 = ihd2 * Math.sin(5.9690260418);
        sx1 = hc - sdx1;
        sx2 = hc - sdx2;
        sx3 = hc + sdx2;
        sx4 = hc + sdx1;
        sy1 = svc - sdy1;
        sy2 = svc - sdy2;
        yAdj = vc - ihd2;

        return "M" + hc + ",0" +
            " L" + sx3 + "," + sy1 +
            " L" + x4 + "," + y1 +
            " L" + sx4 + "," + sy2 +
            " L" + hc + "," + h +
            " L" + sx1 + "," + sy2 +
            " L" + x1 + "," + y1 +
            " L" + sx2 + "," + sy1 +
            " z";
    };

    // 6角星
    PPTXStarShapes.genStar6 = function(w, h, node, slideFactor) {
        var a, swd2, dx1, x1, x2, y2, iwd2, ihd2, sdx2, sx1, sx2, sx3, sx4, sdy1, sy1, sy2, yAdj;
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2, hd4 = h / 4;
        var adj = 28868 * slideFactor;
        var hf = 115470 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }
        a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        swd2 = wd2 * hf / cnstVal1;
        dx1 = swd2 * Math.cos(0.5235987756);
        x1 = hc - dx1;
        x2 = hc + dx1;
        y2 = vc + hd4;
        iwd2 = swd2 * a / maxAdj;
        ihd2 = hd2 * a / maxAdj;
        sdx2 = iwd2 / 2;
        sx1 = hc - iwd2;
        sx2 = hc - sdx2;
        sx3 = hc + sdx2;
        sx4 = hc + iwd2;
        sdy1 = ihd2 * Math.sin(1.0471975512);
        sy1 = vc - sdy1;
        sy2 = vc + sdy1;
        yAdj = vc - ihd2;

        return "M" + x1 + "," + hd4 +
            " L" + sx2 + "," + sy1 +
            " L" + hc + ",0" +
            " L" + sx3 + "," + sy1 +
            " L" + x2 + "," + hd4 +
            " L" + sx4 + "," + vc +
            " L" + x2 + "," + y2 +
            " L" + sx3 + "," + sy2 +
            " L" + hc + "," + h +
            " L" + sx2 + "," + sy2 +
            " L" + x1 + "," + y2 +
            " L" + sx1 + "," + vc +
            " z";
    };

    // 7角星
    PPTXStarShapes.genStar7 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 105146 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 7角星的顶点计算
        var points = [];
        for (var i = 0; i < 14; i++) {
            var angle = (i * Math.PI) / 7 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 8角星
    PPTXStarShapes.genStar8 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 100000 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 8角星的顶点计算
        var points = [];
        for (var i = 0; i < 16; i++) {
            var angle = (i * Math.PI) / 8 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 10角星
    PPTXStarShapes.genStar10 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 105146 * slideFactor;
        var vf = 110557 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 10角星的顶点计算
        var points = [];
        for (var i = 0; i < 20; i++) {
            var angle = (i * Math.PI) / 10 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 12角星
    PPTXStarShapes.genStar12 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 100000 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 12角星的顶点计算
        var points = [];
        for (var i = 0; i < 24; i++) {
            var angle = (i * Math.PI) / 12 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 16角星
    PPTXStarShapes.genStar16 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 100000 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 16角星的顶点计算
        var points = [];
        for (var i = 0; i < 32; i++) {
            var angle = (i * Math.PI) / 16 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 24角星
    PPTXStarShapes.genStar24 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 100000 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 24角星的顶点计算
        var points = [];
        for (var i = 0; i < 48; i++) {
            var angle = (i * Math.PI) / 24 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 32角星
    PPTXStarShapes.genStar32 = function(w, h, node, slideFactor) {
        var hc = w / 2, vc = h / 2, wd2 = w / 2, hd2 = h / 2;
        var adj = 12500 * slideFactor;
        var hf = 100000 * slideFactor;
        var vf = 100000 * slideFactor;
        var maxAdj = 50000 * slideFactor;
        var cnstVal1 = 100000 * slideFactor;
        var shapAdjst = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                var name = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        var a = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        var swd2 = wd2 * hf / cnstVal1;
        var shd2 = hd2 * vf / cnstVal1;
        var iwd2 = swd2 * a / maxAdj;
        var ihd2 = hd2 * a / maxAdj;

        // 32角星的顶点计算
        var points = [];
        for (var i = 0; i < 64; i++) {
            var angle = (i * Math.PI) / 32 - Math.PI / 2;
            var isOuter = i % 2 === 0;
            var r = isOuter ? swd2 : iwd2;
            var rh = isOuter ? shd2 : ihd2;
            var x = hc + r * Math.cos(angle);
            var y = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

export { PPTXStarShapes };

// Also export to global scope for backward compatibility
// window.PPTXStarShapes = PPTXStarShapes; // Removed for ES modules
