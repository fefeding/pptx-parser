import { PPTXUtils } from '../core/utils.js';

interface PPTXStarShapesType {
    genStar4: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar5: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar6: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar7: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar8: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar10: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar12: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar16: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar24: (w: number, h: number, node: any, slideFactor: number) => string;
    genStar32: (w: number, h: number, node: any, slideFactor: number) => string;
}

const PPTXStarShapes: PPTXStarShapesType = {} as any;

    // 4角星
    PPTXStarShapes.genStar4 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 19098 * slideFactor;
        const cnstVal1: number = 50000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            const name: string = shapAdjst["attrs"]["name"];
            if (name == "adj") {
                adj = parseInt(shapAdjst["attrs"]["fmla"].substr(4)) * slideFactor;
            }
        }
        const a: any = (adj < 0) ? 0 : (adj > cnstVal1) ? cnstVal1 : adj;
        const iwd2: any = wd2 * a / cnstVal1;
        const ihd2: any = hd2 * a / cnstVal1;
        const sdx: any = iwd2 * Math.cos(0.7853981634);
        const sdy: any = ihd2 * Math.sin(0.7853981634);
        const sx1: any = hc - sdx;
        const sx2: any = hc + sdx;
        const sy1: any = vc - sdy;
        const sy2: any = vc + sdy;
        const yAdj: any = vc - ihd2;

        return "M0," + vc +
            " L" + sx1 + "," + sy1 +
            " L" + hc + `,0 L` + sx2 + "," + sy1 +
            " L" + w + "," + vc +
            " L" + sx2 + "," + sy2 +
            " L" + hc + "," + h +
            " L" + sx1 + "," + sy2 +
            " z";
    };

    // 5角星
    PPTXStarShapes.genStar5 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 19098 * slideFactor;
        const hf: number = 105146 * slideFactor;
        const vf: number = 110557 * slideFactor;
        const maxAdj: number = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }
        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const svc: any = vc * vf / cnstVal1;
        const dx1: any = swd2 * Math.cos(0.31415926536);
        const dx2: any = swd2 * Math.cos(5.3407075111);
        const dy1: any = shd2 * Math.sin(0.31415926536);
        const dy2: any = shd2 * Math.sin(5.3407075111);
        const x1: any = hc - dx1;
        const x2: any = hc - dx2;
        const x3: any = hc + dx2;
        const x4: any = hc + dx1;
        const y1: any = svc - dy1;
        const y2: any = svc - dy2;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = shd2 * a / maxAdj;
        const sdx1: any = iwd2 * Math.cos(5.9690260418);
        const sdx2: any = iwd2 * Math.cos(0.94247779608);
        const sdy1: any = ihd2 * Math.sin(0.94247779608);
        const sdy2: any = ihd2 * Math.sin(5.9690260418);
        const sx1: any = hc - sdx1;
        const sx2: any = hc - sdx2;
        const sx3: any = hc + sdx2;
        const sx4: any = hc + sdx1;
        const sy1: any = svc - sdy1;
        const sy2: any = svc - sdy2;
        const sy3: any = svc + sdy2;
        const yAdj: any = vc - ihd2;

        return "M" + hc + `,0 L` + sx3 + "," + sy1 +
            " L" + x4 + "," + y1 +
            " L" + sx4 + "," + sy2 +
            " L" + hc + "," + h +
            " L" + sx1 + "," + sy2 +
            " L" + x1 + "," + y1 +
            " L" + sx2 + "," + sy1 +
            " z";
    };

    // 6角星
    PPTXStarShapes.genStar6 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2, hd4: number = h / 4;
        let adj: any = 28868 * slideFactor;
        const hf: any = 115470 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }
        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const dx1: any = swd2 * Math.cos(0.5235987756);
        const x1: any = hc - dx1;
        const x2: any = hc + dx1;
        const y2: any = vc + hd4;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;
        const sdx2: any = iwd2 / 2;
        const sx1: any = hc - iwd2;
        const sx2: any = hc - sdx2;
        const sx3: any = hc + sdx2;
        const sx4: any = hc + iwd2;
        const sdy1: any = ihd2 * Math.sin(1.0471975512);
        const sy1: any = vc - sdy1;
        const sy2: any = vc + sdy1;
        const yAdj: any = vc - ihd2;

        return "M" + x1 + "," + hd4 +
            " L" + sx2 + "," + sy1 +
            " L" + hc + `,0 L` + sx3 + "," + sy1 +
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
    PPTXStarShapes.genStar7 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 105146 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 7角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 14; i++) {
            const angle: number = (i * Math.PI) / 7 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 8角星
    PPTXStarShapes.genStar8 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 100000 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 8角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 16; i++) {
            const angle: number = (i * Math.PI) / 8 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 10角星
    PPTXStarShapes.genStar10 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 105146 * slideFactor;
        const vf: any = 110557 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 10角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 20; i++) {
            const angle: number = (i * Math.PI) / 10 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 12角星
    PPTXStarShapes.genStar12 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 100000 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 12角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 24; i++) {
            const angle: number = (i * Math.PI) / 12 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 16角星
    PPTXStarShapes.genStar16 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 100000 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 16角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 32; i++) {
            const angle: number = (i * Math.PI) / 16 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 24角星
    PPTXStarShapes.genStar24 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 100000 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 24角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 48; i++) {
            const angle: number = (i * Math.PI) / 24 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

    // 32角星
    PPTXStarShapes.genStar32 = function(w: number, h: number, node: any, slideFactor: number): string {
        const hc: number = w / 2, vc: number = h / 2, wd2: number = w / 2, hd2: number = h / 2;
        let adj: any = 12500 * slideFactor;
        const hf: any = 100000 * slideFactor;
        const vf: any = 100000 * slideFactor;
        const maxAdj: any = 50000 * slideFactor;
        const cnstVal1: any = 100000 * slideFactor;
        const shapAdjst: any = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);

        if (shapAdjst !== undefined) {
            Object.keys(shapAdjst).forEach(function (key) {
                const name: any = shapAdjst[key]["attrs"]["name"];
                if (name == "adj") {
                    adj = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "hf") {
                    // hf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                } else if (name == "vf") {
                    // vf = parseInt(shapAdjst[key]["attrs"]["fmla"].substr(4)) * slideFactor;
                }
            });
        }

        const a: any = (adj < 0) ? 0 : (adj > maxAdj) ? maxAdj : adj;
        const swd2: any = wd2 * hf / cnstVal1;
        const shd2: any = hd2 * vf / cnstVal1;
        const iwd2: any = swd2 * a / maxAdj;
        const ihd2: any = hd2 * a / maxAdj;

        // 32角星的顶点计算
        const points: string[] = [];
        for (let i = 0; i < 64; i++) {
            const angle: number = (i * Math.PI) / 32 - Math.PI / 2;
            const isOuter: boolean = i % 2 === 0;
            const r: any = isOuter ? swd2 : iwd2;
            const rh: any = isOuter ? shd2 : ihd2;
            const x: any = hc + r * Math.cos(angle);
            const y: any = vc + rh * Math.sin(angle);
            points.push((i === 0 ? "M" : "L") + x + "," + y);
        }

        return points.join(" ") + " z";
    };

export { PPTXStarShapes };
