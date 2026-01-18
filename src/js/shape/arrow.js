    import { PPTXUtils } from '../core/utils.js';
    const PPTXArrowShapes = {};

    // 右箭头
    PPTXArrowShapes.genRightArrow = function(w, h, node, slideFactor) {
        const shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        var sAdj1, sAdj1_val = 0.25;
        var sAdj2, sAdj2_val = 0.5;
        const max_sAdj2_const = w / h;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                const sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    const sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = 1 - ((sAdj2_val2) / max_sAdj2_const);
                }
            }
        }

        return "polygon points='" + w + " " + h / 2 + "," + sAdj2_val * w + " 0," + sAdj2_val * w + " " + sAdj1_val * h + ",0 " + sAdj1_val * h +
            ",0 " + (1 - sAdj1_val) * h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + ", " + sAdj2_val * w + " " + h + "'";
    };

    // 左箭头
    PPTXArrowShapes.genLeftArrow = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 0.25;
        let sAdj2, sAdj2_val = 0.5;
        let max_sAdj2_const = w / h;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    let sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }

        return "polygon points='0 " + h / 2 + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + "," + w + " " + (1 - sAdj1_val) * h +
            "," + w + " " + sAdj1_val * h + "," + sAdj2_val * w + " " + sAdj1_val * h + ", " + sAdj2_val * w + " 0'";
    };

    // 下箭头
    PPTXArrowShapes.genDownArrow = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 0.25;
        let sAdj2, sAdj2_val = 0.5;
        let max_sAdj2_const = h / w;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }

        return "polygon points='" + (0.5 - sAdj1_val) * w + " 0," + (0.5 - sAdj1_val) * w + " " + (1 - sAdj2_val) * h + ",0 " + (1 - sAdj2_val) * h + "," + (w / 2) + " " + h +
            "," + w + " " + (1 - sAdj2_val) * h + "," + (0.5 + sAdj1_val) * w + " " + (1 - sAdj2_val) * h + ", " + (0.5 + sAdj1_val) * w + " 0'";
    };

    // 上箭头
    PPTXArrowShapes.genUpArrow = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 0.25;
        let sAdj2, sAdj2_val = 0.5;
        let max_sAdj2_const = h / w;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = parseInt(sAdj1.substr(4)) / 200000;
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    let sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }

        return "polygon points='" + (w / 2) + " 0,0 " + sAdj2_val * h + "," + (0.5 - sAdj1_val) * w + " " + sAdj2_val * h + "," + (0.5 - sAdj1_val) * w + " " + h +
            "," + (0.5 + sAdj1_val) * w + " " + h + "," + (0.5 + sAdj1_val) * w + " " + sAdj2_val * h + ", " + w + " " + sAdj2_val * h + "'";
    };

    // 左右箭头
    PPTXArrowShapes.genLeftRightArrow = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 0.25;
        let sAdj2, sAdj2_val = 0.25;
        let max_sAdj2_const = w / h;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    let sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }

        return "polygon points='0 " + h / 2 + "," + sAdj2_val * w + " " + h + "," + sAdj2_val * w + " " + (1 - sAdj1_val) * h + "," + (1 - sAdj2_val) * w + " " + (1 - sAdj1_val) * h +
            "," + (1 - sAdj2_val) * w + " " + h + "," + w + " " + h / 2 + ", " + (1 - sAdj2_val) * w + " 0," + (1 - sAdj2_val) * w + " " + sAdj1_val * h + "," +
            sAdj2_val * w + " " + sAdj1_val * h + "," + sAdj2_val * w + " 0'";
    };

    // 上下箭头
    PPTXArrowShapes.genUpDownArrow = function(w, h, node, slideFactor) {
        let shapAdjst_ary = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "a:avLst", "a:gd"]);
        let sAdj1, sAdj1_val = 0.25;
        let sAdj2, sAdj2_val = 0.25;
        let max_sAdj2_const = h / w;
        if (shapAdjst_ary !== undefined) {
            for (let i = 0; i < shapAdjst_ary.length; i++) {
                let sAdj_name = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "name"]);
                if (sAdj_name == "adj1") {
                    sAdj1 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    sAdj1_val = 0.5 - (parseInt(sAdj1.substr(4)) / 200000);
                } else if (sAdj_name == "adj2") {
                    sAdj2 = PPTXUtils.getTextByPathList(shapAdjst_ary[i], ["attrs", "fmla"]);
                    let sAdj2_val2 = parseInt(sAdj2.substr(4)) / 100000;
                    sAdj2_val = (sAdj2_val2) / max_sAdj2_const;
                }
            }
        }

        return "polygon points='" + w / 2 + " 0,0 " + sAdj2_val * h + "," + sAdj1_val * w + " " + sAdj2_val * h + "," + sAdj1_val * w + " " + (1 - sAdj2_val) * h +
            ",0 " + (1 - sAdj2_val) * h + "," + w / 2 + " " + h + ", " + w + " " + (1 - sAdj2_val) * h + "," + (1 - sAdj1_val) * w + " " + (1 - sAdj2_val) * h + "," +
            (1 - sAdj1_val) * w + " " + sAdj2_val * h + "," + w + " " + sAdj2_val * h + "'";
    };

export { PPTXArrowShapes };
