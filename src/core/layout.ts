/**
 * PPTX 布局工具函数
 * 用于处理幻灯片布局和母版节点
 */

import { PPTXUtils } from './utils';

interface WarpObj {
    slideLayoutTables?: {
        idxTable: any[];
        typeTable: Record<string, any>;
    };
    slideMasterTextStyles?: {
        'p:titleStyle'?: any;
        'p:bodyStyle'?: any;
        'p:otherStyle'?: any;
    };
    defaultTextStyle?: any;
    slideMasterTables?: {
        typeTable: Record<string, any>;
    };
}

/**
 * 获取布局和母版节点
 * @param node - 段落节点
 * @param idx - 索引
 * @param type - 类型
 * @param warpObj - 包装对象
 * @returns 包含 nodeLaout 和 nodeMaster 属性的对象
 */
export function getLayoutAndMasterNode(node: any, idx: number | undefined, type: string | undefined, warpObj: WarpObj): {
    nodeLaout: any;
    nodeMaster: any;
} {
    let pPrNodeLaout: any, pPrNodeMaster: any;
    const pPrNode = node["a:pPr"];
    //lvl
    let lvl = 1;
    const lvlNode = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    if (idx !== undefined) {
        //slidelayout
        pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj.slideLayoutTables?.idxTable?.[idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj.slideLayoutTables?.idxTable?.[idx], ["p:txBody", "a:p", "a:pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj.slideLayoutTables?.idxTable?.[idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
            }
        }
    }
    if (type !== undefined) {
        //slidelayout
        const lvlStr = "a:lvl" + lvl + "pPr";
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
        //masterlayout
        if (type == "title" || type == "ctrTitle") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
        } else if (type == "body" || type == "obj" || type == "subTitle") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
        } else if (type == "shape" || type == "diagram") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
        } else if (type == "textBox") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
        } else {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
    }
    return {
        "nodeLaout": pPrNodeLaout,
        "nodeMaster": pPrNodeMaster
    };
}

// 向后兼容的导出对象
const PPTXLayoutUtils = {
    getLayoutAndMasterNode
};

export { PPTXLayoutUtils };