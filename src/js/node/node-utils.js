import { PPTXUtils } from '../utils/utils.js';

class PPTXNodeUtils {
    /**
     * 处理幻灯片中的节点
     * @param {string} nodeKey - 节点键
     * @param {Object} nodeValue - 节点值
     * @param {Object} nodes - 节点集合
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 来源
     * @param {string} sType - 子类型
     * @param {Object} handlers - 处理函数集合
     * @returns {string} HTML字符串
     */
    static processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers) {
        let result = "";

        switch (nodeKey) {
            case "p:sp":    // Shape, Text
                result = handlers.processSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:cxnSp":    // Shape, Text (with connection)
                result = handlers.processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:pic":    // Picture
                result = handlers.processPicNode(nodeValue, warpObj, source, sType);
                break;
            case "p:graphicFrame":    // Chart, Diagram, Table
                result = handlers.processGraphicFrameNode(nodeValue, warpObj, source, sType);
                break;
            case "p:grpSp":
                result = handlers.processGroupSpNode(nodeValue, warpObj, source);
                break;
            case "mc:AlternateContent": // Equations and formulas as Image
                const mcFallbackNode = PPTXUtils.getTextByPathList(nodeValue, ["mc:Fallback"]);
                result = handlers.processGroupSpNode(mcFallbackNode, warpObj, source);
                break;
            default:
                // No action for unknown node types
        }

        return result;
    }

    /**
     * 处理组节点(包含多个子元素的组)
     * @param {Object} node - 组节点
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 来源
     * @param {number} slideFactor - 幻灯片缩放因子
     * @param {Function} processNodesInSlide - 处理幻灯片节点的函数
     * @returns {string} HTML字符串
     */
    static processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide) {
        const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
        let top, left, width, height;
        let grpStyle = "";
        let sType = "group";
        let rotate = 0;
        let rotStr = "";

        if (xfrmNode !== undefined) {
            let x, y, chx, chy, cx, cy, chcx, chcy;
            if (xfrmNode["a:off"] && xfrmNode["a:off"]["attrs"]) {
                x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
                y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
            }
            if (xfrmNode["a:chOff"] && xfrmNode["a:chOff"]["attrs"]) {
                chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
                chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;
            } else {
                chx = 0;
                chy = 0;
            }
            if (xfrmNode["a:ext"] && xfrmNode["a:ext"]["attrs"]) {
                cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * slideFactor;
                cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * slideFactor;
            }
            if (xfrmNode["a:chExt"] && xfrmNode["a:chExt"]["attrs"]) {
                chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * slideFactor;
                chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * slideFactor;
            } else {
                chcx = 0;
                chcy = 0;
            }

            if (xfrmNode["attrs"]) {
                rotate = parseInt(xfrmNode["attrs"]["rot"]);
            }

            if (y !== undefined && chy !== undefined) {
                top = y - chy;
            }
            if (x !== undefined && chx !== undefined) {
                left = x - chx;
            }
            if (cx !== undefined && chcx !== undefined) {
                width = cx - chcx;
            }
            if (cy !== undefined && chcy !== undefined) {
                height = cy - chcy;
            }

            if (!isNaN(rotate)) {
                rotate = PPTXUtils.angleToDegrees(rotate);
                rotStr += "transform: rotate(" + rotate + "deg) ; transform-origin: center;";
                if (rotate != 0) {
                    top = y;
                    left = x;
                    width = cx;
                    height = cy;
                    sType = "group-rotate";
                }
            }
        }

        if (rotStr !== undefined && rotStr != "") {
            grpStyle += rotStr;
        }

        if (top !== undefined) {
            grpStyle += "top: " + top + "px;";
        }
        if (left !== undefined) {
            grpStyle += "left: " + left + "px;";
        }
        if (width !== undefined) {
            grpStyle += "width:" + width + "px;";
        }
        if (height !== undefined) {
            grpStyle += "height: " + height + "px;";
        }

        const order = PPTXUtils.getTextByPathList(node, ["attrs", "order"]) || 0;
        let result = "<div class='block group' style='z-index: " + order + ";" + grpStyle + "'>";

        // Process all child nodes
        for (const nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (let i = 0; i < node[nodeKey].length; i++) {
                    result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType);
                }
            } else if (typeof node[nodeKey] === 'object' && nodeKey !== "attrs") {
                result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType);
            }
        }

        result += "</div>";
        return result;
    }

}

// 为了保持向后兼容性和全局访问，保留全局赋值
if (typeof window !== 'undefined') {
    window.PPTXNodeUtils = PPTXNodeUtils;
}

export { PPTXNodeUtils };
