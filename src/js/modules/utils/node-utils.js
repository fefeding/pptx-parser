/**
 * 节点工具函数模块
 * 提供PPTX节点处理和索引功能
 */

var PPTXNodeUtils = (function() {
    /**
     * indexNodes - 索引幻灯片节点
     * @param {Object} content - 幻灯片内容
     * @returns {Object} 包含idTable、idxTable和typeTable的对象
     */
    function indexNodes(content) {
        var keys = Object.keys(content);
        var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];

        var idTable = {};
        var idxTable = {};
        var typeTable = {};

        for (var key in spTreeNode) {
            if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
                continue;
            }

            var targetNode = spTreeNode[key];

            if (targetNode.constructor === Array) {
                for (var i = 0; i < targetNode.length; i++) {
                    var nvSpPrNode = targetNode[i]["p:nvSpPr"];
                    var id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                    var idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                    var type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                    if (id !== undefined) {
                        idTable[id] = targetNode[i];
                    }
                    if (idx !== undefined) {
                        idxTable[idx] = targetNode[i];
                    }
                    if (type !== undefined) {
                        typeTable[type] = targetNode[i];
                    }
                }
            } else {
                var nvSpPrNode = targetNode["p:nvSpPr"];
                var id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                var idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                var type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);

                if (id !== undefined) {
                    idTable[id] = targetNode;
                }
                if (idx !== undefined) {
                    idxTable[idx] = targetNode;
                }
                if (type !== undefined) {
                    typeTable[type] = targetNode;
                }
            }
        }

        return { "idTable": idTable, "idxTable": idxTable, "typeTable": typeTable };
    }

    /**
     * processGroupSpNode - 处理组形状节点
     * @param {Object} node - 组形状节点
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 源
     * @returns {string} 生成的HTML
     */
    function processGroupSpNode(node, warpObj, source) {
        var slideFactor = 96 / 914400;
        var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
        if (xfrmNode !== undefined) {
            var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * slideFactor;
            var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * slideFactor;
            
            // 根据ECMA-376标准，a:chOff和a:chExt是可选元素
            // 当不存在时，应该使用父元素的对应值作为默认值
            var chx, chy, chcx, chcy;
            
            if (xfrmNode["a:chOff"] !== undefined && xfrmNode["a:chOff"]["attrs"] !== undefined) {
                chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * slideFactor;
                chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * slideFactor;
            } else {
                // 当a:chOff不存在时，使用a:off的值作为默认值
                chx = x;
                chy = y;
            }
            
            var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * slideFactor;
            var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * slideFactor;
            
            if (xfrmNode["a:chExt"] !== undefined && xfrmNode["a:chExt"]["attrs"] !== undefined) {
                chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * slideFactor;
                chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * slideFactor;
            } else {
                // 当a:chExt不存在时，使用a:ext的值作为默认值
                chcx = cx;
                chcy = cy;
            }
            var rotate = parseInt(xfrmNode["attrs"]["rot"]);
            var rotStr = "";
            var top = y - chy,
                left = x - chx,
                width = cx - chcx,
                height = cy - chcy;

            var sType = "group";
            if (!isNaN(rotate)) {
                rotate = PPTXXmlUtils.angleToDegrees(rotate);
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
        var grpStyle = "";

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
        var order = node["attrs"]["order"];

        var result = "<div class='block group' style='z-index: " + order + ";" + grpStyle + " border:1px solid red;'>";

        // Procsee all child nodes
        for (var nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (var i = 0; i < node[nodeKey].length; i++) {
                    result += processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType);
                }
            } else {
                result += processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType);
            }
        }

        result += "</div>";

        return result;
    }

    /**
     * processNodesInSlide - 处理幻灯片中的节点
     * @param {string} nodeKey - 节点键
     * @param {Object} nodeValue - 节点值
     * @param {Object} nodes - 节点集合
     * @param {Object} warpObj - 包装对象
     * @param {string} source - 源
     * @param {string} sType - 形状类型
     * @returns {string} 生成的HTML
     */
    function processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType) {
        var result = "";

        switch (nodeKey) {
            case "p:sp":    // Shape, Text
                result = processSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:cxnSp":    // Shape, Text (with connection)
                result = processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:pic":    // Picture
                result = processPicNode(nodeValue, warpObj, source, sType);
                break;
            case "p:graphicFrame":    // Chart, Diagram, Table
                result = processGraphicFrameNode(nodeValue, warpObj, source, sType);
                break;
            case "p:grpSp":
                result = processGroupSpNode(nodeValue, warpObj, source);
                break;
            case "mc:AlternateContent": //Equations and formulas as Image
                var mcFallbackNode = PPTXXmlUtils.getTextByPathList(nodeValue, ["mc:Fallback"]);
                result = processGroupSpNode(mcFallbackNode, warpObj, source);
                break;
            default:
                //console.log("nodeKey: ", nodeKey)
        }

        return result;
    }

    return {
        indexNodes: indexNodes,
        processGroupSpNode: processGroupSpNode,
        processNodesInSlide: processNodesInSlide
    };
})();

window.PPTXNodeUtils = PPTXNodeUtils;