import { PPTXUtils } from '../core/utils';
import { PPTXColorUtils } from '../core/color';

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
    static async processNodesInSlide(nodeKey, nodeValue, nodes, warpObj, source, sType, handlers) {
        let result = "";

        switch (nodeKey) {
            case "p:sp":    // Shape, Text
                result = await handlers.processSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:cxnSp":    // Shape, Text (with connection)
                result = await handlers.processCxnSpNode(nodeValue, nodes, warpObj, source, sType);
                break;
            case "p:pic":    // Picture
                result = handlers.processPicNode(nodeValue, warpObj, source, sType);
                break;
            case "p:graphicFrame":    // Chart, Diagram, Table
                result = await handlers.processGraphicFrameNode(nodeValue, warpObj, source, sType);
                break;
            case "p:grpSp":
                result = await handlers.processGroupSpNode(nodeValue, warpObj, source);
                break;
            case "mc:AlternateContent": // Equations and formulas as Image
                // 尝試從選擇內容中獲取數學公式
                const mcChoiceNode = PPTXUtils.getTextByPathList(nodeValue, ["mc:Choice"]);
                if (mcChoiceNode) {
                    // 檢查是否包含OMath內容
                    const graphicData = PPTXUtils.getTextByPathList(mcChoiceNode, ["a:graphic", "a:graphicData"]);
                    if (graphicData) {
                        // 檢查是否包含OMath內容
                        const oMathPara = PPTXUtils.getTextByPathList(graphicData, ["m:oMathPara"]);
                        if (oMathPara) {
                            // 如果找到OMath內容，生成數學公式占位符
                            result = PPTXNodeUtils.generateMathPlaceholder(oMathPara);
                            break;
                        }
                        
                        const oMath = PPTXUtils.getTextByPathList(graphicData, ["m:oMath"]);
                        if (oMath) {
                            result = PPTXNodeUtils.generateMathPlaceholder(oMath);
                            break;
                        }
                    }
                }
                
                // 如果沒有找到數學公式，嘗試從備選內容獲取
                const mcFallbackNode = PPTXUtils.getTextByPathList(nodeValue, ["mc:Fallback"]);
                if (mcFallbackNode) {
                    // 檢查備選內容中是否包含數學公式
                    const fallbackGraphicData = PPTXUtils.getTextByPathList(mcFallbackNode, ["a:graphic", "a:graphicData"]);
                    if (fallbackGraphicData) {
                        const fallbackOMathPara = PPTXUtils.getTextByPathList(fallbackGraphicData, ["m:oMathPara"]);
                        if (fallbackOMathPara) {
                            result = PPTXNodeUtils.generateMathPlaceholder(fallbackOMathPara);
                            break;
                        }
                        
                        const fallbackOMath = PPTXUtils.getTextByPathList(fallbackGraphicData, ["m:oMath"]);
                        if (fallbackOMath) {
                            result = PPTXNodeUtils.generateMathPlaceholder(fallbackOMath);
                            break;
                        }
                    }
                    
                    result = handlers.processGroupSpNode(mcFallbackNode, warpObj, source);
                }
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
    static async processGroupSpNode(node, warpObj, source, slideFactor, processNodesInSlide) {
        let result = "";
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

        const order = node["attrs"]["order"];
        result = "<div class='block group' style='z-index: " + order + ";" + grpStyle + "'>";

        // Process all child nodes
        for (const nodeKey in node) {
            if (node[nodeKey].constructor === Array) {
                for (let i = 0; i < node[nodeKey].length; i++) {
                    result += await processNodesInSlide(nodeKey, node[nodeKey][i], node, warpObj, source, sType);
                }
            } else if (typeof node[nodeKey] === 'object' && nodeKey !== "attrs") {
                result += await processNodesInSlide(nodeKey, node[nodeKey], node, warpObj, source, sType);
            }
        }

        result += "</div>";
        return result;
    }

    /**
     * 生成数学公式占位符
     * @param {any} oMathContent - OMath内容节点
     * @returns {string} HTML表示
     */
    static generateMathPlaceholder(oMathContent: any): string {
        // 尝試從OMath內容中提取簡單的文本表示
        const mathText = PPTXNodeUtils.extractSimpleMathText(oMathContent);
        
        if (mathText) {
            return `<span class="math-placeholder" style="font-style:italic; color:#0066cc;" title="数学公式">${mathText}</span>`;
        }
        
        return "<span class='math-placeholder' style='font-style:italic; color:#0066cc;' title='数学公式'>[数学公式]</span>";
    }

    /**
     * 从OMath内容中提取简单文本
     * @param {any} oMathContent - OMath内容节点
     * @returns {string} 提取的文本
     */
    static extractSimpleMathText(oMathContent: any): string {
        if (!oMathContent) return '';

        // 递归搜索文本节点
        const findTextNodes = (node: any): string[] => {
            let texts: string[] = [];
            
            if (typeof node === 'object' && node !== null) {
                for (const key in node) {
                    if (key === 'm:t' && typeof node[key] === 'string') {
                        // 直接文本节点
                        texts.push(node[key]);
                    } else if (key === 'm:r' && typeof node[key] === 'object') {
                        // 寻找文本运行中的文本
                        const run = node[key];
                        if (Array.isArray(run)) {
                            for (const r of run) {
                                if (r && r['m:t']) {
                                    texts.push(r['m:t']);
                                }
                            }
                        } else if (run && run['m:t']) {
                            texts.push(run['m:t']);
                        }
                    } else if (key === 'm:sSub' || key === 'm:sSup' || key === 'm:f' || key === 'm:e' || key === 'm:num' || key === 'm:den' || key === 'm:d' || key === 'm:r') {
                        // 数学结构节点：下标、上标、分数、分子、分母、分隔符等
                        texts = texts.concat(findTextNodes(node[key]));
                    } else if (typeof node[key] === 'object') {
                        texts = texts.concat(findTextNodes(node[key]));
                    }
                }
            }
            
            return texts;
        };

        const textNodes = findTextNodes(oMathContent);
        return textNodes.join('');
    }
}

export { PPTXNodeUtils };