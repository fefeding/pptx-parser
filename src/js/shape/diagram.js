
import { PPTXUtils } from '../core/utils.js';

const PPTXDiagramUtils = {};

    /**
 * 生成图表/图表演示 (SmartArt)
 * @param {Object} node - 图表节点
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 来源
 * @param {string} sType - 子类型
 * @param {Function} readXmlFile - 读取XML文件的函数
 * @param {Function} getPosition - 获取位置的函数
 * @param {Function} getSize - 获取尺寸的函数
 * @param {Function} processSpNode - 处理形状节点的函数
 * @returns {string} HTML字符串
 */
PPTXDiagramUtils.genDiagram = function(node, warpObj, source, sType, readXmlFile, getPosition, getSize, processSpNode) {
    const order = node["attrs"]["order"];
    const zip = warpObj["zip"];
    const xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    const dgmRelIds = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);

    if (!dgmRelIds) {
        return "";
    }

    // 获取图表相关文件的ID
    const dgmClrFileId = dgmRelIds["r:cs"];
    const dgmDataFileId = dgmRelIds["r:dm"];
    const dgmLayoutFileId = dgmRelIds["r:lo"];
    const dgmQuickStyleFileId = dgmRelIds["r:qs"];

    // 获取文件名
    const dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target;
    const dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target;
    const dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
    const dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;

    // 读取XML文件
    const dgmClr = readXmlFile(zip, dgmClrFileName);
    const dgmData = readXmlFile(zip, dgmDataFileName);
    const dgmLayout = readXmlFile(zip, dgmLayoutFileName);
    const dgmQuickStyle = readXmlFile(zip, dgmQuickStyleFileName);

    // 获取绘图文件内容
    const dgmDrwSpArray = PPTXUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
    const rslt = "";

    if (dgmDrwSpArray !== undefined) {
        const dgmDrwSpArrayLen = dgmDrwSpArray.length;
        for (let i = 0; i < dgmDrwSpArrayLen; i++) {
            const dspSp = dgmDrwSpArray[i];
            rslt += processSpNode(dspSp, node, warpObj, "diagramBg", sType);
        }
    }

    return "<div class='block diagram-content' style='" +
        getPosition(xfrmNode, node, undefined, undefined, sType) +
        getSize(xfrmNode, undefined, undefined) +
        "'>" + rslt + "</div>";
};


export { PPTXDiagramUtils };

// Also export to global scope for backward compatibility
// window.PPTXDiagramUtils = PPTXDiagramUtils; // Removed for ES modules
