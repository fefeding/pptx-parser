
import { PPTXUtils } from '../utils/utils.js';

var PPTXDiagramUtils = {};

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
    var order = node["attrs"]["order"];
    var zip = warpObj["zip"];
    var xfrmNode = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    var dgmRelIds = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);

    if (!dgmRelIds) {
        return "";
    }

    // 获取图表相关文件的ID
    var dgmClrFileId = dgmRelIds["r:cs"];
    var dgmDataFileId = dgmRelIds["r:dm"];
    var dgmLayoutFileId = dgmRelIds["r:lo"];
    var dgmQuickStyleFileId = dgmRelIds["r:qs"];

    // 获取文件名
    var dgmClrFileName = warpObj["slideResObj"][dgmClrFileId].target;
    var dgmDataFileName = warpObj["slideResObj"][dgmDataFileId].target;
    var dgmLayoutFileName = warpObj["slideResObj"][dgmLayoutFileId].target;
    var dgmQuickStyleFileName = warpObj["slideResObj"][dgmQuickStyleFileId].target;

    // 读取XML文件
    var dgmClr = readXmlFile(zip, dgmClrFileName);
    var dgmData = readXmlFile(zip, dgmDataFileName);
    var dgmLayout = readXmlFile(zip, dgmLayoutFileName);
    var dgmQuickStyle = readXmlFile(zip, dgmQuickStyleFileName);

    // 获取绘图文件内容
    var dgmDrwSpArray = PPTXUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
    var rslt = "";

    if (dgmDrwSpArray !== undefined) {
        var dgmDrwSpArrayLen = dgmDrwSpArray.length;
        for (var i = 0; i < dgmDrwSpArrayLen; i++) {
            var dspSp = dgmDrwSpArray[i];
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
