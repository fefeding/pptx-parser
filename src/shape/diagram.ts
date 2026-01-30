import { PPTXUtils } from '../core/utils.js';

interface PPTXDiagramUtilsType {
    genDiagram: (node: any, warpObj: any, source: string, sType: string, readXmlFile: Function, getPosition: Function, getSize: Function, processSpNode: Function) => Promise<string>;
}

const PPTXDiagramUtils = {} as PPTXDiagramUtilsType;

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
PPTXDiagramUtils.genDiagram = async function(node: any, warpObj: any, source: string, sType: string, readXmlFile: Function, getPosition: Function, getSize: Function, processSpNode: Function): Promise<string> {
    const order: string = node["attrs"]["order"];
    const zip: any = warpObj["zip"];
    const xfrmNode: any = PPTXUtils.getTextByPathList(node, ["p:xfrm"]);
    const dgmRelIds: any = PPTXUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "dgm:relIds", "attrs"]);

    if (!dgmRelIds) {
        return "";
    }

    // 获取图表相关文件的ID
    const dgmClrFileId: string = dgmRelIds["r:cs"];
    const dgmDataFileId: string = dgmRelIds["r:dm"];
    const dgmLayoutFileId: string = dgmRelIds["r:lo"];
    const dgmQuickStyleFileId: string = dgmRelIds["r:qs"];

    // 获取文件名
    const dgmClrFileName: string = warpObj["slideResObj"][dgmClrFileId].target;
    const dgmDataFileName: string = warpObj["slideResObj"][dgmDataFileId].target;
    const dgmLayoutFileName: string = warpObj["slideResObj"][dgmLayoutFileId].target;
    const dgmQuickStyleFileName: string = warpObj["slideResObj"][dgmQuickStyleFileId].target;

    // 读取XML文件
    const dgmClr: any = await readXmlFile(zip, dgmClrFileName);
    const dgmData: any = await readXmlFile(zip, dgmDataFileName);
    const dgmLayout: any = await readXmlFile(zip, dgmLayoutFileName);
    const dgmQuickStyle: any = await readXmlFile(zip, dgmQuickStyleFileName);

    // 获取绘图文件内容
    const dgmDrwSpArray: any = PPTXUtils.getTextByPathList(warpObj["digramFileContent"], ["p:drawing", "p:spTree", "p:sp"]);
    let rslt: string = "";

    if (dgmDrwSpArray !== undefined) {
        const dgmDrwSpArrayLen: number = dgmDrwSpArray.length;
        for (let i: number = 0; i < dgmDrwSpArrayLen; i++) {
            const dspSp: any = dgmDrwSpArray[i];
            rslt += processSpNode(dspSp, node, warpObj, "diagramBg", sType);
        }
    }

    return "<div class='block diagram-content' style='" +
        getPosition(xfrmNode, node, undefined, undefined, sType) +
        getSize(xfrmNode, undefined, undefined) +
        "'>" + rslt + "</div>";
};


export { PPTXDiagramUtils };