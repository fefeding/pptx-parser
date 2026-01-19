import { PPTXUtils } from '../core/utils.js';

/**
 * 处理形状节点 (SpNode)
 *
 * @param {Object} node - 形状节点对象
 * @param {Object} pNode - 父节点对象
 * @param {Object} warpObj - 包装对象，包含解析上下文
 * @param {string} source - 来源标识
 * @param {string} sType - 形状类型标识
 * @param {Function} genShape - 形状生成函数
 * @returns {string} 生成的HTML字符串
 */
function processSpNode(node: any, pNode: any, warpObj: any, source: string, sType: string, genShape: Function): string {
    /*
    *  958    <xsd:complexType name="CT_GvmlShape">
    *  959   <xsd:sequence>
    *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
    *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
    *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
    *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
    *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    *  965   </xsd:sequence>
    *  966 </xsd:complexType>
    */

    let id: any = PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
    let name: any = PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
    let idx: any = (PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]) === undefined) ? undefined : PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
    let type: any = (PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) === undefined) ? undefined : PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    const order: any = PPTXUtils.getTextByPathList(node, ["attrs", "order"]);
    let isUserDrawnBg: boolean | undefined;
    if (source == "slideLayoutBg" || source == "slideMasterBg") {
        const userDrawn: any = PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
        if (userDrawn == "1") {
            isUserDrawnBg = true;
        } else {
            isUserDrawnBg = false;
        }
    }
    let slideLayoutSpNode: any;
    let slideMasterSpNode: any;
    let txBoxVal: any;

    if (idx !== undefined) {
        slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
        if (type !== undefined) {
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        } else {
            slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
        }
    } else {
        if (type !== undefined) {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        }
    }

    if (type === undefined) {
        txBoxVal = PPTXUtils.getTextByPathList(node, ["p:nvSpPr", "p:cNvSpPr", "attrs", "txBox"]);
        if (txBoxVal == "1") {
            type = "textBox";
        }
    }
    if (type === undefined) {
        type = PPTXUtils.getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
            //type = PPTXUtils.getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
            if (source == "diagramBg") {
                type = "diagram";
            } else {
                type = "obj"; //default type
            }
        }
    }
    //console.log("processSpNode type:", type, "idx:", idx);
    return genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source);
}

/**
 * 处理连接形状节点 (CxnSpNode)
 *
 * @param {Object} node - 连接形状节点对象
 * @param {Object} pNode - 父节点对象
 * @param {Object} warpObj - 包装对象，包含解析上下文
 * @param {string} source - 来源标识
 * @param {string} sType - 形状类型标识
 * @param {Function} genShape - 形状生成函数
 * @returns {string} 生成的HTML字符串
 */
function processCxnSpNode(node: any, pNode: any, warpObj: any, source: string, sType: string, genShape: Function): string {
    const id: any = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
    const name: any = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
    const idx: any = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    const type: any = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
    const order: any = node["attrs"]["order"];

    return genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source);
}

export { processSpNode, processCxnSpNode };