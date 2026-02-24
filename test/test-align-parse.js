import tXml from '../src/js/core/tXml.js';

// 模拟 PPTXXmlUtils.getTextByPathList 函数
function getTextByPathList(node, path) {
    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }

    if (node === undefined) {
        return undefined;
    }

    let l = path.length;
    for (let i = 0; i < l; i++) {
        node = node[path[i]];
        if (node === undefined) {
            return undefined;
        }
    }

    return node;
}

// 模拟 getHorizontalAlign 函数
function getHorizontalAlign(node, textBodyNode, idx, type, prg_dir, warpObj) {
    let algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    console.log('Initial algn:', algn);
    if (algn === undefined) {
        //let layoutMasterNode = getLayoutAndMasterNode(node, idx, type, warpObj);
        // let pPrNodeLaout = layoutMasterNode.nodeLaout;
        // let pPrNodeMaster = layoutMasterNode.nodeMaster;
        let lvlIdx = 1;
        let lvlNode = getTextByPathList(node, ["a:pPr", "attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvlIdx = parseInt(lvlNode) + 1;
        }
        let lvlStr = "a:lvl" + lvlIdx + "pPr";

        let lstStyle = textBodyNode["a:lstStyle"];
        algn = getTextByPathList(lstStyle, [lvlStr, "attrs", "algn"]);
        console.log('After lstStyle algn:', algn);

        if (algn === undefined && idx !== undefined ) {
            //slidelayout
            algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
            console.log('After slideLayout algn:', algn);
            if (algn === undefined) {
                algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
                console.log('After slideLayout a:p algn:', algn);
                if (algn === undefined) {
                    algn = getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvlIdx - 1), "a:pPr", "attrs", "algn"]);
                    console.log('After slideLayout a:p index algn:', algn);
                }
            }
        }
        if (algn === undefined) {
            if (type !== undefined) {
                //slidelayout
                algn = getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                console.log('After slideLayout typeTable algn:', algn);

                if (algn === undefined) {
                    //masterlayout
                    if (type == "title" || type == "ctrTitle") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr, "attrs", "algn"]);
                        console.log('After slideMasterTextStyles titleStyle algn:', algn);
                    } else if (type == "body" || type == "obj" || type == "subTitle") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                        console.log('After slideMasterTextStyles bodyStyle algn:', algn);
                    } else if (type == "shape" || type == "diagram") {
                        algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr, "attrs", "algn"]);
                        console.log('After slideMasterTextStyles otherStyle algn:', algn);
                    } else if (type == "textBox") {
                        algn = getTextByPathList(warpObj, ["defaultTextStyle", lvlStr, "attrs", "algn"]);
                        console.log('After defaultTextStyle algn:', algn);
                    } else {
                        algn = getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr, "attrs", "algn"]);
                        console.log('After slideMasterTables typeTable algn:', algn);
                    }
                }
            } else {
                algn = getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr, "attrs", "algn"]);
                console.log('After slideMasterTextStyles bodyStyle (default) algn:', algn);
            }
        }
    }

    console.log('Final algn before switch:', algn);
    if (algn === undefined) {
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            return "h-mid";
        } else if (type == "sldNum") {
            return "h-right";
        } else {
            // 默认返回左对齐
            return "h-left";
        }
    }
    if (algn !== undefined) {
        switch (algn) {
            case "l":
                if (prg_dir == "pregraph-rtl"){
                    //return "h-right";
                    return "h-left-rtl";
                }else{
                    return "h-left";
                }
                break;
            case "r":
                if (prg_dir == "pregraph-rtl") {
                    //return "h-left";
                    return "h-right-rtl";
                }else{
                    return "h-right";
                }
                break;
            case "ctr":
                return "h-mid";
                break;
            case "just":
            case "dist":
            default:
                return "h-" + algn;
        }
    }
    //return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
}

// 测试 XML 解析
const xml = `
<p:sp>
    <p:txBody>
        <a:bodyPr wrap="square">
            <a:spAutoFit />
        </a:bodyPr>
        <a:lstStyle />
        <a:p>
            <a:pPr marL="0" marR="0" lvl="0" indent="0" algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" fontAlgn="auto" latinLnBrk="0" hangingPunct="1">
                <a:lnSpc>
                    <a:spcPct val="100000" />
                </a:lnSpc>
                <a:spcBef>
                    <a:spcPts val="0" />
                </a:spcBef>
                <a:spcAft>
                    <a:spcPts val="0" />
                </a:spcAft>
                <a:buClrTx />
                <a:buSzTx />
                <a:buFontTx />
                <a:buNone />
                <a:tabLst />
                <a:defRPr />
            </a:pPr>
            <a:r>
                <a:rPr kumimoji="1" lang="zh-CN" altLang="en-US" sz="2000" b="0" i="0" u="none" strike="noStrike" kern="1200" cap="none" spc="0" normalizeH="0" baseline="0" noProof="0" dirty="0">
                    <a:ln>
                        <a:noFill />
                    </a:ln>
                    <a:solidFill>
                        <a:prstClr val="white" />
                    </a:solidFill>
                    <a:effectLst />
                    <a:uLnTx />
                    <a:uFillTx />
                    <a:latin typeface="PingFang SC Light" panose="020B0400000000000000" charset="-122" />
                    <a:ea typeface="PingFang SC Light" panose="020B0400000000000000" charset="-122" />
                    <a:cs typeface="PingFang SC Light" panose="020B0400000000000000" charset="-122" />
                </a:rPr>
                <a:t>感 谢 观 看</a:t>
            </a:r>
        </a:p>
    </p:txBody>
</p:sp>
`;

const parsed = tXml(xml, { simplify: 1 });
console.log('Parsed XML structure:', JSON.stringify(parsed, null, 2));

// 测试 getHorizontalAlign 函数
const spNode = parsed['p:sp'];
const textBodyNode = spNode['p:txBody'];
const pNode = textBodyNode['a:p'];

console.log('pNode:', JSON.stringify(pNode, null, 2));
console.log('textBodyNode:', JSON.stringify(textBodyNode, null, 2));

const warpObj = {
    slideLayoutTables: {
        idxTable: {}
    },
    slideMasterTextStyles: {},
    defaultTextStyle: {}
};

const result = getHorizontalAlign(pNode, textBodyNode, undefined, 'shape', '', warpObj);
console.log('Final alignment result:', result);
