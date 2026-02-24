import tXml from '../src/js/core/tXml.js';

// 测试 XML 解析
const xml = `
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
`;

const parsed = tXml(xml, { simplify: 1 });
console.log('Parsed XML:', JSON.stringify(parsed, null, 2));

// 测试 getTextByPathList 函数
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

// 测试获取 algn 属性
const pNode = parsed['a:p'];
console.log('pNode:', JSON.stringify(pNode, null, 2));

const algn1 = getTextByPathList(pNode, ["a:pPr", "attrs", "algn"]);
console.log('algn1 (using ["a:pPr", "attrs", "algn"]):', algn1);

const algn2 = getTextByPathList(pNode, ["attrs", "algn"]);
console.log('algn2 (using ["attrs", "algn"]):', algn2);

const hasPPr = pNode["a:pPr"] !== undefined;
console.log('Has a:pPr:', hasPPr);

if (hasPPr) {
    const pPr = pNode["a:pPr"];
    console.log('pPr:', JSON.stringify(pPr, null, 2));
    console.log('pPr.attrs:', JSON.stringify(pPr.attrs, null, 2));
    console.log('pPr.attrs.algn:', pPr.attrs ? pPr.attrs.algn : undefined);
}
