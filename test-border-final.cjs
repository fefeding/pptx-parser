// 最终测试边框解析逻辑
const fs = require('fs');
const path = require('path');

// 模拟PPTXXmlUtils
const PPTXXmlUtils = {
    getTextByPathList: (node, pathList) => {
        let current = node;
        for (const path of pathList) {
            if (current === undefined) return undefined;
            if (path === "attrs") {
                current = current.attrs;
            } else {
                current = current[path];
            }
        }
        return current;
    }
};

// 模拟warpObj - 完全模拟实际的XML结构
const warpObj = {
    themeContent: {
        "a:theme": {
            "a:themeElements": {
                "a:clrScheme": {
                    "a:accent2": {
                        "a:srgbClr": {
                            attrs: {
                                val: "ED7D31"
                            }
                        }
                    }
                },
                "a:fmtScheme": {
                    "a:lnStyleLst": {
                        "a:ln": {
                            attrs: {
                                w: "6350",
                                cap: "flat",
                                cmpd: "sng",
                                algn: "ctr"
                            },
                            "a:solidFill": {
                                "a:schemeClr": {
                                    attrs: {
                                        val: "phClr"
                                    }
                                }
                            },
                            "a:prstDash": {
                                attrs: {
                                    val: "solid"
                                }
                            }
                        }
                    }
                }
            }
        }
    }
};

// 模拟从slide12.xml中提取的节点
const node = {
    "p:spPr": {
        "a:xfrm": {
            "a:off": {
                attrs: {
                    x: "2912166",
                    y: "223630"
                }
            },
            "a:ext": {
                attrs: {
                    cx: "3781838",
                    cy: "646331"
                }
            }
        },
        "a:prstGeom": {
            attrs: {
                prst: "rect"
            },
            "a:avLst": {}
        }
    },
    "p:style": {
        "a:lnRef": {
            attrs: {
                idx: "0"
            },
            "a:schemeClr": {
                attrs: {
                    val: "accent2"
                }
            }
        },
        "a:fillRef": {
            attrs: {
                idx: "3"
            },
            "a:schemeClr": {
                attrs: {
                    val: "accent2"
                }
            }
        },
        "a:effectRef": {
            attrs: {
                idx: "3"
            },
            "a:schemeClr": {
                attrs: {
                    val: "accent2"
                }
            }
        }
    }
};

// 模拟getSchemeColorFromTheme
function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
    if (schemeClr === "a:phClr" && phClr) {
        return phClr;
    }
    if (schemeClr === "a:accent2" && warpObj) {
        const refNode = warpObj.themeContent["a:theme"]["a:themeElements"]["a:clrScheme"][schemeClr];
        return refNode ? refNode["a:srgbClr"].attrs.val : "000000";
    }
    return "000000";
}

// 模拟getSolidFill
function getSolidFill(node, clrMap, phClr, warpObj) {
    if (node === undefined) return undefined;
    
    if (node["a:schemeClr"] !== undefined) {
        const schemeClr = "a:" + node["a:schemeClr"].attrs.val;
        return getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj);
    }
    
    return undefined;
}

// 模拟getFillType
function getFillType(node) {
    if (node["a:solidFill"] !== undefined) {
        return "SOLID_FILL";
    }
    return "NO_FILL";
}

// 从style.js复制的getBorder函数
function getBorder(node, pNode, isSvgMode, type, warpObj) {
    let cssText = "";
    let lineNode;
    let lnRefNode;
    
    console.log('=== 开始测试getBorder函数 ===');
    
    // 从主题获取线条样式
    if (lineNode == undefined) {
        lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        console.log('1. lnRefNode:', lnRefNode);
        
        if (lnRefNode !== undefined) {
            let lnIdx = PPTXXmlUtils.getTextByPathList(lnRefNode, ["attrs", "idx"]);
            console.log('2. lnIdx:', lnIdx);
            
            // 检查lnStyleLst的结构
            const lnStyleLst = warpObj["themeContent"]["a:theme"]["a:themeElements"]["a:fmtScheme"]["a:lnStyleLst"]["a:ln"];
            console.log('3. lnStyleLst类型:', typeof lnStyleLst);
            console.log('4. lnStyleLst是否为数组:', Array.isArray(lnStyleLst));
            
            // 处理lnStyleLst可能是对象而不是数组的情况
            if (Array.isArray(lnStyleLst)) {
                lineNode = lnStyleLst[Number(lnIdx)];
            } else {
                // 如果是对象而不是数组，直接使用
                lineNode = lnStyleLst;
            }
            console.log('5. lineNode from theme:', lineNode);
        }
    }
    
    if (lineNode == undefined) {
        //is table
        cssText = "";
        lineNode = node;
    }
    
    let borderColor;
    let borderWidth = 0;
    let borderType = "solid";
    let strokeDasharray = "0";
    
    if (lineNode !== undefined) {
        // 计算边框宽度
        borderWidth = parseInt(PPTXXmlUtils.getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
        console.log('6. borderWidth:', borderWidth);
        
        if (isNaN(borderWidth) || borderWidth < 1) {
            cssText += (4/3) + "px ";
        } else {
            cssText += borderWidth + "px ";
        }
        
        // 获取边框类型
        borderType = PPTXXmlUtils.getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
        console.log('7. borderType:', borderType);
        
        strokeDasharray = "0";
        switch (borderType) {
            case "solid":
                cssText += "solid";
                strokeDasharray = "0";
                break;
            case "dash":
                cssText += "dashed";
                strokeDasharray = "5";
                break;
            case "dashDot":
                cssText += "dashed";
                strokeDasharray = "5, 5, 1, 5";
                break;
            case "dot":
                cssText += "dotted";
                strokeDasharray = "1, 5";
                break;
            default:
                cssText += "solid";
                strokeDasharray = "0";
        }
        
        // 计算边框颜色
        let fillTyp = getFillType(lineNode);
        console.log('8. fillTyp:', fillTyp);
        
        if (fillTyp === "SOLID_FILL") {
            // 获取lnRef中的颜色作为phClr参数
            if (!lnRefNode) {
                lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
            }
            let phClr = undefined;
            if (lnRefNode !== undefined) {
                phClr = getSolidFill(lnRefNode, undefined, undefined, warpObj);
                console.log('9. phClr from lnRef:', phClr);
            }
            borderColor = getSolidFill(lineNode["a:solidFill"], undefined, phClr, warpObj);
            console.log('10. borderColor:', borderColor);
        }
    }
    
    // 处理borderColor
    if (borderColor === undefined) {
        let lnRefNode = PPTXXmlUtils.getTextByPathList(node, ["p:style", "a:lnRef"]);
        if (lnRefNode !== undefined) {
            borderColor = getSolidFill(lnRefNode, undefined, undefined, warpObj);
            console.log('11. borderColor from lnRef:', borderColor);
        }
    }
    
    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        } else {
            borderColor = "hidden";
        }
    } else {
        if (borderColor && typeof borderColor === 'string') {
            if (!borderColor.startsWith('#') && !borderColor.startsWith('rgb') && !borderColor.startsWith('hsl') && borderColor !== 'none' && borderColor !== 'hidden') {
                borderColor = "#" + borderColor;
            }
        }
    }
    
    console.log('12. final borderColor:', borderColor);
    
    if (isSvgMode) {
        return { "color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray };
    } else {
        cssText += " " + borderColor + " ";
        return cssText + ";";
    }
}

// 测试
function runTest() {
    console.log('=== 测试边框解析逻辑 ===');
    console.log('模拟从slide12.xml中提取的节点结构');
    
    // 测试getBorder函数
    const result = getBorder(node, null, true, "shape", warpObj);
    console.log('=== 测试结果 ===');
    console.log('边框解析结果:', result);
    
    if (result.color === "#ED7D31" && result.width === 0.5 && result.type === "solid") {
        console.log('✅ 测试通过！边框颜色和样式都正确解析。');
        console.log('边框颜色:', result.color);
        console.log('边框宽度:', result.width, 'pt');
        console.log('边框类型:', result.type);
        console.log('\n修复成功！现在slide12.xml中的文本框边框应该可以正确显示了。');
    } else {
        console.log('❌ 测试失败！边框解析结果不正确。');
    }
}

// 运行测试
runTest();