/**
 * PPTX CSS 工具函数
 * 用于生成全局 CSS 样式
 */

interface StyleTableEntry {
    name: string;
    suffix?: string;
    text: string;
}

interface Settings {
    slideMode?: boolean;
    slideType?: string;
}

/**
 * 从样式表生成全局 CSS
 * @param styleTable - 样式表对象
 * @param settings - 设置对象
 * @param slideWidth - 幻灯片宽度
 * @returns 生成的 CSS 文本
 */
export function genGlobalCSS(styleTable: Record<string, StyleTableEntry>, settings: Settings, slideWidth: number): string {
    let cssText = "";
    //console.log("styleTable: ", styleTable)
    for (const key in styleTable) {
        let tagname = "";
        // if (settings.slideMode && settings.slideType == "revealjs") {
        //     tagname = "section";
        // } else {
        //     tagname = "div";
        // }
        //ADD suffix
        cssText += tagname + " ." + styleTable[key].name +
            ((styleTable[key].suffix) ? styleTable[key].suffix : "") +
            "{" + styleTable[key].text + "}\n"; //section > div
    }
    cssText += " .slide{margin-bottom: 5px;}\n";

    if (settings.slideMode && settings.slideType == "divs2slidesjs") {
        //divId
        //console.log("slideWidth: ", slideWidth)
        cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n";
    }
    return cssText;
}

// 向后兼容的导出对象
const PPTXCSSUtils = {
    genGlobalCSS
};

export { PPTXCSSUtils };