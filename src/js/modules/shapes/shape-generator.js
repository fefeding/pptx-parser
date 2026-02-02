/**
 * 形状生成器模块
 * 处理各种形状的生成逻辑
 * 这个模块包含了大量的形状处理函数
 * 注意：由于原始文件过大（14000+行），完整迁移需要更多时间
 * 当前版本提供框架和关键函数签名
 */

/**
 * genShape - 生成形状HTML
 * @param {Object} node - 节点对象
 * @param {Object} pNode - 父节点对象
 * @param {Object} slideLayoutSpNode - Slide布局节点
 * @param {Object} slideMasterSpNode - Slide主节点
 * @param {string} id - ID
 * @param {string} name - 名称
 * @param {string} idx - 索引
 * @param {string} type - 类型
 * @param {number} order - 顺序
 * @param {Object} warpObj - 包装对象
 * @param {boolean} isUserDrawnBg - 是否用户绘制背景
 * @param {string} sType - 源类型
 * @param {string} source - 源
 * @returns {string} 形状HTML字符串
 */

var ShapeGenerator = (function() {
    function genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, isUserDrawnBg, sType, source) {
    // TODO: 完整实现形状生成逻辑
    // 这是从原始文件pptxjs.js迁移的核心函数
    // 需要处理各种形状类型：矩形、圆形、线条、自定义形状等

    // 获取变换信息
    const xfrmList = ["p:spPr", "a:xfrm"];
    const slideXfrmNode = getTextByPathList(node, xfrmList);
    const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
    const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);

    // 获取形状类型
    const shapType = getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    const custShapType = getTextByPathList(node, ["p:spPr", "a:custGeom"]);

    // 翻转处理
    let isFlipV = getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1";
    let isFlipH = getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1";
    let flip = "";
    if (isFlipH && !isFlipV) {
        flip = " scale(-1,1)";
    } else if (!isFlipH && isFlipV) {
        flip = " scale(1,-1)";
    } else if (isFlipH && isFlipV) {
        flip = " scale(-1,-1)";
    }

    // TODO: 继续实现完整逻辑
    // 包括位置计算、尺寸计算、边框、填充、文本等

    return `<!-- Shape ${id} - ${name} (Type: ${type}) -->`;
}

/**
 * processSpNode - 处理形状节点
 * @param {Object} node - 节点对象
 * @param {Object} pNode - 父节点对象
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 源
 * @param {string} sType - 源类型
 * @returns {string} HTML字符串
 */
    function processSpNode(node, pNode, warpObj, source, sType) {
    const id = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "id"]);
    const name = getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]);
    const idx = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "idx"]);
    const type = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    const order = getTextByPathList(node, ["attrs", "order"]);

    let slideLayoutSpNode;
    let slideMasterSpNode;

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

    // TODO: 处理文本框等特殊类型

    return genShape(node, pNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj, undefined, sType, source);
}

/**
 * processCxnSpNode - 处理连接形状节点
 * @param {Object} node - 节点对象
 * @param {Object} pNode - 父节点对象
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 源
 * @param {string} sType - 源类型
 * @returns {string} HTML字符串
 */
    function processCxnSpNode(node, pNode, warpObj, source, sType) {
    const id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
    const name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
    const idx = node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    const type = node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    const order = node["attrs"]["order"];

    return genShape(node, pNode, undefined, undefined, id, name, idx, type, order, warpObj, undefined, sType, source);
}

/**
 * processPicNode - 处理图片节点
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 源
 * @param {string} sType - 源类型
 * @returns {string} HTML字符串
 */
    function processPicNode(node, warpObj, source, sType) {
    // TODO: 实现图片节点处理逻辑
    // 包括读取图片数据、计算位置、设置样式等
    return `<!-- Picture Node -->`;
}

/**
 * processGraphicFrameNode - 处理图形框架节点（图表、表格等）
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 源
 * @param {string} sType - 源类型
 * @returns {string} HTML字符串
 */
    function processGraphicFrameNode(node, warpObj, source, sType) {
    // TODO: 实现图形框架节点处理逻辑
    // 包括图表、表格、SmartArt等
    return `<!-- GraphicFrame Node -->`;
}

/**
 * processGroupSpNode - 处理组合形状节点
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @param {string} source - 源
 * @returns {string} HTML字符串
 */
    function processGroupSpNode(node, warpObj, source) {
    // TODO: 实现组合形状节点处理逻辑
    // 包括位置、旋转、缩放等变换
    return `<!-- Group Shape Node -->`;
}


    return {
        genShape: genShape,
        processSpNode: processSpNode,
        processCxnSpNode: processCxnSpNode,
        processPicNode: processPicNode,
        processGraphicFrameNode: processGraphicFrameNode,
        processGroupSpNode: processGroupSpNode
    };
})();