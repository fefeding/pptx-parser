
const PPTXLayoutUtils = {};

    /**
 * Get layout and master nodes
 * @param {Object} node - The paragraph node
 * @param {Number} idx - The index
 * @param {String} type - The type
 * @param {Object} warpObj - The warp object
 * @returns {Object} Object with nodeLaout and nodeMaster properties
 */
PPTXLayoutUtils.getLayoutAndMasterNode = function(node, idx, type, warpObj) {
    var pPrNodeLaout, pPrNodeMaster;
    var pPrNode = node["a:pPr"];
    //lvl
    var lvl = 1;
    var lvlNode = PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
    if (lvlNode !== undefined) {
        lvl = parseInt(lvlNode) + 1;
    }
    if (idx !== undefined) {
        //slidelayout
        pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
            }
        }
    }
    if (type !== undefined) {
        //slidelayout
        var lvlStr = "a:lvl" + lvl + "pPr";
        if (pPrNodeLaout === undefined) {
            pPrNodeLaout = PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
        //masterlayout
        if (type == "title" || type == "ctrTitle") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
        } else if (type == "body" || type == "obj" || type == "subTitle") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
        } else if (type == "shape" || type == "diagram") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
        } else if (type == "textBox") {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
        } else {
            pPrNodeMaster = PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
        }
    }
    return {
        "nodeLaout": pPrNodeLaout,
        "nodeMaster": pPrNodeMaster
    };
};

    // Export to window

export { PPTXLayoutUtils };

// Also export to global scope for backward compatibility
// window.PPTXLayoutUtils = PPTXLayoutUtils; // Removed for ES modules
