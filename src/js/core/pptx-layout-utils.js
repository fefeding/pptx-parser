/**
 * pptx-layout-utils.js
 * Utilities for handling slide layout and master nodes
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXLayoutUtils = {};

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
        var lvlNode = window.PPTXUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
        if (lvlNode !== undefined) {
            lvl = parseInt(lvlNode) + 1;
        }
        if (idx !== undefined) {
            //slidelayout
            pPrNodeLaout = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:lstStyle", "a:lvl" + lvl + "pPr"]);
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", "a:pPr"]);
                if (pPrNodeLaout === undefined) {
                    pPrNodeLaout = window.PPTXUtils.getTextByPathList(warpObj["slideLayoutTables"]["idxTable"][idx], ["p:txBody", "a:p", (lvl - 1), "a:pPr"]);
                }
            }
        }
        if (type !== undefined) {
            //slidelayout
            var lvlStr = "a:lvl" + lvl + "pPr";
            if (pPrNodeLaout === undefined) {
                pPrNodeLaout = window.PPTXUtils.getTextByPathList(warpObj, ["slideLayoutTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
            //masterlayout
            if (type == "title" || type == "ctrTitle") {
                pPrNodeMaster = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:titleStyle", lvlStr]);
            } else if (type == "body" || type == "obj" || type == "subTitle") {
                pPrNodeMaster = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:bodyStyle", lvlStr]);
            } else if (type == "shape" || type == "diagram") {
                pPrNodeMaster = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTextStyles", "p:otherStyle", lvlStr]);
            } else if (type == "textBox") {
                pPrNodeMaster = window.PPTXUtils.getTextByPathList(warpObj, ["defaultTextStyle", lvlStr]);
            } else {
                pPrNodeMaster = window.PPTXUtils.getTextByPathList(warpObj, ["slideMasterTables", "typeTable", type, "p:txBody", "a:lstStyle", lvlStr]);
            }
        }
        return {
            "nodeLaout": pPrNodeLaout,
            "nodeMaster": pPrNodeMaster
        };
    };

    // Export to window
    window.PPTXLayoutUtils = PPTXLayoutUtils;

})();
