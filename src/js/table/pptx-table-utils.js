/**
 * pptx-table-utils.js
 * Utilities for handling table-related operations
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXTableUtils = {};

    /**
     * Get table borders style
     * @param {Object} node - The table borders node
     * @param {Object} warpObj - The warp object
     * @returns {String} CSS border style
     */
    PPTXTableUtils.getTableBorders = function(node, warpObj) {
        var borderStyle = "";
        if (node["a:bottom"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:bottom"]["a:ln"]
                }
            }
            var borders = window.PPTXShapeFillsUtils.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-bottom");
        }
        if (node["a:top"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:top"]["a:ln"]
                }
            }
            var borders = window.PPTXShapeFillsUtils.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-top");
        }
        if (node["a:right"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:right"]["a:ln"]
                }
            }
            var borders = window.PPTXShapeFillsUtils.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-right");
        }
        if (node["a:left"] !== undefined) {
            var obj = {
                "p:spPr": {
                    "a:ln": node["a:left"]["a:ln"]
                }
            }
            var borders = window.PPTXShapeFillsUtils.getBorder(obj, undefined, false, "shape", warpObj);
            borderStyle += borders.replace("border", "border-left");
        }

        return borderStyle;
    };

    // Export to window
    window.PPTXTableUtils = PPTXTableUtils;

})();
