/**
 * pptx-table-utils.js
 * Utilities for handling table-related operations
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXTableUtils = {};

    /**
     * Get HTML bullet character
     * Convert special bullet characters to their HTML unicode equivalents
     * @param {String} typefaceNode - The font typeface (e.g., "Wingdings", "Wingdings 2")
     * @param {String} buChar - The bullet character
     * @returns {String} HTML entity for the bullet character
     */
    PPTXTableUtils.getHtmlBullet = function(typefaceNode, buChar) {
        //http://www.alanwood.net/demos/wingdings.html
        //not work for IE11
        switch (buChar) {
            case "§":
                return "&#9632;"; // U+25A0 | Black square
            case "q":
                return "&#10065;"; // U+2751 | Lower right shadowed white square
            case "v":
                return "&#10070;"; // U+2756 | Black diamond minus white X
            case "Ø":
                return "&#11162;"; // U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
            case "ü":
                return "&#10004;"; // U+2714 | Heavy check mark
            default:
                if (typefaceNode == "Wingdings" || typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3") {
                    var wingCharCode = window.TextUtils.getDingbatToUnicode(typefaceNode, buChar);
                    if (wingCharCode !== null) {
                        return "&#" + wingCharCode + ";";
                    }
                }
                return "&#" + (buChar.charCodeAt(0)) + ";";
        }
    };

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
