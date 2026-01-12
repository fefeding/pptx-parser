/**
 * pptx-css-utils.js
 * Utilities for generating CSS
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXCSSUtils = {};

    /**
     * Generate global CSS from style table
     * @param {Object} styleTable - The style table object
     * @param {Object} settings - The settings object
     * @param {Number} slideWidth - The slide width
     * @returns {String} Generated CSS text
     */
    PPTXCSSUtils.genGlobalCSS = function(styleTable, settings, slideWidth) {
        var cssText = "";
        //console.log("styleTable: ", styleTable)
        for (var key in styleTable) {
            var tagname = "";
            // if (settings.slideMode && settings.slideType == "revealjs") {
            //     tagname = "section";
            // } else {
            //     tagname = "div";
            // }
            //ADD suffix
            cssText += tagname + " ." + styleTable[key]["name"] +
                ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
                "{" + styleTable[key]["text"] + "}\n"; //section > div
        }
        //cssText += " .slide{margin-bottom: 5px;}\n"; // TODO

        if (settings.slideMode && settings.slideType == "divs2slidesjs") {
            //divId
            //console.log("slideWidth: ", slideWidth)
            cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n"; // TODO
        }
        return cssText;
    };

    // Export to window
    window.PPTXCSSUtils = PPTXCSSUtils;

})();
