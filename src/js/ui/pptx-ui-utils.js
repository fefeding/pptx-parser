/**
 * pptx-ui-utils.js
 * Utilities for UI-related operations like progress bars and slide wrappers
 * Extracted from pptxjs.js for better code organization
 */

(function() {
    'use strict';

    var PPTXUIUtils = {};

    /**
     * Update the loading progress bar
     * @param {Number} percent - The progress percentage (0-100)
     */
    PPTXUIUtils.updateProgressBar = function(percent) {
        var progressBarElement = document.querySelector(".slides-loading-progress-bar");
        if (progressBarElement) {
            progressBarElement.style.width = percent + "%";
            progressBarElement.innerHTML = "<span style='text-align: center;'>Loading...(" + percent + "%)</span>";
        }
    };

    /**
     * Create or get the slides wrapper element
     * @param {String} divId - The container div ID
     * @returns {HTMLElement} The slides wrapper element
     */
    PPTXUIUtils.getSlidesWrapper = function(divId) {
        var wrapper = document.getElementById("all_slides_warpper");
        if (wrapper === null) {
            var slides = document.querySelectorAll("#" + divId + " .slide");
            if (slides.length > 0) {
                wrapper = document.createElement("div");
                wrapper.id = "all_slides_warpper";
                wrapper.className = "slides";
                var firstSlide = slides[0];
                var parent = firstSlide.parentNode;
                parent.insertBefore(wrapper, firstSlide);
                
                // Move each slide into the wrapper
                for (var i = 0; i < slides.length; i++) {
                    wrapper.appendChild(slides[i]);
                }
            }
        }
        return wrapper;
    };

    /**
     * Update the slides wrapper height based on scale
     * @param {String} divId - The container div ID
     * @param {String} slidesScale - The scale percentage (e.g., "100", "50")
     * @param {Boolean} isSlideMode - Whether in slide mode
     * @param {String} slideType - The slide type ("divs2slidesjs", "revealjs", etc.)
     * @param {Number} numOfSlides - Number of slides (for slide mode)
     */
    PPTXUIUtils.updateWrapperHeight = function(divId, slidesScale, isSlideMode, slideType, numOfSlides) {
        var sScale = slidesScale;
        var trnsfrmScl = "";
        
        if (sScale != "") {
            var numsScale = parseInt(sScale);
            var scaleVal = numsScale / 100;
            if (isSlideMode && slideType != "revealjs") {
                trnsfrmScl = 'transform:scale(' + scaleVal + '); transform-origin:top';
            }
        }

        var slideElements = document.querySelectorAll("#" + divId + " .slide");
        var slidesHeight = slideElements.length > 0 ? slideElements[0].clientHeight : 0;
        var sScaleVal = (sScale != "") ? scaleVal : 1;
        
        var allSlidesWrapper = document.getElementById("all_slides_warpper");
        if (allSlidesWrapper) {
            var num = isSlideMode ? numOfSlides : slideElements.length;
            allSlidesWrapper.style.cssText = trnsfrmScl + ";height: " + (num * slidesHeight * sScaleVal) + "px";
        }
    };

    /**
     * Remove the loading message
     */
    PPTXUIUtils.removeLoadingMessage = function() {
        var loadingMsg = document.querySelector(".slides-loadnig-msg");
        if (loadingMsg) {
            loadingMsg.remove();
        }
    };

    /**
     * Add reveal class to container for reveal.js integration
     * @param {String} divId - The container div ID
     */
    PPTXUIUtils.addRevealClass = function(divId) {
        var revealElem = document.getElementById(divId);
        if (revealElem) {
            revealElem.classList.add("reveal");
        }
    };

    // Export to window
    window.PPTXUIUtils = PPTXUIUtils;

})();
