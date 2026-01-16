/**
 * pptx-slide-mode.js
 * 幻灯片模式管理模块
 * 负责初始化和退出幻灯片演示模式
 */

/**
 * 初始化幻灯片模式
 * 
 * @param {string} divId - 容器ID
 * @param {Object} settings - 配置对象
 * @param {Function} updateWrapperHeight - 更新包装器高度的函数
 */
function initSlideMode(divId, settings, updateWrapperHeight) {
    if (settings.slideType == "" || settings.slideType == "divs2slidesjs") {
        var slideElements = document.querySelectorAll("#" + divId + " .slide");
        var slidesHeight = slideElements.length > 0 ? slideElements[0].clientHeight : 0;
        // Hide all slide elements
        for (var i = 0; i < slideElements.length; i++) {
            slideElements[i].style.display = "none";
        }
        setTimeout(function () {
            var slideConf = settings.slideModeConfig;
            var loadingMsg = document.querySelector(".slides-loadnig-msg");
            if (loadingMsg) {
                loadingMsg.remove();
            }
            pptxjslideObj.init({
                divId: divId,
                slides: document.querySelectorAll("#" + divId + " .slide"),
                totalSlides: document.querySelectorAll("#" + divId + " .slide").length,
                slideCount: 1,
                prevSlide: -1,
                isInit: false,
                isSlideMode: true,
                isEnbleNextBtn: true,
                isEnblePrevBtn: false,
                isAutoSlideMode: false,
                isLoopMode: false,
                loopIntrval: null,
                first: slideConf.first,
                nav: slideConf.nav,
                showPlayPauseBtn: settings.showPlayPauseBtn,
                showFullscreenBtn: settings.showFullscreenBtn,
                navTxtColor: slideConf.navTxtColor,
                keyBoardShortCut: slideConf.keyBoardShortCut,
                showSlideNum: slideConf.showSlideNum,
                showTotalSlideNum: slideConf.showTotalSlideNum,
                autoSlide: slideConf.autoSlide,
                timeBetweenSlides: slideConf.autoSlide,
                randomAutoSlide: slideConf.randomAutoSlide,
                loop: slideConf.loop,
                background: slideConf.background,
                slctdBgClr: slideConf.background,
                transition: slideConf.transition,
                transitionTime: slideConf.transitionTime
            });

            updateWrapperHeight(divId, settings.slidesScale, true, "divs2slidesjs", 1);

        }, 1500);
    } else if (settings.slideType == "revealjs") {
        // Remove loading message first
        if (window.PPTXUIUtils && PPTXUIUtils.removeLoadingMessage) {
            PPTXUIUtils.removeLoadingMessage();
        }
        var revealjsPath = "";
        if (settings.revealjsPath != "") {
            revealjsPath = settings.revealjsPath;
        } else {
            revealjsPath = "./revealjs/reveal.js";
        }
        var script = document.createElement('script');
        script.src = revealjsPath;
        script.onload = function() {
            if (typeof Reveal !== 'undefined') {
                Reveal.initialize(settings.revealjsConfig);
            }
        };
        script.onerror = function() {
            console.error('Failed to load reveal.js script');
        };
        document.head.appendChild(script);
    }
}

/**
 * 退出幻灯片模式
 * 
 * @param {string} divId - 容器ID
 * @param {Object} settings - 配置对象
 * @param {Function} updateWrapperHeight - 更新包装器高度的函数
 */
function exitSlideMode(divId, settings, updateWrapperHeight) {
    // Show all slide elements
    var slideElements = document.querySelectorAll("#" + divId + " .slide");
    for (var i = 0; i < slideElements.length; i++) {
        slideElements[i].style.display = "block";
    }
    // If pptxjslideObj exists and has a destroy method, call it
    if (typeof window.pptxjslideObj !== 'undefined' && window.pptxjslideObj && typeof window.pptxjslideObj.destroy === 'function') {
        window.pptxjslideObj.destroy();
    }
    // Update wrapper height for normal mode
    updateWrapperHeight(divId, settings.slidesScale, false, settings.slideType, null);
}

export { initSlideMode, exitSlideMode };
