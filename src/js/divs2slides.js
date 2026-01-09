/**
 * divs2slides.js
 * Ver : 1.3.2
 * update: 14/05/2018
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://github.com/meshesha/divs2slides
 * 
 * New: 
 *  - fixed fullscreen (fullscreen on div only insted all page)
 *  - converted to vanilla JavaScript (no jQuery dependency)
 */
(function(){
    
    var orginalMainDivWidth,
        orginalMainDivHeight,
        orginalSlidesWarpperScale,
        orginalSlideTop,
        orginalSlideLeft,
        orginalSlidesToolbarWidth,
        orginalSlidesToolbarTop;
    
    // Helper functions to replace jQuery methods
    var dom = {
        hide: function(element, duration) {
            if (duration && duration > 0) {
                element.style.transition = 'opacity ' + duration + 'ms';
                element.style.opacity = '0';
                setTimeout(() => { element.style.display = 'none'; }, duration);
            } else {
                element.style.display = 'none';
            }
        },
        show: function(element, duration) {
            element.style.display = '';
            if (duration && duration > 0) {
                element.style.transition = 'opacity ' + duration + 'ms';
                setTimeout(() => { element.style.opacity = '1'; }, 10);
            }
        },
        fadeOut: function(element, duration) {
            if (duration && duration > 0) {
                element.style.transition = 'opacity ' + duration + 'ms';
                element.style.opacity = '0';
                setTimeout(() => { element.style.display = 'none'; }, duration);
            } else {
                element.style.opacity = '0';
                element.style.display = 'none';
            }
        },
        fadeIn: function(element, duration) {
            element.style.display = '';
            element.style.opacity = '0';
            if (duration && duration > 0) {
                element.style.transition = 'opacity ' + duration + 'ms';
                setTimeout(() => { element.style.opacity = '1'; }, 10);
            } else {
                element.style.opacity = '1';
            }
        },
        slideUp: function(element, duration) {
            if (duration && duration > 0) {
                element.style.transition = 'height ' + duration + 'ms';
                element.style.height = '0';
                setTimeout(() => { element.style.display = 'none'; }, duration);
            } else {
                element.style.height = '0';
                element.style.display = 'none';
            }
        },
        slideDown: function(element, duration) {
            element.style.display = '';
            const originalHeight = element.scrollHeight + 'px';
            element.style.height = '0';
            if (duration && duration > 0) {
                element.style.transition = 'height ' + duration + 'ms';
                setTimeout(() => { element.style.height = originalHeight; }, 10);
            } else {
                element.style.height = originalHeight;
            }
        },
        css: function(element, property, value) {
            if (typeof property === 'string' && value !== undefined) {
                element.style[property] = value;
            } else if (typeof property === 'object') {
                for (var key in property) {
                    element.style[key] = property[key];
                }
            } else if (typeof property === 'string') {
                return window.getComputedStyle(element)[property];
            }
        },
        attr: function(element, attributes, value) {
            if (typeof attributes === 'string' && value !== undefined) {
                element.setAttribute(attributes, value);
            } else if (typeof attributes === 'object') {
                for (var key in attributes) {
                    element.setAttribute(key, attributes[key]);
                }
            } else if (typeof attributes === 'string') {
                return element.getAttribute(attributes);
            }
        },
        html: function(element, content) {
            if (content !== undefined) {
                element.innerHTML = content;
            } else {
                return element.innerHTML;
            }
        },
        prepend: function(parent, child) {
            parent.insertBefore(child, parent.firstChild);
        },
        wrapAll: function(elements, wrapperHtml) {
            var wrapper = document.createElement('div');
            wrapper.innerHTML = wrapperHtml;
            var wrapperElement = wrapper.firstChild;
            elements[0].parentNode.insertBefore(wrapperElement, elements[0]);
            for (var i = 0; i < elements.length; i++) {
                wrapperElement.appendChild(elements[i]);
            }
            return wrapperElement;
        },
        bind: function(element, event, handler) {
            element.addEventListener(event, handler);
        },
        on: function(element, event, handler) {
            element.addEventListener(event, handler);
        },
        isVisible: function(element) {
            return element.offsetParent !== null;
        },
        offset: function(element) {
            var rect = element.getBoundingClientRect();
            return {
                top: rect.top + window.pageYOffset,
                left: rect.left + window.pageXOffset
            };
        },
        find: function(parent, selector) {
            return parent.querySelectorAll(selector);
        },
        children: function(parent) {
            return parent.children;
        },
        length: function(collection) {
            return collection.length;
        }
    };
    
    var pptxjslideObj = {
        init: function(options){
            pptxjslideObj.data = options;
            var data = pptxjslideObj.data;
            var divId = data.divId;
            var container = document.getElementById(divId);
            var slides = data.slides || container.querySelectorAll('.slide');
            var isInit = data.isInit;
            
            // Hide all slides
            for (var i = 0; i < slides.length; i++) {
                dom.hide(slides[i]);
            }
            
            if(data.slctdBgClr != false){
                var preBgClr = dom.css(document.body, 'background-color');
                data.prevBgColor = preBgClr;
                dom.css(document.body, 'background-color', data.slctdBgClr);
            }
            
            if (data.nav && !isInit){
                data.isInit = true;
                // Create navigators 
                var toolbar = document.createElement('div');
                dom.attr(toolbar, {
                    'class': 'slides-toolbar',
                    'style': 'width: 90%; padding: 10px; text-align: center;font-size:18px; color: ' + data.navTxtColor + ';'
                });
                dom.prepend(container, toolbar);
                
                // Next button
                var nextBtn = document.createElement('img');
                dom.attr(nextBtn, {
                    'id': 'slides-next',
                    'class': 'slides-nav',
                    'alt': 'Next Slide',
                    'style': 'float: right;cursor: pointer;opacity: 0.7;'
                });
                dom.bind(nextBtn, 'click', pptxjslideObj.nextSlide);
                dom.prepend(toolbar, nextBtn);
                if(data.showTotalSlideNum){
                    var totalSpan = document.createElement('span');
                    dom.attr(totalSpan, 'id', 'slides-total-slides-num');
                    dom.html(totalSpan, data.totalSlides);
                    dom.prepend(toolbar, totalSpan);
                }
                
                if(data.showSlideNum && data.showTotalSlideNum){
                    var separatorSpan = document.createElement('span');
                    dom.attr(separatorSpan, 'id', 'slides-slides-num-separator');
                    dom.html(separatorSpan, ' / ');
                    dom.prepend(toolbar, separatorSpan);
                }
                
                if(data.showSlideNum){
                    var slideNumSpan = document.createElement('span');
                    dom.attr(slideNumSpan, 'id', 'slides-slide-num');
                    dom.html(slideNumSpan, data.slideCount);
                    dom.prepend(toolbar, slideNumSpan);
                }
                if(data.showFullscreenBtn){
                    var fullscreenBtn = document.createElement('img');
                    dom.attr(fullscreenBtn, {
                        'id': 'slides-full-screen',
                        'class': 'slides-nav-play',
                        'alt': 'fullscreen Slide',
                        'style': 'float: left;cursor: pointer;opacity: 0.7; padding: 0 10px 0 10px'
                    });
                    dom.bind(fullscreenBtn, 'click', function() { 
                        pptxjslideObj.fullscreen();
                    });
                    dom.prepend(toolbar, fullscreenBtn);
                }
                if(data.showPlayPauseBtn){
                    var playPauseBtn = document.createElement('img');
                    dom.attr(playPauseBtn, {
                        'id': 'slides-play-pause',
                        'class': 'slides-nav-play',
                        'alt': 'Play/Pause Slide',
                        'style': 'float: left;cursor: pointer;opacity: 0.7;  padding: 0 10px 0 10px'
                    });
                    dom.html(playPauseBtn, '<span style="font-size:80%;">&#x23ef;</span>');
                    dom.bind(playPauseBtn, 'click', function() { 
                        if(data.isSlideMode){
                            pptxjslideObj.startAutoSlide();
                        }
                    });
                    dom.prepend(toolbar, playPauseBtn);
                }
                // Previous button
                var prevBtn = document.createElement('img');
                dom.attr(prevBtn, {
                    'id': 'slides-prev',
                    'class': 'slides-nav',
                    'alt': 'Prev. Slide',
                     'style': 'float: left;cursor: pointer; opacity: 0.7;'
                });
                dom.bind(prevBtn, 'click', pptxjslideObj.prevSlide);
                dom.prepend(toolbar, prevBtn);

                // Mouseover/mouseout events for navigation buttons
                var navButtons = toolbar.querySelectorAll('.slides-nav, .slides-nav-play');
                for (var j = 0; j < navButtons.length; j++) {
                    (function(btn) {
                        dom.bind(btn, 'mouseover', function(){
                            dom.css(btn, 'opacity', '1');
                        });
                        dom.bind(btn, 'mouseout', function(){
                            dom.css(btn, 'opacity', '0.7');
                        });
                    })(navButtons[j]);
                }
                if(data.slideCount == 1){
                    dom.css(document.getElementById('slides-prev'), 'display', 'none');
                }else if(data.slideCount == data.totalSlides){
                    dom.css(document.getElementById('slides-next'), 'display', 'none');
                }else{
                    dom.css(document.getElementById('slides-next'), 'display', 'block');
                }
            }else{
                var toolbar = container.querySelector('.slides-toolbar');
                if (toolbar) {
                    dom.css(toolbar, 'display', 'block');
                }
                data.isEnbleNextBtn = true;
                data.isEnblePrevBtn = true;
            }
            
            if(document.getElementById('all_slides_warpper') === null){
                dom.wrapAll(slides, '<div id="all_slides_warpper"></div>');
            }
            // Go to first slide
            pptxjslideObj.gotoSlide(1);
        },
        nextSlide: function(){
            var data = pptxjslideObj.data;
            var isLoop = data.isLoop;
            var isAutoMode = data.isAutoSlideMode;
            if (data.slideCount < data.totalSlides){
                    pptxjslideObj.gotoSlide(data.slideCount+1);
                    if(!isAutoMode) dom.css(document.getElementById('slides-next'), 'display', 'block');
            }else{
                if(isLoop){
                    pptxjslideObj.gotoSlide(1);
                }else{
                    if(!isAutoMode) dom.css(document.getElementById('slides-next'), 'display', 'none');
                }
            }
            if(!isAutoMode){
                if(data.slideCount > 1){
                    dom.css(document.getElementById('slides-prev'), 'display', 'block');
                }else{
                    dom.css(document.getElementById('slides-prev'), 'display', 'none');
                }
                if(data.slideCount == data.totalSlides && !isLoop){
                    dom.css(document.getElementById('slides-next'), 'display', 'none');
                }
            }
        },
        prevSlide: function(){
            var data = pptxjslideObj.data;
            var isAutoMode = data.isAutoSlideMode;
            if (data.slideCount > 1){
                pptxjslideObj.gotoSlide(data.slideCount-1);
            }
            if(!isAutoMode){
                if(data.slideCount == 1){
                    dom.css(document.getElementById('slides-prev'), 'display', 'none');
                }else{
                    dom.css(document.getElementById('slides-prev'), 'display', 'block');
                }
                dom.css(document.getElementById('slides-next'), 'display', 'block');
            }
        },
        gotoSlide: function(idx){
            var index = idx - 1;
            var data = pptxjslideObj.data;
            var slides = data.slides || document.getElementById(data.divId).querySelectorAll('.slide');
            var prevSlidNum = data.prevSlide;
            var transType = data.transition; /*"slid","fade","default" */
            if(transType=="random"){
                var tType = ["","default","fade","slid"];
                var randomNum = Math.floor(Math.random() * 3) + 1; //random number between 1 to 3
                transType = tType[randomNum];
            }
            var transTime = 1000*(data.transitionTime);
            if (slides[index]){
                var nextSlide = slides[index];
                if (prevSlidNum >= 0 && slides[prevSlidNum] && dom.isVisible(slides[prevSlidNum])){
                    if(transType=="default"){
                        dom.hide(slides[prevSlidNum], transTime);
                    }else if(transType=="fade"){
                        dom.fadeOut(slides[prevSlidNum], transTime);
                    }else if(transType=="slid"){
                        dom.slideUp(slides[prevSlidNum], transTime);
                    }
                }
                if(transType=="default"){
                    dom.show(nextSlide, transTime); 
                }else if(transType=="fade"){
                    dom.fadeIn(nextSlide, transTime);
                }else if(transType=="slid"){
                    dom.slideDown(nextSlide, transTime);
                }
                data.prevSlide = index;
                pptxjslideObj.data.slideCount = idx;
                var slideNumElement = document.getElementById('slides-slide-num');
                if (slideNumElement) {
                    dom.html(slideNumElement, idx);
                }
            }
            return this;
        },
        keyDown: function(event){
            event.preventDefault();
            var key = event.keyCode;
            //console.log(key);
            var data = pptxjslideObj.data;
            switch(key){
                case(37): // Left arrow
                case(8): // Backspace
                    if(data.isSlideMode && data.isEnblePrevBtn){
                        pptxjslideObj.prevSlide();
                    }
                    break;
                case(39): // Right arrow
                case(32): // Space 
                case(13): // Enter 
                    if(data.isSlideMode  && data.isEnbleNextBtn){
                        pptxjslideObj.nextSlide();
                    }
                    break; 
                case(46): //Delete
                    //if in auto mode , stop auto mode TODO
                    if(data.isSlideMode){
                        var div_id = data.divId;
                        var container = document.getElementById(div_id);
                        var slides = container.querySelectorAll('.slide');
                        for (var i = 0; i < slides.length; i++) {
                            dom.hide(slides[i]);
                        }
                        pptxjslideObj.gotoSlide(1);               //bugFix to ver. 1.2.1
                    }
                    break;
                case(27): //Esc
                    if(data.isSlideMode){
                        pptxjslideObj.closeSileMode();
                        data.isSlideMode = false;
                    }
                    break;
                case(116): //F5
                    if(!data.isSlideMode){
                        pptxjslideObj.startSlideMode();
                        data.isSlideMode = true;
                        if(data.isAutoSlideMode || data.isLoopMode){
                            clearInterval(data.loopIntrval);
                            data.isAutoSlideMode = false;
                            data.isLoopMode = false;
                        }
                        
                    }
                    break;
                case(113): // F2
                    if(data.isSlideMode){
                        pptxjslideObj.fullscreen();
                    }
                    break;
                case(119): // F8
                    if(data.isSlideMode){
                        pptxjslideObj.startAutoSlide();
                        //TODO : ADD indication that it is in auto slide mode
                    }
                break;
            }
            return true;
        },
        startSlideMode: function(){
            pptxjslideObj.init();
        },
        closeSileMode: function(){
            var data = pptxjslideObj.data;
            data.isSlideMode = false;
            var div_id= data.divId;
            var container = document.getElementById(div_id);
            var toolbar = container.querySelector('.slides-toolbar');
            if (toolbar) {
                dom.css(toolbar, 'display', 'none');
            }
            var slides = container.querySelectorAll('.slide');
            for (var i = 0; i < slides.length; i++) {
                dom.show(slides[i]);
            }
            dom.css(document.body, 'background-color', pptxjslideObj.data.prevBgColor);
            if(data.isLoopMode){
                clearInterval(data.loopIntrval);
                data.isLoopMode = false;
            }
            pptxjslideObj.exitFullscreenMod();
        },
        startAutoSlide: function(){
            var data = pptxjslideObj.data;
            var isAutoSlideOption = data.timeBetweenSlides
            var isAutoSlideMode = data.isAutoSlideMode;
            if(!isAutoSlideMode && isAutoSlideOption !== false){
                data.isAutoSlideMode = true;
                //var isLoopOption = data.isLoop;
                var isStrtLoop =  data.isLoopMode;
                //hide and disable next and prev btn
                if(data.nav){
                    var div_Id = data.divId;
                    var container = document.getElementById(div_Id);
                    var navButtons = container.querySelectorAll('.slides-toolbar .slides-nav');
                    for (var i = 0; i < navButtons.length; i++) {
                        dom.hide(navButtons[i]);
                    }
                }
                data.isEnbleNextBtn = false;
                data.isEnblePrevBtn = false;
                ///////////////////////////////
                
                var t = isAutoSlideOption + data.transitionTime;
                
                var slideNums = data.totalSlides;
                var isRandomSlide = data.randomAutoSlide;
                
                if(!isStrtLoop){
                    var timeBtweenSlides = t*1000; //milisecons
                    data.isLoopMode = true;
                    data.loopIntrval = setInterval(function(){
                        if(isRandomSlide){
                            var randomSlideNum = Math.floor(Math.random() * slideNums) + 1;
                            pptxjslideObj.gotoSlide(randomSlideNum);
                        }else{
                            pptxjslideObj.nextSlide();
                        }
                    }, timeBtweenSlides);
                }else{
                    clearInterval(data.loopIntrval);
                    data.isLoopMode = false;                
                }
            }else{
                clearInterval(data.loopIntrval);
                data.isAutoSlideMode = false;
                data.isLoopMode = false;
                //show and enable next and prev btn
                if(data.nav){
                    var div_Id = data.divId;
                    var container = document.getElementById(div_Id);
                    var navButtons = container.querySelectorAll('.slides-toolbar .slides-nav');
                    for (var i = 0; i < navButtons.length; i++) {
                        dom.show(navButtons[i]);
                    }
                }
                data.isEnbleNextBtn = true;
                data.isEnblePrevBtn = true;    
            }
        },
        fullscreen: function(){
            if (!document.fullscreenElement &&    
                !document.mozFullScreenElement && !document.webkitFullscreenElement && !document.msFullscreenElement ) {  // current working methods
                var data = pptxjslideObj.data;
                var div_Id = data.divId;
                var container = document.getElementById(div_Id);
                if (document.documentElement.requestFullscreen) {
                    container.requestFullscreen();
                } else if (document.documentElement.msRequestFullscreen) {
                    container.msRequestFullscreen();
                } else if (document.documentElement.mozRequestFullScreen) {
                    container.mozRequestFullScreen();
                } else if (document.documentElement.webkitRequestFullscreen) {
                    container.webkitRequestFullscreen(Element.ALLOW_KEYBOARD_INPUT);
                }
                var winWidth = window.innerWidth;
                var winHeight = window.innerHeight;
                //Need to save:
                orginalMainDivWidth = dom.css(container, 'width');
                orginalMainDivHeight = dom.css(container, 'height');
                var m = dom.css(container.querySelector('#all_slides_warpper'), 'transform');
                orginalSlidesWarpperScale = m.substring(m.indexOf('(') + 1, m.indexOf(')')).split(",")
                var slideElement = container.querySelector('#all_slides_warpper .slide');
                var slideOffset = dom.offset(slideElement);
                orginalSlideTop = slideOffset.top;
                orginalSlideLeft = slideOffset.left;
                var toolbar = container.querySelector('.slides-toolbar');
                orginalSlidesToolbarWidth = dom.css(toolbar, 'width');
                var toolbarOffset = dom.offset(toolbar);
                orginalSlidesToolbarTop = toolbarOffset.top;

                dom.attr(container, 'style', "width: " + (winWidth - 10) + "px; height: " + (winHeight - 10) + "px;");

                dom.css(container.querySelector('#all_slides_warpper'), {
                    "transform":"scale(1)"
                });

                var slideWidth = dom.css(slideElement, 'width');
                var sildeHeight = dom.css(slideElement, 'height');
                dom.css(slideElement, {
                    "top": ((winHeight - parseInt(sildeHeight))/2) + "px",
                    "left": ((winWidth - parseInt(slideWidth))/2) + "px"
                });

                if(data.nav && toolbar){
                    dom.css(toolbar, {
                        "width": "99%",
                        "top": "20px"
                    });
                }
            } else {
                if (document.exitFullscreen) {
                    document.exitFullscreen();
                } else if (document.msExitFullscreen) {
                    document.msExitFullscreen();
                } else if (document.mozCancelFullScreen) {
                    document.mozCancelFullScreen();
                } else if (document.webkitExitFullscreen) {
                    document.webkitExitFullscreen();
                }

                pptxjslideObj.exitFullscreenMod();
            }
            
        },
        exitFullscreenMod: function(){
            var data = pptxjslideObj.data;
            var div_Id = data.divId;
            var container = document.getElementById(div_Id);
            //saved:
            /*
            orginalMainDivWidth
            orginalMainDivHeight
            orginalSlidesWarpperScale
            orginalSlideTop
            orginalSlideLeft
            orginalSlidesToolbarWidth
            orginalSlidesToolbarTop
            */
            dom.attr(container, {
                style: "width: " + orginalMainDivWidth + "px; height: " + orginalMainDivHeight + "px;"
            });
            console.log(orginalSlidesWarpperScale[0])
            dom.css(container.querySelector('#all_slides_warpper'), {
                "transform":"scale(" + orginalSlidesWarpperScale[0] + ")"
            });

            var slideElement = container.querySelector('#all_slides_warpper .slide');
            dom.css(slideElement, {
                "top": "0px", /**orginalSlideTop +  */
                "left": "0px" /**orginalSlideLeft +  */
            });

            if(data.nav){
                var toolbar = container.querySelector('.slides-toolbar');
                if (toolbar) {
                    dom.css(toolbar, {
                        "width": orginalSlidesToolbarWidth + "px",
                        "top": orginalSlidesToolbarTop + "px"
                    });
                }
            }
        }

    };
    // Expose to global for backward compatibility
    window.pptxjslideObj = pptxjslideObj;
})();
