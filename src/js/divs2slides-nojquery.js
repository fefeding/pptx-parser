/**
 * divs2slides-nojquery.js
 * No jQuery version of divs2slides.js
 * Converted from jQuery plugin to native JavaScript
 */

// Immediately invoked function without jQuery dependency
(function() {
    
    // Main divs2slides function
    window.divs2slides = function(container, options) {
        var settings = Object.assign({}, getDefaults(), options);
        
        // Initialize variables
        var currentSlide = settings.first - 1;
        var totalSlides = 0;
        var slides = [];
        var isPlaying = false;
        var playInterval;
        
        // Get all slide elements
        if (container.querySelectorAll) {
            slides = container.querySelectorAll('.slide');
            totalSlides = slides.length;
        }
        
        if (totalSlides === 0) {
            console.warn('No slides found');
            return;
        }
        
        // Create navigation if enabled
        if (settings.nav) {
            createNavigation();
        }
        
        // Show initial slide
        showSlide(currentSlide);
        
        // Auto slide functionality
        if (settings.autoSlide && settings.autoSlide !== false) {
            var interval = parseInt(settings.autoSlide) * 1000;
            startAutoSlide(interval);
        }
        
        function getDefaults() {
            return {
                first: 1,
                nav: true,
                showPlayPauseBtn: true,
                navTxtColor: "black",
                keyBoardShortCut: true,
                showSlideNum: true,
                showTotalSlideNum: true,
                autoSlide: false,
                randomAutoSlide: false,
                loop: false,
                background: false,
                transition: "default",
                transitionTime: 1
            };
        }
        
        function createNavigation() {
            var navDiv = document.createElement('div');
            navDiv.className = 'slides-nav';
            navDiv.style.color = settings.navTxtColor;
            
            // Previous button
            var prevBtn = document.createElement('button');
            prevBtn.innerHTML = settings.navPrevTxt || '&lt;';
            prevBtn.className = 'nav-prev';
            prevBtn.addEventListener('click', function() {
                previousSlide();
            });
            navDiv.appendChild(prevBtn);
            
            // Slide counter
            if (settings.showSlideNum || settings.showTotalSlideNum) {
                var counter = document.createElement('span');
                counter.className = 'slide-counter';
                counter.innerHTML = ' ';
                if (settings.showSlideNum && settings.showTotalSlideNum) {
                    counter.innerHTML = (currentSlide + 1) + ' / ' + totalSlides;
                } else if (settings.showSlideNum) {
                    counter.innerHTML = (currentSlide + 1);
                } else if (settings.showTotalSlideNum) {
                    counter.innerHTML = totalSlides;
                }
                navDiv.appendChild(counter);
            }
            
            // Next button
            var nextBtn = document.createElement('button');
            nextBtn.innerHTML = settings.navNextTxt || '&gt;';
            nextBtn.className = 'nav-next';
            nextBtn.addEventListener('click', function() {
                nextSlide();
            });
            navDiv.appendChild(nextBtn);
            
            // Play/Pause button
            if (settings.showPlayPauseBtn) {
                var playPauseBtn = document.createElement('button');
                playPauseBtn.innerHTML = '⏸';
                playPauseBtn.className = 'play-pause';
                playPauseBtn.addEventListener('click', function() {
                    toggleAutoSlide();
                });
                navDiv.appendChild(playPauseBtn);
            }
            
            container.appendChild(navDiv);
        }
        
        function showSlide(index) {
            // Hide all slides
            for (var i = 0; i < slides.length; i++) {
                if (slides[i].style) {
                    slides[i].style.display = 'none';
                }
            }
            
            // Show current slide
            if (slides[index] && slides[index].style) {
                slides[index].style.display = 'block';
                currentSlide = index;
                
                // Update counter if exists
                var counter = document.querySelector('.slide-counter');
                if (counter) {
                    if (settings.showSlideNum && settings.showTotalSlideNum) {
                        counter.innerHTML = (currentSlide + 1) + ' / ' + totalSlides;
                    } else if (settings.showSlideNum) {
                        counter.innerHTML = (currentSlide + 1);
                    }
                }
            }
        }
        
        function nextSlide() {
            var next = currentSlide + 1;
            if (next >= totalSlides) {
                if (settings.loop) {
                    next = 0;
                } else {
                    next = totalSlides - 1;
                }
            }
            showSlide(next);
        }
        
        function previousSlide() {
            var prev = currentSlide - 1;
            if (prev < 0) {
                if (settings.loop) {
                    prev = totalSlides - 1;
                } else {
                    prev = 0;
                }
            }
            showSlide(prev);
        }
        
        function startAutoSlide(interval) {
            if (isPlaying) return;
            
            isPlaying = true;
            playInterval = setInterval(function() {
                if (settings.randomAutoSlide) {
                    var randomSlide = Math.floor(Math.random() * totalSlides);
                    showSlide(randomSlide);
                } else {
                    nextSlide();
                }
            }, interval);
            
            // Update play/pause button
            var playPauseBtn = document.querySelector('.play-pause');
            if (playPauseBtn) {
                playPauseBtn.innerHTML = '⏸';
            }
        }
        
        function stopAutoSlide() {
            if (!isPlaying) return;
            
            isPlaying = false;
            if (playInterval) {
                clearInterval(playInterval);
                playInterval = null;
            }
            
            // Update play/pause button
            var playPauseBtn = document.querySelector('.play-pause');
            if (playPauseBtn) {
                playPauseBtn.innerHTML = '▶';
            }
        }
        
        function toggleAutoSlide() {
            if (isPlaying) {
                stopAutoSlide();
            } else {
                var interval = parseInt(settings.autoSlide) * 1000;
                startAutoSlide(interval);
            }
        }
        
        function exitFullscreenMod() {
            stopAutoSlide();
            // Reset slide display
            for (var i = 0; i < slides.length; i++) {
                if (slides[i].style) {
                    slides[i].style.display = 'block';
                }
            }
        }
        
        // Keyboard shortcuts
        if (settings.keyBoardShortCut) {
            document.addEventListener('keydown', function(e) {
                switch(e.which) {
                    case 37: // Left arrow
                        e.preventDefault();
                        previousSlide();
                        break;
                    case 39: // Right arrow
                        e.preventDefault();
                        nextSlide();
                        break;
                    case 32: // Spacebar
                        e.preventDefault();
                        toggleAutoSlide();
                        break;
                    case 27: // Escape
                        e.preventDefault();
                        exitFullscreenMod();
                        break;
                }
            });
        }
        
        // Public API
        return {
            next: nextSlide,
            previous: previousSlide,
            goTo: showSlide,
            startAutoSlide: startAutoSlide,
            stopAutoSlide: stopAutoSlide,
            toggleAutoSlide: toggleAutoSlide,
            getCurrentSlide: function() { return currentSlide; },
            getTotalSlides: function() { return totalSlides; },
            destroy: function() {
                stopAutoSlide();
                // Remove navigation
                var nav = document.querySelector('.slides-nav');
                if (nav) {
                    nav.remove();
                }
                // Show all slides
                for (var i = 0; i < slides.length; i++) {
                    if (slides[i].style) {
                        slides[i].style.display = 'block';
                    }
                }
            }
        };
    };
    
})();