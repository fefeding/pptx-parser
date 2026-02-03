/**
 * PPTX.js Constants
 * 常量定义文件
 */

var PPTXConstants = (function() {
    // RTL 语言数组
    var RTL_LANGS_ARRAY = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR", "ur-PK"];

    // 尺寸转换因子
    var SLIDE_FACTOR = 96 / 914400; // 1 EMU = 1 / 914400 inch, 1 inch = 96px
    var FONT_SIZE_FACTOR = 4 / 3.2;

    // 默认配置
    var DEFAULT_SETTINGS = {
        pptxFileUrl: "",
        fileInputId: "",
        slidesScale: "", // Change Slides scale by percent
        slideMode: false, /** true,false*/
        slideType: "divs2slidesjs", /*'divs2slidesjs' (default) , 'revealjs'(https://revealjs.com)  -TODO*/
        revealjsPath: "", /*path to js file of revealjs - TODO*/
        keyBoardShortCut: false, /** true,false ,condition: slideMode: true XXXXX - need to remove - this is doublcated*/
        mediaProcess: true, /** true,false: if true then process video and audio files */
        jsZipV2: false,
        themeProcess: true, /*true (default) , false, "colorsAndImageOnly"*/
        incSlide: {
            width: 0,
            height: 0
        },
        slideModeConfig: {
            first: 1,
            nav: true, /** true,false : show or not nav buttons*/
            navTxtColor: "black", /** color */
            keyBoardShortCut: true, /** true,false ,condition: */
            showSlideNum: true, /** true,false */
            showTotalSlideNum: true, /** true,false */
            autoSlide: true, /** false or seconds , F8 to active ,keyBoardShortCut: true */
            randomAutoSlide: false, /** true,false ,autoSlide:true */
            loop: false,  /** true,false */
            background: false, /** false or color*/
            transition: "default", /** transition type: "slid","fade","default","random" , to show transition efects :transitionTime > 0.5 */
            transitionTime: 1 /** transition time between slides in seconds */
        },
        revealjsConfig: {}
    };

    return {
        RTL_LANGS_ARRAY: RTL_LANGS_ARRAY,
        SLIDE_FACTOR: SLIDE_FACTOR,
        FONT_SIZE_FACTOR: FONT_SIZE_FACTOR,
        DEFAULT_SETTINGS: DEFAULT_SETTINGS
    };
})();
