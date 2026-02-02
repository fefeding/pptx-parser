/**
 * Progress Utils
 * 进度条工具函数
 */

var PPTXProgressUtils = (function() {
    /**
     * 更新进度条
     * @param {number} percent - 进度百分比
     */
    function updateProgressBar(percent) {
        // console.log("percent: ", percent)
        var progressBarElemtnt = $(".slides-loading-progress-bar");
        progressBarElemtnt.width(percent + "%");
        progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
    }

    return {
        updateProgressBar: updateProgressBar
    };
})();
