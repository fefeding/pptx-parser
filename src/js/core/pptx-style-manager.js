var StyleManager = function() {
    this.styleTable = {};
};

    /**
 * 获取或创建样式对应的 CSS 类名
 * @param {String} styleText - 样式文本
 * @param {String} prefix - CSS 类名前缀 (可选)
 * @returns {String} CSS 类名
 */
StyleManager.prototype.getStyleClassName = function(styleText, prefix) {
    var cssName;
    if (styleText in this.styleTable) {
        cssName = this.styleTable[styleText]["name"];
    } else {
        prefix = prefix || "_css_";
        cssName = prefix + (Object.keys(this.styleTable).length + 1);
        this.styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }
    return cssName;
};

    /**
 * 获取样式表
 * @returns {Object} 样式表对象
 */
StyleManager.prototype.getStyleTable = function() {
    return this.styleTable;
};

    /**
 * 生成全局 CSS 样式
 * @returns {String} CSS 样式字符串
 */
StyleManager.prototype.generateGlobalCSS = function() {
    var css = "";
    for (var styleText in this.styleTable) {
        var cssName = this.styleTable[styleText]["name"];
        css += "." + cssName + " {" + styleText + "}\n";
    }
    return css;
};

    /**
 * 重置样式表
 */
StyleManager.prototype.reset = function() {
    this.styleTable = {};
};

    // 单例模式
var instance = null;

const PPTXStyleManager = {
    getInstance: function() {
        if (!instance) {
            instance = new StyleManager();
        }
        return instance;
    },
    // 快捷方法
    getStyleClassName: function(styleText, prefix) {
        return this.getInstance().getStyleClassName(styleText, prefix);
    },
    getStyleTable: function() {
        return this.getInstance().getStyleTable();
    },
    generateGlobalCSS: function() {
        return this.getInstance().generateGlobalCSS();
    },
    reset: function() {
        this.getInstance().reset();
    }
};

export { PPTXStyleManager };

// Also export to global scope for backward compatibility
window.PPTXStyleManager = PPTXStyleManager;
