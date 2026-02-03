/**
 * XML 工具函数模块
 * 提供XML节点遍历和查询功能
 */

var PPTXXmlUtils = (function() {
    /**
     * getTextByPathStr - 通过路径字符串获取XML文本
     * @param {Object} node - XML节点
     * @param {string} pathStr - 路径字符串（空格分隔）
     * @returns {*} 获取的值
     */
    function getTextByPathStr(node, pathStr) {
        return getTextByPathList(node, pathStr.trim().split(/\s+/));
    }

    /**
     * getTextByPathList - 通过路径数组获取XML文本
     * @param {Object} node - XML节点
     * @param {string[]} path - 路径数组
     * @returns {*} 获取的值
     */
    function getTextByPathList(node, path) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }

        if (node === undefined) {
            return undefined;
        }

        var l = path.length;
        for (var i = 0; i < l; i++) {
            node = node[path[i]];
            if (node === undefined) {
                return undefined;
            }
        }

        return node;
    }

    /**
     * setTextByPathList - 通过路径数组设置XML文本
     * @param {Object} node - XML节点
     * @param {string[]} path - 路径数组
     * @param {*} value - 要设置的值
     */
    function setTextByPathList(node, path, value) {
        if (path.constructor !== Array) {
            throw Error("Error of path type! path is not array.");
        }

        if (node === undefined) {
            return undefined;
        }

        Object.prototype.set = function (parts, value) {
            var obj = this;
            var lent = parts.length;
            for (var i = 0; i < lent; i++) {
                var p = parts[i];
                if (obj[p] === undefined) {
                    if (i == lent - 1) {
                        obj[p] = value;
                    } else {
                        obj[p] = {};
                    }
                }
                obj = obj[p];
            }
            return obj;
        };

        node.set(path, value);
    }

    /**
     * eachElement - 遍历节点数组或单个节点
     * @param {Object|Array} node - XML节点或节点数组
     * @param {Function} doFunction - 对每个节点执行的函数
     * @returns {string} 所有函数返回值的拼接
     */
    function eachElement(node, doFunction) {
        if (node === undefined) {
            return;
        }
        var result = "";
        if (node.constructor === Array) {
            var l = node.length;
            for (var i = 0; i < l; i++) {
                result += doFunction(node[i], i);
            }
        } else {
            result += doFunction(node, 0);
        }
        return result;
    }

    /**
     * angleToDegrees - 将角度转换为度数
     * @param {number} angle - 角度值（EMU单位）
     * @returns {number} 转换后的度数
     */
    function angleToDegrees(angle) {
        if (angle == "" || angle == null) {
            return 0;
        }
        return Math.round(angle / 60000);
    }

    /**
     * degreesToRadians - 将度数转换为弧度
     * @param {number} degrees - 度数
     * @returns {number} 弧度
     */
    function degreesToRadians(degrees) {
        if (degrees == "" || degrees == null || degrees == undefined) {
            return 0;
        }
        return degrees * (Math.PI / 180);
    }

    /**
     * escapeHtml - 转义HTML特殊字符
     * @param {string} text - 原始文本
     * @returns {string} 转义后的文本
     */
    function escapeHtml(text) {
        var map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, function (m) { return map[m]; });
    }

    /**
     * readXmlFile - 读取XML文件并解析为对象
     * @param {Object} zip - JSZip实例
     * @param {string} filename - 文件名
     * @param {boolean} isSlideContent - 是否为幻灯片内容
     * @param {number} appVersion - 应用版本
     * @returns {Object} 解析后的XML对象
     */
    function readXmlFile(zip, filename, isSlideContent, appVersion) {
        try {
            var fileContent = zip.file(filename).asText();
            if (isSlideContent && appVersion <= 12) {
                //< office2007
                //remove "<!CDATA[ ... ]]>" tag
                fileContent = fileContent.replace(/<!\[CDATA\[(.*?)\]\]>/g, '$1');
            }
            var xmlData = tXml(fileContent, { simplify: 1 });
            if (xmlData["?xml"] !== undefined) {
                return xmlData["?xml"];
            } else {
                return xmlData;
            }
        } catch (e) {
            //console.log("error readXmlFile: the file '" + filename + "' not exit")
            return null;
        }
    }

    return {
        getTextByPathStr: getTextByPathStr,
        getTextByPathList: getTextByPathList,
        setTextByPathList: setTextByPathList,
        eachElement: eachElement,
        angleToDegrees: angleToDegrees,
        degreesToRadians: degreesToRadians,
        escapeHtml: escapeHtml,
        readXmlFile: readXmlFile
    };
})();
