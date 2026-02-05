/**
 * 形状工厂模块
 * 负责根据形状类型调用相应的生成器
 */

var PPTXShapeFactory = (function() {
    var shapeGenerators = {};

    /**
     * 注册形状生成器
     * @param {string} shapeType - 形状类型
     * @param {Function} generator - 生成器函数
     */
    function registerGenerator(shapeType, generator) {
        shapeGenerators[shapeType] = generator;
    }

    /**
     * 批量注册形状生成器
     * @param {Array} shapeTypes - 形状类型数组
     * @param {Function} generator - 生成器函数
     */
    function registerGenerators(shapeTypes, generator) {
        shapeTypes.forEach(function(shapeType) {
            shapeGenerators[shapeType] = generator;
        });
    }

    /**
     * 生成形状
     * @param {string} shapeType - 形状类型
     * @param {Object} params - 形状参数
     * @returns {string} SVG字符串
     */
    function generateShape(shapeType, params) {
        var generator = shapeGenerators[shapeType];
        if (generator) {
            return generator(params);
        }
        console.warn("Unsupported shape type: " + shapeType);
        return "";
    }

    /**
     * 检查是否支持该形状类型
     * @param {string} shapeType - 形状类型
     * @returns {boolean} 是否支持
     */
    function isSupported(shapeType) {
        return shapeGenerators.hasOwnProperty(shapeType);
    }

    /**
     * 获取所有支持的形状类型
     * @returns {Array} 形状类型数组
     */
    function getSupportedTypes() {
        return Object.keys(shapeGenerators);
    }

    return {
        registerGenerator: registerGenerator,
        registerGenerators: registerGenerators,
        generateShape: generateShape,
        isSupported: isSupported,
        getSupportedTypes: getSupportedTypes
    };
})();

window.PPTXShapeFactory = PPTXShapeFactory;