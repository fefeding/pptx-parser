import { PPTXUtils } from '../core/utils.js';
import { PPTXColorUtils } from '../core/color.js';

    /**
 * 提取形状的基本几何属性和变换信息
 * @param {Object} node - 形状节点
 * @param {number} slideFactor - 坐标转换因子（默认 96/914400）
 * @param {Object} [pNode] - 父节点（用于某些属性继承）
 * @param {Object} [slideLayoutSpNode] - 幻灯片布局形状节点
 * @param {Object} [slideMasterSpNode] - 幻灯片母版形状节点
 * @returns {Object} 包含形状属性的对象
 */
function extractShapeProperties(node, slideFactor, pNode, slideLayoutSpNode, slideMasterSpNode) {
    const xfrmList = ["p:spPr", "a:xfrm"];
    const slideXfrmNode = PPTXUtils.getTextByPathList(node, xfrmList);
    const slideLayoutXfrmNode = PPTXUtils.getTextByPathList(slideLayoutSpNode, xfrmList);
    const slideMasterXfrmNode = PPTXUtils.getTextByPathList(slideMasterSpNode, xfrmList);

    const shapType = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    const custShapType = PPTXUtils.getTextByPathList(node, ["p:spPr", "a:custGeom"]);

    // 翻转处理
    let isFlipV = false;
    let isFlipH = false;
    let flip = "";
    if (PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1") {
        isFlipV = true;
    }
    if (PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
        isFlipH = true;
    }
    if (isFlipH && !isFlipV) {
        flip = " scale(-1,1)";
    } else if (!isFlipH && isFlipV) {
        flip = " scale(1,-1)";
    } else if (isFlipH && isFlipV) {
        flip = " scale(-1,-1)";
    }

    // 旋转角度
    const rotate = PPTXUtils.angleToDegrees(
        PPTXUtils.getTextByPathList(slideXfrmNode, ["attrs", "rot"])
    );

    // 文字旋转角度
    let txtRotate;
    const txtXframeNode = PPTXUtils.getTextByPathList(node, ["p:txXfrm"]);
    if (txtXframeNode !== undefined) {
        const txtXframeRot = PPTXUtils.getTextByPathList(txtXframeNode, ["attrs", "rot"]);
        if (txtXframeRot !== undefined) {
            txtRotate = PPTXUtils.angleToDegrees(txtXframeRot) + 90;
        }
    } else {
        txtRotate = rotate;
    }

    // 位置和尺寸
    const off = PPTXUtils.getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
    const ext = PPTXUtils.getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
    
    let x = 0, y = 0, w = 0, h = 0;
    if (off && off["x"] !== undefined && off["y"] !== undefined) {
        x = parseInt(off["x"]) * slideFactor;
        y = parseInt(off["y"]) * slideFactor;
    }
    if (ext && ext["cx"] !== undefined && ext["cy"] !== undefined) {
        w = parseInt(ext["cx"]) * slideFactor;
        h = parseInt(ext["cy"]) * slideFactor;
    }

    const shpId = PPTXUtils.getTextByPathList(node, ["attrs", "order"]);

    return {
        shapType: shapType,
        custShapType: custShapType,
        w: w,
        h: h,
        x: x,
        y: y,
        rotate: rotate,
        flip: flip,
        txtRotate: txtRotate,
        shpId: shpId,
        slideXfrmNode: slideXfrmNode,
        slideLayoutXfrmNode: slideLayoutXfrmNode,
        slideMasterXfrmNode: slideMasterXfrmNode
    };
}

const PPTXShapePropertyExtractor = {
    extractShapeProperties
};

export { PPTXShapePropertyExtractor };

// Also export to global scope for backward compatibility
// window.PPTXShapePropertyExtractor = PPTXShapePropertyExtractor; // Removed for ES modules

// Also add extractShapeProperties to PPTXShapeUtils for backward compatibility
// if (!window.PPTXShapeUtils) {
//     window.PPTXShapeUtils = {};
// }
// window.PPTXShapeUtils.extractShapeProperties = extractShapeProperties; // Removed for ES modules
