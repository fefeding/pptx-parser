import { PPTXUtils } from './utils.js';

/**
 * 颜色工具模块
 * 处理 PPTX 中的颜色、填充、图案等
 */

/**
 * 获取填充类型
 * @param {Object} node - 节点对象
 * @returns {string} 填充类型
 */
function getFillType(node) {
    //SOLID_FILL
    //PIC_FILL
    //GRADIENT_FILL
    //PATTERN_FILL
    //NO_FILL
    let fillType = "";
    if (node["a:noFill"] !== undefined) {
        fillType = "NO_FILL";
    }
    if (node["a:solidFill"] !== undefined) {
        fillType = "SOLID_FILL";
    }
    if (node["a:gradFill"] !== undefined) {
        fillType = "GRADIENT_FILL";
    }
    if (node["a:pattFill"] !== undefined) {
        fillType = "PATTERN_FILL";
    }
    if (node["a:blipFill"] !== undefined) {
        fillType = "PIC_FILL";
    }
    if (node["a:grpFill"] !== undefined) {
        fillType = "GROUP_FILL";
    }

    return fillType;
}
/**
 * 获取渐变填充
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {Object} 渐变填充对象
 */
function getGradientFill(node, warpObj) {
    const gsLst = node["a:gsLst"]["a:gs"];
    //get start color
    const color_ary = [];
    const tint_ary = [];
    for (let i = 0; i < gsLst.length; i++) {
        const lo_color = getSolidFill(gsLst[i], undefined, undefined, warpObj);
        color_ary[i] = lo_color;
    }
    //get rot
    const lin = node["a:lin"];
    let rot = 0;
    if (lin !== undefined) {
        rot = PPTXUtils.angleToDegrees(lin["attrs"]["ang"]) + 90;
    }
    return {
        "color": color_ary,
        "rot": rot
    };
}
/**
 * 获取图片填充
 * @param {string} type - 类型
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {Object} 图片填充对象
 */
function getPicFill(type, node, warpObj) {
    //rId
    // 图像属性处理已实现 - 支持平铺、拉伸、裁剪等属性
    // 参考: http://officeopenxml.com/drwPic-tile.php
    let img;
    const rId = node["a:blip"]["attrs"]["r:embed"];
    let imgPath;

    if (type == "slideBg" || type == "slide") {
        imgPath = PPTXUtils.getTextByPathList(warpObj, ["slideResObj", rId, "target"]);
    } else if (type == "slideLayoutBg") {
        imgPath = PPTXUtils.getTextByPathList(warpObj, ["layoutResObj", rId, "target"]);
    } else if (type == "slideMasterBg") {
        imgPath = PPTXUtils.getTextByPathList(warpObj, ["masterResObj", rId, "target"]);
    } else if (type == "themeBg") {
        imgPath = PPTXUtils.getTextByPathList(warpObj, ["themeResObj", rId, "target"]);
    } else if (type == "diagramBg") {
        imgPath = PPTXUtils.getTextByPathList(warpObj, ["diagramResObj", rId, "target"]);
    }

    if (imgPath === undefined) {
        return undefined;
    }

    const imgCache = PPTXUtils.getTextByPathList(warpObj, ["loaded-images", imgPath]);
    let imgData = null;

    if (imgCache === undefined) {
        imgPath = PPTXUtils.escapeHtml(imgPath);

        const imgExt = imgPath.split(".").pop();
        if (imgExt == "xml") {
            return undefined;
        }

        // 尝试解析图片路径，处理相对路径问题
        let imgFile = warpObj["zip"].file(imgPath);
        if (!imgFile && !imgPath.startsWith("ppt/")) {
            // 尝试添加 ppt/ 前缀
            imgFile = warpObj["zip"].file(`ppt/${imgPath}`);
        }

        if (!imgFile) {
            return undefined;
        }

        const imgArrayBuffer = imgFile.asArrayBuffer();
        const imgMimeType = PPTXUtils.getMimeType(imgExt);
        img = PPTXUtils.arrayBufferToBlobUrl(imgArrayBuffer, imgMimeType);
        imgData = PPTXUtils.base64ArrayBuffer(imgArrayBuffer);
        // 缓存对象，包含 URL 和 base64 数据
        const cacheObj = { url: img, data: imgData };
        PPTXUtils.setTextByPathList(warpObj, ["loaded-images", imgPath], cacheObj);
    } else {
        // 从缓存中提取 URL 和数据
        if (typeof imgCache === 'object' && imgCache.url) {
            img = imgCache.url;
            imgData = imgCache.data;
        } else {
            // 向后兼容：如果缓存中是字符串，则作为 URL
            img = imgCache;
            // 没有 base64 数据，imgData 保持 null
        }
    }

    // 为了保持向后兼容，默认返回图片 URL 字符串
    // 添加图像属性信息 - 支持平铺、拉伸或显示部分图像
    let fillProps = { img, imgData };

    // 解析 a:stretch 元素 - 拉伸填充
    if (node["a:stretch"] !== undefined) {
        const fillRect = node["a:stretch"]["a:fillRect"];
        const rectAttrs = fillRect !== undefined && fillRect["attrs"] !== undefined ? fillRect["attrs"] : null;

        fillProps = {
            img,
            imgData,
            stretch: true,
            tile: false,
            cropRect: null,
            fillRect: rectAttrs ? {
                l: parseInt(rectAttrs["l"]) / 100000,
                t: parseInt(rectAttrs["t"]) / 100000,
                r: parseInt(rectAttrs["r"]) / 100000,
                b: parseInt(rectAttrs["b"]) / 100000
            } : null
        };
    }
    // 解析 a:tile 元素 - 平铺填充
    else if (node["a:tile"] !== undefined) {
        const tileAttrs = node["a:tile"]["attrs"];

        fillProps = {
            img,
            imgData,
            stretch: false,
            tile: true,
            cropRect: null,
            fillRect: null,
            tileProps: tileAttrs ? {
                tx: tileAttrs["tx"] ? parseInt(tileAttrs["tx"]) / 100000 : 0,
                ty: tileAttrs["ty"] ? parseInt(tileAttrs["ty"]) / 100000 : 0,
                sx: tileAttrs["sx"] ? parseInt(tileAttrs["sx"]) / 100000 : 1,
                sy: tileAttrs["sy"] ? parseInt(tileAttrs["sy"]) / 100000 : 1,
                algn: tileAttrs["algn"] || "tl"
            } : null
        };
    }

    // 解析裁剪信息
    const srcRect = PPTXUtils.getTextByPathList(node, ["a:srcRect", "attrs"]);
    if (srcRect !== undefined && typeof fillProps === 'object') {
        fillProps.cropRect = {
            l: parseInt(srcRect["l"]) / 100000,
            t: parseInt(srcRect["t"]) / 100000,
            r: parseInt(srcRect["r"]) / 100000,
            b: parseInt(srcRect["b"]) / 100000
        };
    }

    return fillProps;
}
/**
 * 获取图案填充
 * @param {Object} node - 节点对象
 * @param {Object} warpObj - 包装对象
 * @returns {Array} 渐变数组
 */
function getPatternFill(node, warpObj) {
    let fgColor = "", bgColor = "", prst = "";
    const bgClr = node["a:bgClr"];
    const fgClr = node["a:fgClr"];
    prst = node["attrs"]["prst"];
    fgColor = getSolidFill(fgClr, undefined, undefined, warpObj);
    bgColor = getSolidFill(bgClr, undefined, undefined, warpObj);

    const linear_gradient = getLinerGrandient(prst, bgColor, fgColor);
    return linear_gradient;
}

/**
 * 获取线性渐变
 * @param {string} prst - 预设图案类型
 * @param {string} bgColor - 背景色
 * @param {string} fgColor - 前景色
 * @returns {Array} 渐变数组
 */
const getLinerGrandient = (prst, bgColor, fgColor) => {
    // dashDnDiag (Dashed Downward Diagonal)-V
    // dashHorz (Dashed Horizontal)-V
    // dashUpDiag(Dashed Upward DIagonal)-V
    // dashVert(Dashed Vertical)-V
    // diagBrick(Diagonal Brick)-V
    // divot(Divot)-VX
    // dkDnDiag(Dark Downward Diagonal)-V
    // dkHorz(Dark Horizontal)-V
    // dkUpDiag(Dark Upward Diagonal)-V
    // dkVert(Dark Vertical)-V
    // dotDmnd(Dotted Diamond)-VX
    // dotGrid(Dotted Grid)-V
    // horzBrick(Horizontal Brick)-V
    // lgCheck(Large Checker Board)-V
    // lgConfetti(Large Confetti)-V
    // lgGrid(Large Grid)-V
    // ltDnDiag(Light Downward Diagonal)-V
    // ltHorz(Light Horizontal)-V
    // ltUpDiag(Light Upward Diagonal)-V
    // ltVert(Light Vertical)-V
    // narHorz(Narrow Horizontal)-V
    // narVert(Narrow Vertical)-V
    // openDmnd(Open Diamond)-V
    // pct10(10 %)-V
    // pct20(20 %)-V
    // pct25(25 %)-V
    // pct30(30 %)-V
    // pct40(40 %)-V
    // pct5(5 %)-V
    // pct50(50 %)-V
    // pct60(60 %)-V
    // pct70(70 %)-V
    // pct75(75 %)-V
    // pct80(80 %)-V
    // pct90(90 %)-V
    // smCheck(Small Checker Board) -V
    // smConfetti(Small Confetti)-V
    // smGrid(Small Grid) -V
    // solidDmnd(Solid Diamond)-V
    // sphere(Sphere)-V
    // trellis(Trellis)-VX
    // wave(Wave)-V
    // wdDnDiag(Wide Downward Diagonal)-V
    // wdUpDiag(Wide Upward Diagonal)-V
    // weave(Weave)-V
    // zigZag(Zig Zag)-V
    // shingle(Shingle)-V
    // plaid(Plaid)-V
    // cross (Cross)
    // diagCross(Diagonal Cross)
    // dnDiag(Downward Diagonal)
    // horz(Horizontal)
    // upDiag(Upward Diagonal)
    // vert(Vertical)
    let size = "";
    switch (prst) {
        case "smGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "4px 4px"];
            break
        case "dotGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1px)  #" + bgColor + ";", "8px 8px"];
            break
        case "lgGrid":
            return ["linear-gradient(to right,  #" + fgColor + " -1px, transparent 1.5px ), " +
                "linear-gradient(to bottom,  #" + fgColor + " -1px, transparent 1.5px)  #" + bgColor + ";", "8px 8px"];
            break
        case "wdUpDiag":
            //return ["repeating-linear-gradient(-45deg,  #" + bgColor + ", #" + bgColor + " 1px,#" + fgColor + " 5px);"];
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            // return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 1px),  #" + fgColor + " 50%, transparent calc(50% + 1px),  transparent 100%) " +
            //     "#" + bgColor + ";", "6px 6px"];
            break
        case "dkUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
            break
        case "ltUpDiag":
            return ["repeating-linear-gradient(-45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "wdDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 4px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "dkDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , #" + bgColor + " 5px)" + "#" + fgColor + ";"];
            break
        case "ltDnDiag":
            return ["repeating-linear-gradient(45deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "dkHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
            break
        case "ltHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "narHorz":
            return ["repeating-linear-gradient(0deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "dkVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + bgColor + " 7px)" + "#" + fgColor + ";"];
            break
        case "ltVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 5px, #" + fgColor + " 7px)" + "#" + bgColor + ";"];
            break
        case "narVert":
            return ["repeating-linear-gradient(90deg, transparent 1px , transparent 2px, #" + fgColor + " 4px)" + "#" + bgColor + ";"];
            break
        case "lgCheck":
        case "smCheck":
            let pos = "";
            if (prst == "lgCheck") {
                size = "8px 8px";
                pos = "0 0, 4px 4px, 4px 4px, 8px 8px";
            } else {
                size = "4px 4px";
                pos = "0 0, 2px 2px, 2px 2px, 4px 4px";
            }
            return ["linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 0, transparent 75%,  #" + fgColor + " 0) " +
                "#" + bgColor + ";", size, pos];
            break
        // case "smCheck":
        //     return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
        //         "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%)  " +
        //         "#" + bgColor + ";", "4px 4px"];
        //     break 

        case "dashUpDiag":
            return ["repeating-linear-gradient(152deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "dashDnDiag":
            return ["repeating-linear-gradient(45deg, #" + fgColor + ", #" + fgColor + " 5% , transparent 0, transparent 70%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "diagBrick":
            return ["linear-gradient(45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 15%,  #" + fgColor + " 30%, transparent 30%), " +
                "linear-gradient(-45deg, transparent 65%,  #" + fgColor + " 80%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "horzBrick":
            return ["linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(335deg, #" + bgColor + " 1.6px, transparent 1.6px), " +
                "linear-gradient(155deg, #" + bgColor + " 1.6px, transparent 1.6px) " +
                "#" + fgColor + ";", "4px 4px", "0 0.15px, 0.3px 2.5px, 2px 2.15px, 2.35px 0.4px"];
            break

        case "dashVert":
            return ["linear-gradient(0deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(90deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "dashHorz":
            return ["linear-gradient(90deg,  #" + bgColor + " 30%, transparent 30%)," +
                "linear-gradient(0deg,transparent, transparent 40%, #" + fgColor + " 40%, #" + fgColor + " 60% , transparent 60%)" +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "solidDmnd":
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "8px 8px"];
            break
        case "openDmnd":
            return ["linear-gradient(45deg, transparent 0%, transparent calc(50% - 0.5px),  #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%), " +
                "linear-gradient(-45deg, transparent 0%, transparent calc(50% - 0.5px) , #" + fgColor + " 50%, transparent calc(50% + 0.5px),  transparent 100%) " +
                "#" + bgColor + ";", "8px 8px"];
            break

        case "dotDmnd":
            return ["radial-gradient(#" + fgColor + " 15%, transparent 0), " +
                "radial-gradient(#" + fgColor + " 15%, transparent 0) " +
                "#" + bgColor + ";", "4px 4px", "0 0, 2px 2px"];
            break
        case "zigZag":
        case "wave":
            if (prst == "zigZag") size = "0";
            else size = "1px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%) 50px " + size + ", " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "lgConfetti":
        case "smConfetti":
            if (prst == "lgConfetti") size = "4px 4px";
            else size = "2px 2px";
            return ["linear-gradient(135deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px, " +
                "linear-gradient(225deg,  #" + fgColor + " 25%, transparent 25%), " +
                "linear-gradient(315deg,  #" + fgColor + " 25%, transparent 25%) 50px 1px , " +
                "linear-gradient(45deg,  #" + fgColor + " 25%, transparent 25%) " +
                "#" + bgColor + ";", size];
            break
        // case "weave":
        //     return ["linear-gradient(45deg,  #" + bgColor + " 5%, transparent 25%) 50px 0, " +
        //         "linear-gradient(135deg,  #" + bgColor + " 25%, transparent 25%) 50px 0, " +
        //         "linear-gradient(45deg,  #" + bgColor + " 25%, transparent 25%) " +
        //         "#" + fgColor + ";", "4px 4px"];
        //     //background: linear-gradient(45deg, #dca 12%, transparent 0, transparent 88%, #dca 0),
        //     //linear-gradient(135deg, transparent 37 %, #a85 0, #a85 63 %, transparent 0),
        //     //linear-gradient(45deg, transparent 37 %, #dca 0, #dca 63 %, transparent 0) #753;
        //     // background-size: 25px 25px;
        //     break;

        case "plaid":
            return ["linear-gradient(0deg, transparent, transparent 25%, #" + fgColor + "33 25%, #" + fgColor + "33 50%)," +
                "linear-gradient(90deg, transparent, transparent 25%, #" + fgColor + "66 25%, #" + fgColor + "66 50%) " +
                "#" + bgColor + ";", "4px 4px"];
            /**
                background-color: #6677dd;
                background-image: 
                repeating-linear-gradient(0deg, transparent, transparent 35px, rgba(255, 255, 255, 0.2) 35px, rgba(255, 255, 255, 0.2) 70px), 
                repeating-linear-gradient(90deg, transparent, transparent 35px, rgba(255,255,255,0.4) 35px, rgba(255,255,255,0.4) 70px);
             */
            break;
        case "sphere":
            return ["radial-gradient(#" + fgColor + " 50%, transparent 50%)," +
                "#" + bgColor + ";", "4px 4px"];
            break
        case "weave":
        case "shingle":
            return ["linear-gradient(45deg, #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px), " +
                "linear-gradient(-45deg,  #" + bgColor + " 1.31px , #" + fgColor + " 1.4px, #" + fgColor + " 1.5px, transparent 1.5px, transparent 4.2px, #" + fgColor + " 4.2px, #" + fgColor + " 4.3px, transparent 4.31px) 0 4px, " +
                "#" + bgColor + ";", "4px 8px"];
            break
        //background:
        //linear-gradient(45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px),
        //linear-gradient(-45deg, #708090 1.31px, #d9ecff 1.4px, #d9ecff 1.5px, transparent 1.5px, transparent 4.2px, #d9ecff 4.2px, #d9ecff 4.3px, transparent 4.31px)0 4px;
        //background-color:#708090;
        //background-size: 4px 8px;
        case "pct5":
        case "pct10":
        case "pct20":
        case "pct25":
        case "pct30":
        case "pct40":
        case "pct50":
        case "pct60":
        case "pct70":
        case "pct75":
        case "pct80":
        case "pct90":
        //case "dotDmnd":
        case "trellis":
        case "divot":
            let px_pr_ary;
            switch (prst) {
                case "pct5":
                    px_pr_ary = ["0.3px", "10%", "2px 2px"];
                    break
                case "divot":
                    px_pr_ary = ["0.3px", "40%", "4px 4px"];
                    break
                case "pct10":
                    px_pr_ary = ["0.3px", "20%", "2px 2px"];
                    break
                case "pct20":
                    //case "dotDmnd":
                    px_pr_ary = ["0.2px", "40%", "2px 2px"];
                    break
                case "pct25":
                    px_pr_ary = ["0.2px", "50%", "2px 2px"];
                    break
                case "pct30":
                    px_pr_ary = ["0.5px", "50%", "2px 2px"];
                    break
                case "pct40":
                    px_pr_ary = ["0.5px", "70%", "2px 2px"];
                    break
                case "pct50":
                    px_pr_ary = ["0.09px", "90%", "2px 2px"];
                    break
                case "pct60":
                    px_pr_ary = ["0.3px", "90%", "2px 2px"];
                    break
                case "pct70":
                case "trellis":
                    px_pr_ary = ["0.5px", "95%", "2px 2px"];
                    break
                case "pct75":
                    px_pr_ary = ["0.65px", "100%", "2px 2px"];
                    break
                case "pct80":
                    px_pr_ary = ["0.85px", "100%", "2px 2px"];
                    break
                case "pct90":
                    px_pr_ary = ["1px", "100%", "2px 2px"];
                    break
            }
            return ["radial-gradient(#" + fgColor + " " + px_pr_ary[0] + ", transparent " + px_pr_ary[1] + ")," +
                "#" + bgColor + ";", px_pr_ary[2]];
            break
        default:
            return [0, 0];
    }
}

function getSolidFill(node, clrMap, phClr, warpObj) {

    if (node === undefined) {
        return undefined;
    }

    //console.log("getSolidFill node: ", node)
    let color = "";
    let clrNode;
    if (node["a:srgbClr"] !== undefined) {
        clrNode = node["a:srgbClr"];
        color = PPTXUtils.getTextByPathList(clrNode, ["attrs", "val"]); //#...
    } else if (node["a:schemeClr"] !== undefined) { //a:schemeClr
        clrNode = node["a:schemeClr"];
        const schemeClr = PPTXUtils.getTextByPathList(clrNode, ["attrs", "val"]);
        color = getSchemeColorFromTheme("a:" + schemeClr, clrMap, phClr, warpObj);
        //console.log("schemeClr: ", schemeClr, "color: ", color)
    } else if (node["a:scrgbClr"] !== undefined) {
        clrNode = node["a:scrgbClr"];
        //<a:scrgbClr r="50%" g="50%" b="50%"/>  //Need to test/////////////////////////////////////////////
        const defBultColorVals = clrNode["attrs"];
        const red = (defBultColorVals["r"].indexOf("%") != -1) ? defBultColorVals["r"].split("%").shift() : defBultColorVals["r"];
        const green = (defBultColorVals["g"].indexOf("%") != -1) ? defBultColorVals["g"].split("%").shift() : defBultColorVals["g"];
        const blue = (defBultColorVals["b"].indexOf("%") != -1) ? defBultColorVals["b"].split("%").shift() : defBultColorVals["b"];
        //const scrgbClr = red + "," + green + "," + blue;
        color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100));
        //console.log("scrgbClr: " + scrgbClr);

    } else if (node["a:prstClr"] !== undefined) {
        clrNode = node["a:prstClr"];
        //<a:prstClr val="black"/>  //Need to test/////////////////////////////////////////////
        const prstClr = PPTXUtils.getTextByPathList(clrNode, ["attrs", "val"]); //node["a:prstClr"]["attrs"]["val"];
        color = getColorName2Hex(prstClr);
        //console.log("blip prstClr: ", prstClr, " => hexClr: ", color);
    } else if (node["a:hslClr"] !== undefined) {
        clrNode = node["a:hslClr"];
        //<a:hslClr hue="14400000" sat="100%" lum="50%"/>  //Need to test/////////////////////////////////////////////
        const defBultColorVals = clrNode["attrs"];
        const hue = Number(defBultColorVals["hue"]) / 100000;
        const sat = Number((defBultColorVals["sat"].indexOf("%") != -1) ? defBultColorVals["sat"].split("%").shift() : defBultColorVals["sat"]) / 100;
        const lum = Number((defBultColorVals["lum"].indexOf("%") != -1) ? defBultColorVals["lum"].split("%").shift() : defBultColorVals["lum"]) / 100;
        //const hslClr = defBultColorVals["hue"] + "," + defBultColorVals["sat"] + "," + defBultColorVals["lum"];
        const hsl2rgb = hslToRgb(hue, sat, lum);
        color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b);
        // cnvrtHslColor2Hex - 已通过 hslToRgb 实现，无需额外函数
    } else if (node["a:sysClr"] !== undefined) {
        clrNode = node["a:sysClr"];
        //<a:sysClr val="windowText" lastClr="000000"/>  //Need to test/////////////////////////////////////////////
        const sysClr = PPTXUtils.getTextByPathList(clrNode, ["attrs", "lastClr"]);
        if (sysClr !== undefined) {
            color = sysClr;
        }
    }
    //console.log("color: [%cstart]", "color: #" + color, tinycolor(color).toHslString(), color)

    //fix color -------------------------------------------------------- 
    // 透明度、色相偏移、饱和度偏移等颜色修正已实现
    //
    //1. "alpha":
    //Specifies the opacity as expressed by a percentage value.
    // [Example: The following represents a green solid fill which is 50 % opaque
    // < a: solidFill >
    //     <a:srgbClr val="00FF00">
    //         <a:alpha val="50%" />
    //     </a:srgbClr>
    // </a: solidFill >
    let isAlpha = false;
    const alpha = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:alpha", "attrs", "val"])) / 100000;
    //console.log("alpha: ", alpha)
    if (!isNaN(alpha)) {
        // const al_color = new colz.Color(color);
        // al_color.setAlpha(alpha);
        // const ne_color = al_color.rgba.toString();
        // color = (rgba2hex(ne_color))
        const al_color = tinycolor(color);
        al_color.setAlpha(alpha);
        color = al_color.toHex8()
        isAlpha = true;
        //console.log("al_color: ", al_color, ", color: ", color)
    }
    //2. "alphaMod":
    // Specifies the opacity as expressed by a percentage relative to the input color.
    //     [Example: The following represents a green solid fill which is 50 % opaque
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:alphaMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //3. "alphaOff":
    // Specifies the opacity as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the opacity beyond 100 %, decreases never decrease
    // the opacity below 0 %.
    // [Example: The following represents a green solid fill which is 90 % opaque
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:alphaOff val="-10%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //4. "blue":
    //Specifies the value of the blue component.The assigned value is specified as a
    //percentage with 0 % indicating minimal blue and 100 % indicating maximum blue.
    //  [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //      to value RRGGBB = (00, FF, FF)
    //          <a: solidFill >
    //              <a:srgbClr val="00FF00">
    //                  <a:blue val="100%" />
    //              </a:srgbClr>
    //          </a: solidFill >
    //5. "blueMod"
    // Specifies the blue component as expressed by a percentage relative to the input color
    // component.Increases never increase the blue component beyond 100 %, decreases
    // never decrease the blue component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    //     to value RRGGBB = (00, 00, 80)
    //     < a: solidFill >
    //         <a:srgbClr val="0000FF">
    //             <a:blueMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //6. "blueOff"
    // Specifies the blue component as expressed by a percentage offset increase or decrease
    // to the input color component.Increases never increase the blue component
    // beyond 100 %, decreases never decrease the blue component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    // to value RRGGBB = (00, 00, CC)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:blueOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //7. "comp" - This element specifies that the color rendered should be the complement of its input color with the complement
    // being defined as such.Two colors are called complementary if, when mixed they produce a shade of grey.For
    // instance, the complement of red which is RGB(255, 0, 0) is cyan.(<a:comp/>)

    //8. "gamma" - This element specifies that the output color rendered by the generating application should be the sRGB gamma
    //              shift of the input color.

    //9. "gray" - This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue
    //              primaries.

    //10. "green":
    // Specifies the value of the green component. The assigned value is specified as a
    // percentage with 0 % indicating minimal green and 100 % indicating maximum green.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, 00, FF)
    // to value RRGGBB = (00, FF, FF)
    //     < a: solidFill >
    //         <a:srgbClr val="0000FF">
    //             <a:green val="100%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //11. "greenMod":
    // Specifies the green component as expressed by a percentage relative to the input color
    // component.Increases never increase the green component beyond 100 %, decreases
    // never decrease the green component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, 80, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:greenMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //12. "greenOff":
    // Specifies the green component as expressed by a percentage offset increase or decrease
    // to the input color component.Increases never increase the green component
    // beyond 100 %, decreases never decrease the green component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, CC, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:greenOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //13. "hue" (This element specifies a color using the HSL color model):
    // This element specifies the input color with the specified hue, but with its saturation and luminance unchanged.
    // < a: solidFill >
    //     <a:hslClr hue="14400000" sat="100%" lum="50%">
    // </a:solidFill>
    // <a:solidFill>
    //     <a:hslClr hue="0" sat="100%" lum="50%">
    //         <a:hue val="14400000"/>
    //     <a:hslClr/>
    // </a:solidFill>

    //14. "hueMod" (This element specifies a color using the HSL color model):
    // Specifies the hue as expressed by a percentage relative to the input color.
    // [Example: The following manipulates the fill color from having RGB value RRGGBB = (00, FF, 00) to value RRGGBB = (FF, FF, 00)
    //         < a: solidFill >
    //             <a:srgbClr val="00FF00">
    //                 <a:hueMod val="50%" />
    //             </a:srgbClr>
    //         </a: solidFill >

    const hueMod = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:hueMod", "attrs", "val"])) / 100000;
    //console.log("hueMod: ", hueMod)
    if (!isNaN(hueMod)) {
        color = applyHueMod(color, hueMod, isAlpha);
    }
    //15. "hueOff"(This element specifies a color using the HSL color model):
    // Specifies the actual angular value of the shift.The result of the shift shall be between 0
    // and 360 degrees.Shifts resulting in angular values less than 0 are treated as 0. Shifts
    // resulting in angular values greater than 360 are treated as 360.
    // [Example:
    //     The following increases the hue angular value by 10 degrees.
    //     < a: solidFill >
    //         <a:hslClr hue="0" sat="100%" lum="50%"/>
    //             <a:hueOff val="600000"/>
    //     </a: solidFill >
    // 15. "hueOff"
    // Specifies the hue offset for a color adjustment transform
    const hueOff = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:hueOff", "attrs", "val"])) / 100000;
    if (!isNaN(hueOff)) {
        const hslColor = tinycolor(color).toHsl();
        hslColor.h = (hslColor.h + hueOff * 360) % 360;
        if (hslColor.h < 0) hslColor.h += 360;
        color = tinycolor(hslColor).toHexString().substring(1);
        // 保留原有的 alpha 通道
        if (isAlpha) {
            const alphaVal = tinycolor(color).getAlpha();
            color = tinycolor(color).setAlpha(alphaVal).toHex8().substring(1);
        }
    }

    //16. "inv" (inverse)
    //This element specifies the inverse of its input color.
    //The inverse of red (1, 0, 0) is cyan (0, 1, 1 ).
    // The following represents cyan, the inverse of red:
    // <a:solidFill>
    //     <a:srgbClr val="FF0000">
    //         <a:inv />
    //     </a:srgbClr>
    // </a:solidFill>

    //17. "invGamma" - This element specifies that the output color rendered by the generating application should be the inverse sRGB
    //                  gamma shift of the input color.

    //18. "lum":
    // This element specifies the input color with the specified luminance, but with its hue and saturation unchanged.
    // Typically luminance values fall in the range[0 %, 100 %].
    // The following two solid fills are equivalent:
    // <a:solidFill>
    //     <a:hslClr hue="14400000" sat="100%" lum="50%">
    // </a:solidFill>
    // <a:solidFill>
    //     <a:hslClr hue="14400000" sat="100%" lum="0%">
    //         <a:lum val="50%" />
    //     <a:hslClr />
    // </a:solidFill>
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    // to value RRGGBB = (00, 66, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lum val="20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    //19. "lumMod":
    // Specifies the luminance as expressed by a percentage relative to the input color.
    // Increases never increase the luminance beyond 100 %, decreases never decrease the
    // luminance below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, 75, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lumMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    const lumMod = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:lumMod", "attrs", "val"])) / 100000;
    //console.log("lumMod: ", lumMod)
    if (!isNaN(lumMod)) {
        color = applyLumMod(color, lumMod, isAlpha);
    }
    //const lumMod_color = applyLumMod(color, 0.5);
    //console.log("lumMod_color: ", lumMod_color)
    //20. "lumOff"
    // Specifies the luminance as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the luminance beyond 100 %, decreases never
    // decrease the luminance below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, 99, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:lumOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    const lumOff = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:lumOff", "attrs", "val"])) / 100000;
    //console.log("lumOff: ", lumOff)
    if (!isNaN(lumOff)) {
        color = applyLumOff(color, lumOff, isAlpha);
    }

    //21. "red":
    // Specifies the value of the red component.The assigned value is specified as a percentage
    // with 0 % indicating minimal red and 100 % indicating maximum red.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (FF, FF, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:red val="100%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //22. "redMod":
    // Specifies the red component as expressed by a percentage relative to the input color
    // component.Increases never increase the red component beyond 100 %, decreases never
    // decrease the red component below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    //     to value RRGGBB = (80, 00, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="FF0000">
    //             <a:redMod val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    //23. "redOff":
    // Specifies the red component as expressed by a percentage offset increase or decrease to
    // the input color component.Increases never increase the red component beyond 100 %,
    //     decreases never decrease the red component below 0 %.
    //     [Example: The following manipulates the fill from having RGB value RRGGBB = (FF, 00, 00)
    //     to value RRGGBB = (CC, 00, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="FF0000">
    //             <a:redOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >

    //23. "sat":
    // This element specifies the input color with the specified saturation, but with its hue and luminance unchanged.
    // Typically saturation values fall in the range[0 %, 100 %].
    // [Example:
    //     The following two solid fills are equivalent:
    //     <a:solidFill>
    //         <a:hslClr hue="14400000" sat="100%" lum="50%">
    //     </a:solidFill>
    //     <a:solidFill>
    //         <a:hslClr hue="14400000" sat="0%" lum="50%">
    //             <a:sat val="100000" />
    //         <a:hslClr />
    //     </a:solidFill>
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (40, C0, 40)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:sat val="50%" />
    //         </a:srgbClr>
    //     <a: solidFill >
    // end example]

    //24. "satMod":
    // Specifies the saturation as expressed by a percentage relative to the input color.
    // Increases never increase the saturation beyond 100 %, decreases never decrease the
    // saturation below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (66, 99, 66)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:satMod val="20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    const satMod = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:satMod", "attrs", "val"])) / 100000;
    if (!isNaN(satMod)) {
        color = applySatMod(color, satMod, isAlpha);
    }
    //25. "satOff":
    // Specifies the saturation as expressed by a percentage offset increase or decrease to the
    // input color.Increases never increase the saturation beyond 100 %, decreases never
    // decrease the saturation below 0 %.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (19, E5, 19)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:satOff val="-20%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // 25. "satOff"
    // Specifies the saturation offset for a color adjustment transform
    const satOff = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:satOff", "attrs", "val"])) / 100000;
    if (!isNaN(satOff)) {
        let hslColor = tinycolor(color).toHsl();
        hslColor.s = Math.min(100, Math.max(0, hslColor.s + satOff * 100));
        color = tinycolor(hslColor).toHexString().substring(1);
        // 保留原有的 alpha 通道
        if (isAlpha) {
            let alphaVal = tinycolor(color).getAlpha();
            color = tinycolor(color).setAlpha(alphaVal).toHex8().substring(1);
        }
    }

    //26. "shade":
    // This element specifies a darker version of its input color.A 10 % shade is 10 % of the input color combined with 90 % black.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (00, BC, 00)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:shade val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    // end example]
    const shade = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:shade", "attrs", "val"])) / 100000;
    if (!isNaN(shade)) {
        color = applyShade(color, shade, isAlpha);
    }
    //27.  "tint":
    // This element specifies a lighter version of its input color.A 10 % tint is 10 % of the input color combined with
    // 90 % white.
    // [Example: The following manipulates the fill from having RGB value RRGGBB = (00, FF, 00)
    //     to value RRGGBB = (BC, FF, BC)
    //     < a: solidFill >
    //         <a:srgbClr val="00FF00">
    //             <a:tint val="50%" />
    //         </a:srgbClr>
    //     </a: solidFill >
    const tint = parseInt(PPTXUtils.getTextByPathList(clrNode, ["a:tint", "attrs", "val"])) / 100000;
    if (!isNaN(tint)) {
        color = applyTint(color, tint, isAlpha);
    }
    //console.log("color [%cfinal]: ", "color: #" + color, tinycolor(color).toHslString(), color)

    return color;
}
function toHex(n) {
    let hex = n.toString(16);
    while (hex.length < 2) { hex = "0" + hex; }
    return hex;
}
function hslToRgb(hue, sat, light) {
    let t1, t2, r, g, b;
    hue = hue / 60;
    if (light <= 0.5) {
        t2 = light * (sat + 1);
    } else {
        t2 = light + sat - (light * sat);
    }
    t1 = light * 2 - t2;
    r = hueToRgb(t1, t2, hue + 2) * 255;
    g = hueToRgb(t1, t2, hue) * 255;
    b = hueToRgb(t1, t2, hue - 2) * 255;
    return { r: r, g: g, b: b };
}
function hueToRgb(t1, t2, hue) {
    if (hue < 0) hue += 6;
    if (hue >= 6) hue -= 6;
    if (hue < 1) return (t2 - t1) * hue + t1;
    else if (hue < 3) return t2;
    else if (hue < 4) return (t2 - t1) * (4 - hue) + t1;
    else return t1;
}
function getColorName2Hex(name) {
    let hex;
    const colorName = ['white', 'AliceBlue', 'AntiqueWhite', 'Aqua', 'Aquamarine', 'Azure', 'Beige', 'Bisque', 'black', 'BlanchedAlmond', 'Blue', 'BlueViolet', 'Brown', 'BurlyWood', 'CadetBlue', 'Chartreuse', 'Chocolate', 'Coral', 'CornflowerBlue', 'Cornsilk', 'Crimson', 'Cyan', 'DarkBlue', 'DarkCyan', 'DarkGoldenRod', 'DarkGray', 'DarkGrey', 'DarkGreen', 'DarkKhaki', 'DarkMagenta', 'DarkOliveGreen', 'DarkOrange', 'DarkOrchid', 'DarkRed', 'DarkSalmon', 'DarkSeaGreen', 'DarkSlateBlue', 'DarkSlateGray', 'DarkSlateGrey', 'DarkTurquoise', 'DarkViolet', 'DeepPink', 'DeepSkyBlue', 'DimGray', 'DimGrey', 'DodgerBlue', 'FireBrick', 'FloralWhite', 'ForestGreen', 'Fuchsia', 'Gainsboro', 'GhostWhite', 'Gold', 'GoldenRod', 'Gray', 'Grey', 'Green', 'GreenYellow', 'HoneyDew', 'HotPink', 'IndianRed', 'Indigo', 'Ivory', 'Khaki', 'Lavender', 'LavenderBlush', 'LawnGreen', 'LemonChiffon', 'LightBlue', 'LightCoral', 'LightCyan', 'LightGoldenRodYellow', 'LightGray', 'LightGrey', 'LightGreen', 'LightPink', 'LightSalmon', 'LightSeaGreen', 'LightSkyBlue', 'LightSlateGray', 'LightSlateGrey', 'LightSteelBlue', 'LightYellow', 'Lime', 'LimeGreen', 'Linen', 'Magenta', 'Maroon', 'MediumAquaMarine', 'MediumBlue', 'MediumOrchid', 'MediumPurple', 'MediumSeaGreen', 'MediumSlateBlue', 'MediumSpringGreen', 'MediumTurquoise', 'MediumVioletRed', 'MidnightBlue', 'MintCream', 'MistyRose', 'Moccasin', 'NavajoWhite', 'Navy', 'OldLace', 'Olive', 'OliveDrab', 'Orange', 'OrangeRed', 'Orchid', 'PaleGoldenRod', 'PaleGreen', 'PaleTurquoise', 'PaleVioletRed', 'PapayaWhip', 'PeachPuff', 'Peru', 'Pink', 'Plum', 'PowderBlue', 'Purple', 'RebeccaPurple', 'Red', 'RosyBrown', 'RoyalBlue', 'SaddleBrown', 'Salmon', 'SandyBrown', 'SeaGreen', 'SeaShell', 'Sienna', 'Silver', 'SkyBlue', 'SlateBlue', 'SlateGray', 'SlateGrey', 'Snow', 'SpringGreen', 'SteelBlue', 'Tan', 'Teal', 'Thistle', 'Tomato', 'Turquoise', 'Violet', 'Wheat', 'White', 'WhiteSmoke', 'Yellow', 'YellowGreen'];
    const colorHex = ['ffffff', 'f0f8ff', 'faebd7', '00ffff', '7fffd4', 'f0ffff', 'f5f5dc', 'ffe4c4', '000000', 'ffebcd', '0000ff', '8a2be2', 'a52a2a', 'deb887', '5f9ea0', '7fff00', 'd2691e', 'ff7f50', '6495ed', 'fff8dc', 'dc143c', '00ffff', '00008b', '008b8b', 'b8860b', 'a9a9a9', 'a9a9a9', '006400', 'bdb76b', '8b008b', '556b2f', 'ff8c00', '9932cc', '8b0000', 'e9967a', '8fbc8f', '483d8b', '2f4f4f', '2f4f4f', '00ced1', '9400d3', 'ff1493', '00bfff', '696969', '696969', '1e90ff', 'b22222', 'fffaf0', '228b22', 'ff00ff', 'dcdcdc', 'f8f8ff', 'ffd700', 'daa520', '808080', '808080', '008000', 'adff2f', 'f0fff0', 'ff69b4', 'cd5c5c', '4b0082', 'fffff0', 'f0e68c', 'e6e6fa', 'fff0f5', '7cfc00', 'fffacd', 'add8e6', 'f08080', 'e0ffff', 'fafad2', 'd3d3d3', 'd3d3d3', '90ee90', 'ffb6c1', 'ffa07a', '20b2aa', '87cefa', '778899', '778899', 'b0c4de', 'ffffe0', '00ff00', '32cd32', 'faf0e6', 'ff00ff', '800000', '66cdaa', '0000cd', 'ba55d3', '9370db', '3cb371', '7b68ee', '00fa9a', '48d1cc', 'c71585', '191970', 'f5fffa', 'ffe4e1', 'ffe4b5', 'ffdead', '000080', 'fdf5e6', '808000', '6b8e23', 'ffa500', 'ff4500', 'da70d6', 'eee8aa', '98fb98', 'afeeee', 'db7093', 'ffefd5', 'ffdab9', 'cd853f', 'ffc0cb', 'dda0dd', 'b0e0e6', '800080', '663399', 'ff0000', 'bc8f8f', '4169e1', '8b4513', 'fa8072', 'f4a460', '2e8b57', 'fff5ee', 'a0522d', 'c0c0c0', '87ceeb', '6a5acd', '708090', '708090', 'fffafa', '00ff7f', '4682b4', 'd2b48c', '008080', 'd8bfd8', 'ff6347', '40e0d0', 'ee82ee', 'f5deb3', 'ffffff', 'f5f5f5', 'ffff00', '9acd32'];
    const findIndx = colorName.indexOf(name);
    if (findIndx != -1) {
        hex = colorHex[findIndx];
    }
    return hex;
}
function getSchemeColorFromTheme(schemeClr, clrMap, phClr, warpObj) {
    //<p:clrMap ...> in slide master
    // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1" slideLayoutClrOvride
    //console.log("getSchemeColorFromTheme: schemeClr: ", schemeClr, ",clrMap: ", clrMap)
    let slideLayoutClrOvride;
    let color;
    if (clrMap !== undefined) {
        slideLayoutClrOvride = clrMap;//PPTXUtils.getTextByPathList(clrMap, ["p:sldMaster", "p:clrMap", "attrs"])
    } else {
        let sldClrMapOvr = PPTXUtils.getTextByPathList(warpObj["slideContent"], ["p:sld", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
        if (sldClrMapOvr !== undefined) {
            slideLayoutClrOvride = sldClrMapOvr;
        } else {
sldClrMapOvr = PPTXUtils.getTextByPathList(warpObj["slideLayoutContent"], ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);
            if (sldClrMapOvr !== undefined) {
                slideLayoutClrOvride = sldClrMapOvr;
            } else {
                slideLayoutClrOvride = PPTXUtils.getTextByPathList(warpObj["slideMasterContent"], ["p:sldMaster", "p:clrMap", "attrs"]);
            }

        }
    }
    //console.log("getSchemeColorFromTheme slideLayoutClrOvride: ", slideLayoutClrOvride);
    const schmClrName = schemeClr.substr(2);
    if (schmClrName == "phClr" && phClr !== undefined) {
        color = phClr;
    } else {
        if (slideLayoutClrOvride !== undefined) {
            switch (schmClrName) {
                case "tx1":
                case "tx2":
                case "bg1":
                case "bg2":
                    schemeClr = "a:" + slideLayoutClrOvride[schmClrName];
                    break;
            }
        } else {
            switch (schmClrName) {
                case "tx1":
                    schemeClr = "a:dk1";
                    break;
                case "tx2":
                    schemeClr = "a:dk2";
                    break;
                case "bg1":
                    schemeClr = "a:lt1";
                    break;
                case "bg2":
                    schemeClr = "a:lt2";
                    break;
            }
        }
        //console.log("getSchemeColorFromTheme:  schemeClr: ", schemeClr);
        const refNode = PPTXUtils.getTextByPathList(warpObj["themeContent"], ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
color = PPTXUtils.getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
        //console.log("themeContent: color", color);
        if (color === undefined && refNode !== undefined) {
            color = PPTXUtils.getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
        }
    }
    //console.log(color)
    return color;
}

function extractChartData(serNode) {

    const dataMat = [];

    if (serNode === undefined) {
        return dataMat;
    }

    if (serNode["c:xVal"] !== undefined) {
        let dataRow = [];
        PPTXUtils.eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
        dataRow = [];
        PPTXUtils.eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
    } else {
        PPTXUtils.eachElement(serNode, function (innerNode, index) {
dataRow = [];
            const colName = PPTXUtils.getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

            // Category (string or number)
            const rowNames = {};
            if (PPTXUtils.getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                PPTXUtils.eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function (innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            } else if (PPTXUtils.getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                PPTXUtils.eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            }

            // Value
            if (PPTXUtils.getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                PPTXUtils.eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function (innerNode, index) {
                    dataRow.push({ x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"]) });
                    return "";
                });
            }

            dataMat.push({ key: colName, values: dataRow, xlabels: rowNames });
            return "";
        });
    }

    return dataMat;
}

    // ===== Node functions =====
    /**
 * getTextByPathStr
 * @param {Object} node
 * @param {string} pathStr
 */
function getTextByPathStr(node, pathStr) {
    return PPTXUtils.getTextByPathList(node, pathStr.trim().split(/\s+/));
}

    /**
 * PPTXUtils.getTextByPathList
 * @param {Object} node
 * @param {string Array} path
 */

    /**
 * PPTXUtils.setTextByPathList
 * @param {Object} node
 * @param {string Array} path
 * @param {string} value
 */

    // ===== Color functions =====
    /**
 * applyShade
 * @param {string} rgbStr
 * @param {number} shadeValue
 */
function applyShade(rgbStr, shadeValue, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applyShade  color: ", color, ", shadeValue: ", shadeValue)
    if (shadeValue >= 1) {
        shadeValue = 1;
    }
    const cacl_l = Math.min(color.l * shadeValue, 1);//;color.l * shadeValue + (1 - shadeValue);
    // if (isAlpha)
    //     return color.lighten(tintValue).toHex8();
    // return color.lighten(tintValue).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

    /**
 * applyTint
 * @param {string} rgbStr
 * @param {number} tintValue
 */
function applyTint(rgbStr, tintValue, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applyTint  color: ", color, ", tintValue: ", tintValue)
    if (tintValue >= 1) {
        tintValue = 1;
    }
    const cacl_l = color.l * tintValue + (1 - tintValue);
    // if (isAlpha)
    //     return color.lighten(tintValue).toHex8();
    // return color.lighten(tintValue).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

    /**
 * applyLumOff
 * @param {string} rgbStr
 * @param {number} offset
 */
function applyLumOff(rgbStr, offset, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applyLumOff  color.l: ", color.l, ", offset: ", offset, ", color.l + offset : ", color.l + offset)
    let lum = offset + color.l;
    if (lum >= 1) {
        if (isAlpha)
            return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex8();
        return tinycolor({ h: color.h, s: color.s, l: 1, a: color.a }).toHex();
    }
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHex();
}

    /**
 * applyLumMod
 * @param {string} rgbStr
 * @param {number} multiplier
 */
function applyLumMod(rgbStr, multiplier, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applyLumMod  color.l: ", color.l, ", multiplier: ", multiplier, ", color.l * multiplier : ", color.l * multiplier)
    let cacl_l = color.l * multiplier;
    if (cacl_l >= 1) {
        cacl_l = 1;
    }
    if (isAlpha)
        return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: color.s, l: cacl_l, a: color.a }).toHex();
}

    // /**
    //  * applyHueMod
    //  * @param {string} rgbStr
    //  * @param {number} multiplier
    //  */
function applyHueMod(rgbStr, multiplier, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applyLumMod  color.h: ", color.h, ", multiplier: ", multiplier, ", color.h * multiplier : ", color.h * multiplier)

    let cacl_h = color.h * multiplier;
    if (cacl_h >= 360) {
        cacl_h = cacl_h - 360;
    }
    if (isAlpha)
        return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
    return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
}

    // /**
    //  * applyHueOff
    //  * @param {string} rgbStr
    //  * @param {number} offset
    //  */
    // function applyHueOff(rgbStr, offset, isAlpha) {
    //     const color = tinycolor(rgbStr).toHsl();
    //     //console.log("applyLumMod  color.h: ", color.h, ", offset: ", offset, ", color.h * offset : ", color.h * offset)

    //     const cacl_h = color.h * offset;
    //     if (cacl_h >= 360) {
    //         cacl_h = cacl_h - 360;
    //     }
    //     if (isAlpha)
    //         return tinycolor({ h: cocacl_h, s: color.s, l: color.l, a: color.a }).toHex8();
    //     return tinycolor({ h: cacl_h, s: color.s, l: color.l, a: color.a }).toHex();
    // }
    // /**
    //  * applySatMod
    //  * @param {string} rgbStr
    //  * @param {number} multiplier
    //  */
function applySatMod(rgbStr, multiplier, isAlpha) {
    const color = tinycolor(rgbStr).toHsl();
    //console.log("applySatMod  color.s: ", color.s, ", multiplier: ", multiplier, ", color.s * multiplier : ", color.s * multiplier)
    let cacl_s = color.s * multiplier;
    if (cacl_s >= 1) {
        cacl_s = 1;
    }
    //return;
    // if (isAlpha)
    //     return tinycolor(rgbStr).saturate(multiplier * 100).toHex8();
    // return tinycolor(rgbStr).saturate(multiplier * 100).toHex();
    if (isAlpha)
        return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex8();
    return tinycolor({ h: color.h, s: cacl_s, l: color.l, a: color.a }).toHex();
}

    /**
 * rgba2hex
 * @param {string} rgbaStr
 */
function rgba2hex(rgbaStr) {
    let a,
        rgb = rgbaStr.replace(/\s/g, '').match(/^rgba?\((\d+),(\d+),(\d+),?([^,\s)]+)?/i),
        alpha = (rgb && rgb[4] || "").trim(),
        hex = rgb ?
            (rgb[1] | 1 << 8).toString(16).slice(1) +
            (rgb[2] | 1 << 8).toString(16).slice(1) +
            (rgb[3] | 1 << 8).toString(16).slice(1) : rgbaStr;

    if (alpha !== "") {
        a = alpha;
    } else {
        a = 1;
    }
    // multiply before convert to HEX
    a = ((a * 255) | 1 << 8).toString(16).slice(1)
    hex = hex + a;

    return hex;
}

    ///////////////////////Amir////////////////
function getMiddleStops(s) {
    const sArry = ['0%', '100%'];
    if (s == 0) {
        return sArry;
    } else {
        let i = s;
        while (i--) {
            const middleStop = 100 - ((100 / (s + 1)) * (i + 1)), // AM: Ex - For 3 middle stops, progression will be 25%, 50%, and 75%, plus 0% and 100% at the ends.
                middleStopString = middleStop + "%";
            sArry.splice(-1, 0, middleStopString);
        } // AM: add into stopsArray before 100%
    }
    return sArry
}
function SVGangle(deg, svgHeight, svgWidth) {
    let w = parseFloat(svgWidth),
        h = parseFloat(svgHeight),
        ang = parseFloat(deg),
        o = 2,
        n = 2,
        wc = w / 2,
        hc = h / 2,
        tx1 = 2,
        ty1 = 2,
        tx2 = 2,
        ty2 = 2,
        k = (((ang % 360) + 360) % 360),
        j = (360 - k) * Math.PI / 180,
        i = Math.tan(j),
        l = hc - i * wc;

    if (k == 0) {
        tx1 = w,
            ty1 = hc,
            tx2 = 0,
            ty2 = hc
    } else if (k < 90) {
        n = w,
            o = 0
    } else if (k == 90) {
        tx1 = wc,
            ty1 = 0,
            tx2 = wc,
            ty2 = h
    } else if (k < 180) {
        n = 0,
            o = 0
    } else if (k == 180) {
        tx1 = 0,
            ty1 = hc,
            tx2 = w,
            ty2 = hc
    } else if (k < 270) {
        n = 0,
            o = h
    } else if (k == 270) {
        tx1 = wc,
            ty1 = h,
            tx2 = wc,
            ty2 = 0
    } else {
        n = w,
            o = h;
    }
    // AM: I could not quite figure out what m, n, and o are supposed to represent from the original code on visualcsstools.com.
    const m = o + (n / i);
    tx1 = tx1 == 2 ? i * (m - l) / (Math.pow(i, 2) + 1) : tx1;
    ty1 = ty1 == 2 ? i * tx1 + l : ty1;
    tx2 = tx2 == 2 ? w - tx1 : tx2;
    ty2 = ty2 == 2 ? h - ty1 : ty2;
    const x1 = Math.round(tx2 / w * 100 * 100) / 100,
        y1 = Math.round(ty2 / h * 100 * 100) / 100,
        x2 = Math.round(tx1 / w * 100 * 100) / 100,
        y2 = Math.round(ty1 / h * 100 * 100) / 100;
    return [x1, y1, x2, y2];
}
function getBase64ImageDimensions(imgSrc) {
    try {
        let base64Data = imgSrc;
        // 如果是以 data:image/ 开头的 URI，提取 base64 部分
        if (imgSrc && imgSrc.indexOf("data:image/") === 0) {
            base64Data = imgSrc.replace(/^data:image\/\w+;base64,/, '');
        }
        // 如果输入是 null、undefined 或不是字符串，返回 [0,0]
        if (!base64Data || typeof base64Data !== 'string') {
            return [0, 0];
        }
        // 移除可能的换行符和空格
        base64Data = base64Data.replace(/\s/g, '');
        // 检查字符串是否可能是有效的 base64（只包含 base64 字符）
        if (!/^[A-Za-z0-9+/]+=?=?$/.test(base64Data)) {
            return [0, 0];
        }
        // 解码 base64 为二进制字符串
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        
        // 检查 PNG 格式
        if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
            // PNG: IHDR 块起始于偏移 8，宽度在偏移 8+4 = 12，高度在偏移 16
            const width = (bytes[12] << 24) | (bytes[13] << 16) | (bytes[14] << 8) | bytes[15];
            const height = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            return [width, height];
        }
        
        // 检查 JPEG 格式
        if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
            let offset = 2;
            while (offset < bytes.length) {
                // 读取标记
                if (bytes[offset] !== 0xFF) break;
                const marker = bytes[offset + 1];
                // 帧开始标记 (SOF0)
                if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
                    // 高度在偏移 offset+5 (2字节)，宽度在偏移 offset+7 (2字节)
                    let height = (bytes[offset + 5] << 8) | bytes[offset + 6];
                    let width = (bytes[offset + 7] << 8) | bytes[offset + 8];
                    return [width, height];
                }
                // 跳转到下一个标记：标记长度是接下来的2字节（大端序）
                const length = (bytes[offset + 2] << 8) | bytes[offset + 3];
                offset += 2 + length;
            }
        }
    } catch (e) {
        // 发生错误时返回 [0,0]
    }
    
    // 默认返回 [0,0] 避免破坏现有代码
    return [0, 0];
}

const PPTXColorUtils = {
    getFillType,
    getGradientFill,
    getPicFill,
    getPatternFill,
    getLinerGrandient,
    getSolidFill,
    toHex,
    hslToRgb,
    hueToRgb,
    getColorName2Hex,
    getSchemeColorFromTheme,
    extractChartData,
    getTextByPathStr,
    applyShade,
    applyTint,
    applyLumOff,
    applyLumMod,
    applyHueMod,
    applySatMod,
    rgba2hex,
    getMiddleStops,
    SVGangle,
    getBase64ImageDimensions
}

export { PPTXColorUtils };

// Also export to global scope for backward compatibility

