/**
 * 节点工具函数模块
 * 
 * 处理 PPTX 节点的各种操作，包括：
 * - 幻灯片节点处理
 * - 图表生成
 * - SmartArt 图表处理
 * - 节点索引和查询
 * 
 * @module utils/node
 */

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';
import { PPTXTextUtils } from './text.js';
import { PPTXShapeUtils } from '../shape/shape.js';
import { genChart } from './chart.js';
import { SLIDE_FACTOR } from '../core/constants.js';

/**
 * 生成 Diagram HTML
 * @param {Object} node - 节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源类型
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function genDiagram(node, wrapObj, source, shapeType, settings) {
    const order = node.attrs.order;
    const zip = wrapObj.zip;
    const xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:xfrm']);
    const dgmRelIds = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'dgm:relIds', 'attrs']);
    const dgmClrFileId = dgmRelIds['r:cs'];
    const dgmDataFileId = dgmRelIds['r:dm'];
    const dgmLayoutFileId = dgmRelIds['r:lo'];
    const dgmQuickStyleFileId = dgmRelIds['r:qs'];

    const dgmClrFileName = wrapObj.slideResObj[dgmClrFileId].target;
    const dgmDataFileName = wrapObj.slideResObj[dgmDataFileId].target;
    const dgmLayoutFileName = wrapObj.slideResObj[dgmLayoutFileId].target;
    const dgmQuickStyleFileName = wrapObj.slideResObj[dgmQuickStyleFileId].target;

    const dgmClr = await PPTXXmlUtils.readXmlFile(zip, dgmClrFileName);
    const dgmData = await PPTXXmlUtils.readXmlFile(zip, dgmDataFileName);
    const dgmLayout = await PPTXXmlUtils.readXmlFile(zip, dgmLayoutFileName);
    const dgmQuickStyle = await PPTXXmlUtils.readXmlFile(zip, dgmQuickStyleFileName);

    const dgmDrwSpArray = PPTXXmlUtils.getTextByPathList(wrapObj.diagramContent, ['p:drawing', 'p:spTree', 'p:sp']);
    let result = '';

    if (dgmDrwSpArray !== undefined) {
        const results = [];
        for (const dspSp of dgmDrwSpArray) {
            const txBody = PPTXXmlUtils.getTextByPathList(dspSp, ['p:txBody', 'a:p', 'a:r', 'a:t']);
            results.push(processSpNode(dspSp, node, wrapObj, 'diagramBg', shapeType));
        }
        const resolvedResults = await Promise.all(results);
        result = resolvedResults.join('');
    }

    const position = PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined, shapeType);
    const size = PPTXXmlUtils.getSize(xfrmNode, undefined, undefined);

    return `<div class='block diagram-content' style='${position}${size}'>${result}</div>`;
}

/**
 * 索引幻灯片节点
 * @param {Object} content - 幻灯片内容
 * @returns {Object} 包含idTable、idxTable和typeTable的对象
 */
function indexNodes(content) {
    const keys = Object.keys(content);
    const spTreeNode = content[keys[0]]['p:cSld']['p:spTree'];

    const idTable = {};
    const idxTable = {};
    const typeTable = {};

    for (const key in spTreeNode) {
        if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') {
            continue;
        }

        const targetNode = spTreeNode[key];

        if (Array.isArray(targetNode)) {
            for (const node of targetNode) {
                const nvSpPrNode = node['p:nvSpPr'];
                const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
                const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
                const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);

                if (id !== undefined) idTable[id] = node;
                if (idx !== undefined) idxTable[idx] = node;
                if (type !== undefined) typeTable[type] = node;
            }
        } else {
            const nvSpPrNode = targetNode['p:nvSpPr'];
            const id = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id']);
            const idx = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx']);
            const type = PPTXXmlUtils.getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type']);

            if (id !== undefined) idTable[id] = targetNode;
            if (idx !== undefined) idxTable[idx] = targetNode;
            if (type !== undefined) typeTable[type] = targetNode;
        }
    }

    return { idTable, idxTable, typeTable };
}

/**
 * 处理组形状节点
 * @param {Object} node - 组形状节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processGroupSpNode(node, wrapObj, source, settings) {
    const xfrmNode = PPTXXmlUtils.getTextByPathList(node, ['p:grpSpPr', 'a:xfrm']);
    
    let groupStyle = '';
    let shapeType = 'group';
    let top, left, width, height;

    if (xfrmNode !== undefined) {
        const x = Math.round(parseInt(xfrmNode['a:off'].attrs.x) * SLIDE_FACTOR * 100) / 100;
        const y = Math.round(parseInt(xfrmNode['a:off'].attrs.y) * SLIDE_FACTOR * 100) / 100;

        // 根据ECMA-376标准，a:chOff和a:chExt是可选元素
        // 当不存在时，应该使用父元素的对应值作为默认值
        let childX, childY, childCx, childCy;

        if (xfrmNode['a:chOff'] !== undefined && xfrmNode['a:chOff'].attrs !== undefined) {
            childX = Math.round(parseInt(xfrmNode['a:chOff'].attrs.x) * SLIDE_FACTOR * 100) / 100;
            childY = Math.round(parseInt(xfrmNode['a:chOff'].attrs.y) * SLIDE_FACTOR * 100) / 100;
        } else {
            // 当a:chOff不存在时，使用a:off的值作为默认值
            childX = x;
            childY = y;
        }

        const cx = Math.round(parseInt(xfrmNode['a:ext'].attrs.cx) * SLIDE_FACTOR * 100) / 100;
        const cy = Math.round(parseInt(xfrmNode['a:ext'].attrs.cy) * SLIDE_FACTOR * 100) / 100;

        if (xfrmNode['a:chExt'] !== undefined && xfrmNode['a:chExt'].attrs !== undefined) {
            childCx = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cx) * SLIDE_FACTOR * 100) / 100;
            childCy = Math.round(parseInt(xfrmNode['a:chExt'].attrs.cy) * SLIDE_FACTOR * 100) / 100;
        } else {
            // 当a:chExt不存在时，使用a:ext的值作为默认值
            childCx = cx;
            childCy = cy;
        }

        const rotate = parseInt(xfrmNode.attrs.rot);
        let rotationStyle = '';
        
        top = y - childY;
        left = x - childX;
        width = cx - childCx;
        height = cy - childCy;

        if (!isNaN(rotate)) {
            const degrees = PPTXXmlUtils.angleToDegrees(rotate);
            rotationStyle = `transform: rotate(${degrees}deg); transform-origin: center;`;
            if (degrees !== 0) {
                top = y;
                left = x;
                width = cx;
                height = cy;
                shapeType = 'group-rotate';
            }
        }

        if (rotationStyle) groupStyle += rotationStyle;
    }

    if (top !== undefined) groupStyle += `top: ${top}px;`;
    if (left !== undefined) groupStyle += `left: ${left}px;`;
    if (width !== undefined) groupStyle += `width: ${width}px;`;
    if (height !== undefined) groupStyle += `height: ${height}px;`;

    const order = node.attrs.order;
    let result = `<div class='block group' style='z-index: ${order};${groupStyle}'>`;

    // Process all child nodes
    for (const nodeKey in node) {
        if (Array.isArray(node[nodeKey])) {
            for (const childNode of node[nodeKey]) {
                result += await processNodesInSlide(nodeKey, childNode, node, wrapObj, source, shapeType, settings);
            }
        } else {
            result += await processNodesInSlide(nodeKey, node[nodeKey], node, wrapObj, source, shapeType, settings);
        }
    }

    result += '</div>';
    return result;
}

/**
 * 处理幻灯片中的节点
 * @param {string} nodeKey - 节点键
 * @param {Object} nodeValue - 节点值
 * @param {Object} nodes - 节点集合
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processNodesInSlide(nodeKey, nodeValue, nodes, wrapObj, source, shapeType, settings) {
    switch (nodeKey) {
        case 'p:sp':    // Shape, Text
            return await processSpNode(nodeValue, nodes, wrapObj, source, shapeType, settings);
        case 'p:cxnSp':    // Shape, Text (with connection)
            return await processCxnSpNode(nodeValue, nodes, wrapObj, source, shapeType, settings);
        case 'p:pic':    // Picture
            return await processPicNode(nodeValue, wrapObj, source, shapeType, settings);
        case 'p:graphicFrame':    // Chart, Diagram, Table
            return await processGraphicFrameNode(nodeValue, wrapObj, source, shapeType, settings);
        case 'p:grpSp':
            return await processGroupSpNode(nodeValue, wrapObj, source, settings);
        case 'mc:AlternateContent': // Equations and formulas as Image
            const mcFallbackNode = PPTXXmlUtils.getTextByPathList(nodeValue, ['mc:Fallback']);
            return await processGroupSpNode(mcFallbackNode, wrapObj, source, settings);
        default:
            return '';
    }
}

/**
 * 处理形状节点
 * @param {Object} node - 形状节点
 * @param {Object} parentNode - 父节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'id']);
    const name = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name']);
    let idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
    let type = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
    const order = PPTXXmlUtils.getTextByPathList(node, ['attrs', 'order']);

    let isUserDrawnBg;
    if (source === 'slideLayoutBg' || source === 'slideMasterBg') {
        const userDrawn = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'attrs', 'userDrawn']);
        isUserDrawnBg = userDrawn === '1';
    }

    let slideLayoutSpNode;
    let slideMasterSpNode;

    if (idx !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.idxTable[idx];
        if (type !== undefined) {
            slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
        } else {
            slideMasterSpNode = wrapObj.slideMasterTables.idxTable[idx];
        }
    } else if (type !== undefined) {
        slideLayoutSpNode = wrapObj.slideLayoutTables.typeTable[type];
        slideMasterSpNode = wrapObj.slideMasterTables.typeTable[type];
    }

    if (type === undefined) {
        const txBoxVal = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox']);
        if (txBoxVal === '1') {
            type = 'textBox';
        }
    }

    if (type === undefined) {
        type = PPTXXmlUtils.getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
        if (type === undefined) {
            type = source === 'diagramBg' ? 'diagram' : 'obj';
        }
    }

    const result = await PPTXShapeUtils.genShape(node, parentNode, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, wrapObj, isUserDrawnBg, shapeType, source, settings);
    return result;
}

/**
 * 处理连接形状节点
 * @param {Object} node - 连接形状节点
 * @param {Object} parentNode - 父节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processCxnSpNode(node, parentNode, wrapObj, source, shapeType, settings) {
    const id = node['p:nvCxnSpPr']['p:cNvPr'].attrs.id;
    const name = node['p:nvCxnSpPr']['p:cNvPr'].attrs.name;
    const idx = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined 
        ? undefined 
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.idx;
    const type = node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined 
        ? undefined 
        : node['p:nvCxnSpPr']['p:nvPr']['p:ph'].attrs.type;
    const order = node.attrs.order;

    return await PPTXShapeUtils.genShape(node, parentNode, undefined, undefined, id, name, idx, type, order, wrapObj, undefined, shapeType, source, settings);
}

/**
 * 处理图片节点
 * @param {Object} node - 图片节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processPicNode(node, wrapObj, source, shapeType, settings) {
    const order = node.attrs.order;
    const rid = node['p:blipFill']['a:blip'].attrs['r:embed'];
    
    let resObj;
    if (source === 'slideMasterBg') {
        resObj = wrapObj.masterResObj;
    } else if (source === 'slideLayoutBg') {
        resObj = wrapObj.layoutResObj;
    } else {
        resObj = wrapObj.slideResObj;
    }
    
    const imgName = resObj[rid]?.target;

    if (imgName === undefined) {
        console.warn('Image reference not found in resObj for rid:', rid);
        return '';
    }

    const imgFileExt = PPTXXmlUtils.extractFileExtension(imgName).toLowerCase();
    const zip = wrapObj.zip;
    
    // 确定上下文类型用于路径解析
    let context = 'slide';
    if (source === 'slideMasterBg') {
        context = 'master';
    } else if (source === 'slideLayoutBg') {
        context = 'layout';
    }
    
    // 使用改进的媒体文件查找方法
    const imgFile = PPTXXmlUtils.findMediaFile(zip, imgName, context, '');
    if (imgFile === null) {
        console.warn('Image file not found in processPicNode:', imgName);
        return '';
    }
    
    const imgArrayBuffer = await imgFile.async("arraybuffer");
    let xfrmNode = node['p:spPr']?.['a:xfrm'];
    
    if (xfrmNode === undefined) {
        const idx = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx']);
        if (idx !== undefined) {
            xfrmNode = PPTXXmlUtils.getTextByPathList(wrapObj.slideLayoutTables, ['idxTable', idx, 'p:spPr', 'a:xfrm']);
        }
    }

    // 计算旋转角度
    let rotate = 0;
    const rotateNode = PPTXXmlUtils.getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot']);
    if (rotateNode !== undefined) {
        rotate = PPTXXmlUtils.angleToDegrees(rotateNode);
    }

    // 处理视频
    let mediaSupportFlag = false;
    let mediaPicFlag = false;
    let isVideoLink = false;
    let videoBlob, videoFile;
    
    const vdoNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile']);
    const mediaProcess = settings.mediaProcess;
    
    if (vdoNode !== undefined && mediaProcess) {
        const vdoRid = vdoNode.attrs['r:link'];
        videoFile = resObj[vdoRid].target;
        const checkIfLink = PPTXXmlUtils.IsVideoLink(videoFile);
        
        if (checkIfLink) {
            videoFile = PPTXXmlUtils.escapeHtml(videoFile);
            isVideoLink = true;
            mediaSupportFlag = true;
            mediaPicFlag = true;
        } else {
            const vdoFileExt = PPTXXmlUtils.extractFileExtension(videoFile).toLowerCase();
            if (['mp4', 'webm', 'ogg'].includes(vdoFileExt)) {
                const vdoFileObj = PPTXXmlUtils.findMediaFile(zip, videoFile, context, '');
                if (vdoFileObj === null) {
                    console.warn('Video file not found:', videoFile);
                } else {
                    const uInt8Array = await vdoFileObj.async("arraybuffer");
                    const vdoMimeType = PPTXXmlUtils.getMimeType(vdoFileExt);
                    const blob = new Blob([uInt8Array], { type: vdoMimeType });
                    videoBlob = URL.createObjectURL(blob);
                    mediaSupportFlag = true;
                    mediaPicFlag = true;
                }
            }
        }
    }

    // 处理音频
    let audioPlayerFlag = false;
    let audioBlob;
    let audioObj;
    
    const audioNode = PPTXXmlUtils.getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile']);
    
    if (audioNode !== undefined && mediaProcess) {
        const audioRid = audioNode.attrs['r:link'];
        const audioFile = resObj[audioRid].target;
        const audioFileExt = PPTXXmlUtils.extractFileExtension(audioFile).toLowerCase();
        
        if (['mp3', 'wav', 'ogg'].includes(audioFileExt)) {
            const audioFileObj = PPTXXmlUtils.findMediaFile(zip, audioFile, context, '');
            if (audioFileObj === null) {
                console.warn('Audio file not found:', audioFile);
            } else {
                const uInt8ArrayAudio = await audioFileObj.async("arraybuffer");
                const blobAudio = new Blob([uInt8ArrayAudio]);
                audioBlob = URL.createObjectURL(blobAudio);
                
                const cx = parseInt(xfrmNode['a:ext'].attrs.cx) * 20;
                const cy = parseInt(xfrmNode['a:ext'].attrs.cy);
                const x = parseInt(xfrmNode['a:off'].attrs.x) / 2.5;
                const y = parseInt(xfrmNode['a:off'].attrs.y);
                
                audioObj = {
                    'a:ext': { attrs: { cx, cy } },
                    'a:off': { attrs: { x, y } }
                };
                
                audioPlayerFlag = true;
                mediaSupportFlag = true;
                mediaPicFlag = true;
            }
        }
    }

    const mimeType = PPTXXmlUtils.getMimeType(imgFileExt);
    const position = mediaProcess && audioPlayerFlag 
        ? PPTXXmlUtils.getPosition(audioObj, node, undefined, undefined)
        : PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined);
    const size = mediaProcess && audioPlayerFlag
        ? PPTXXmlUtils.getSize(audioObj, undefined, undefined)
        : PPTXXmlUtils.getSize(xfrmNode, undefined, undefined);

    let result = `<div class='block content' style='${position}${size} z-index: ${order};transform: rotate(${rotate}deg);'>`;
    
    if ((vdoNode === undefined && audioNode === undefined) || !mediaProcess || !mediaSupportFlag) {
        const base64Data = PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer);
        result += `<img src='data:${mimeType};base64,${base64Data}' style='width: 100%; height: 100%'/>`;
    } else if ((vdoNode !== undefined || audioNode !== undefined) && mediaProcess && mediaSupportFlag) {
        if (vdoNode !== undefined && !isVideoLink) {
            result += `<video src='${videoBlob}' controls style='width: 100%; height: 100%'>Your browser does not support the video tag.</video>`;
        } else if (vdoNode !== undefined && isVideoLink) {
            result += `<iframe src='${videoFile}' controls style='width: 100%; height: 100%'></iframe>`;
        }
        if (audioNode !== undefined) {
            result += `<audio id="audio_player" controls><source src="${audioBlob}"></audio>`;
        }
    }
    
    if (!mediaSupportFlag && mediaPicFlag) {
        result += `<span style='color:red;font-size:40px;position: absolute;'>This media file Not supported by HTML5</span>`;
    }
    
    result += '</div>';
    return result;
}

/**
 * 处理图形框架节点
 * @param {Object} node - 图形框架节点
 * @param {Object} wrapObj - 包装对象
 * @param {string} source - 源
 * @param {string} shapeType - 形状类型
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 生成的HTML
 */
async function processGraphicFrameNode(node, wrapObj, source, shapeType, settings) {
    const graphicTypeUri = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri']);

    switch (graphicTypeUri) {
        case 'http://schemas.openxmlformats.org/drawingml/2006/table':
            return await PPTXTextUtils.genTable(node, wrapObj);
        case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
            return await genChart(node, wrapObj);
        case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
            return await genDiagram(node, wrapObj, source, shapeType, settings);
        case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
            let oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj']);
            if (oleObjNode === undefined) {
                oleObjNode = PPTXXmlUtils.getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj']);
            }
            if (oleObjNode !== undefined) {
                return await processGroupSpNode(oleObjNode, wrapObj, source, settings);
            }
            return '';
        default:
            return '';
    }
}

/**
 * 处理形状属性节点
 * @param {Object} node - 形状属性节点
 * @param {Object} wrapObj - 包装对象
 */
function processSpPrNode(node, wrapObj) {
    // TODO: Implement shape properties processing
}

/**
 * 获取幻灯片背景
 * @param {Object} wrapObj - 包装对象
 * @param {Object} slideSize - 幻灯片尺寸
 * @param {number} index - 幻灯片索引
 * @param {Object} settings - 设置对象
 * @returns {Promise<string>} 背景HTML
 */
async function getBackground(wrapObj, slideSize, index, settings) {
    const slideContent = wrapObj.slideContent;
    const slideLayoutContent = wrapObj.slideLayoutContent;
    const slideMasterContent = wrapObj.slideMasterContent;

    const nodesSldLayout = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree']);
    const nodesSldMaster = PPTXXmlUtils.getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree']);
    const showMasterSp = PPTXXmlUtils.getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp']);
    
    const bgColor = await PPTXStyleUtils.getSlideBackgroundFill(wrapObj, index);
    let result = `<div class='slide-background-${index}' style='width:${slideSize.width}px; height:${slideSize.height}px;${bgColor}'>`;

    if (nodesSldLayout !== undefined) {
        for (const nodeKey in nodesSldLayout) {
            if (Array.isArray(nodesSldLayout[nodeKey])) {
                for (const node of nodesSldLayout[nodeKey]) {
                    const phType = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    if (phType !== 'pic') {
                        result += await processNodesInSlide(nodeKey, node, nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings);
                    }
                }
            } else {
                const phType = PPTXXmlUtils.getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                if (phType !== 'pic') {
                    result += await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], nodesSldLayout, wrapObj, 'slideLayoutBg', 'group', settings);
                }
            }
        }
    }
    
    if (nodesSldMaster !== undefined && (showMasterSp === '1' || showMasterSp === undefined)) {
        for (const nodeKey in nodesSldMaster) {
            if (Array.isArray(nodesSldMaster[nodeKey])) {
                for (const node of nodesSldMaster[nodeKey]) {
                    const phType = PPTXXmlUtils.getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type']);
                    result += await processNodesInSlide(nodeKey, node, nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings);
                }
            } else {
                result += await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], nodesSldMaster, wrapObj, 'slideMasterBg', 'group', settings);
            }
        }
    }
    
    return result;
}

// 创建兼容的 PPTXNodeUtils 对象
const PPTXNodeUtils = {
    indexNodes,
    processGroupSpNode,
    processNodesInSlide,
    processSpNode,
    processCxnSpNode,
    processPicNode,
    processGraphicFrameNode,
    processSpPrNode,
    getBackground,
    genDiagram
};

export { PPTXNodeUtils };
export default PPTXNodeUtils;
