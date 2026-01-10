/**
 * PPTXHtml - HTML 转换逻辑模块
 * 提取自 pptxjs.js
 */

(function () {
    var $ = window.jQuery;

    // 全局变量引用
    var PPTXUtils = window.PPTXUtils;
    var settings = window.settings; // 将在 pptxjs.js 中设置
    var PPTXParser = window.PPTXParser; // 从 PPTXParser 获取变量

    // 生成全局 CSS
    function genGlobalCSS() {
        var cssText = "";
        // 从 PPTXParser 获取 styleTable
        var styleTable = PPTXParser.styleTable || {};
        var slideWidth = PPTXParser.slideWidth || 960;
        for (var key in styleTable) {
            var tagname = "";
            // if (settings.slideMode && settings.slideType == "revealjs") {
            //     tagname = "section";
            // } else {
            //     tagname = "div";
            // }
            //ADD suffix
            cssText += tagname + " ." + styleTable[key]["name"] +
                ((styleTable[key]["suffix"]) ? styleTable[key]["suffix"] : "") +
                "{" + styleTable[key]["text"] + "}\n"; //section > div
        }
        //cssText += " .slide{margin-bottom: 5px;}\n"; // TODO

        if (settings.slideMode && settings.slideType == "divs2slidesjs") {
            //divId
            //console.log("slideWidth: ", slideWidth)
            cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + slideWidth + "px;}\n"; // TODO
        }
        return cssText;
    }

    // 生成表格 HTML
    function genTable(node, warpObj) {
        var order = node["attrs"]["order"];
        var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        /////////////////////////////////////////Amir////////////////////////////////////////////////
        var getTblPr = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
        var getColsGrid = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
        var tblDir = "";
        if (getTblPr !== undefined) {
            var isRTL = getTblPr["attrs"]["rtl"];
            tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
        }
        var firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
        var firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
        var lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
        var lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
        var bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
        var bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
        //console.log("getTblPr: ", getTblPr);
        var tblStylAttrObj = {
            isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttribute == "1") ? 1 : 0,
            isFrstColAttr: (firstColAttr !== undefined && firstColAttribute == "1") ? 1 : 0,
            isLstRowAttr: (lastRowAttr !== undefined && lastRowAttribute == "1") ? 1 : 0,
            isLstColAttr: (lastColAttr !== undefined && lastColAttribute == "1") ? 1 : 0,
            isBandRowAttr: (bandRowAttr !== undefined && bandRowAttribute == "1") ? 1 : 0,
            isBandColAttr: (bandColAttr !== undefined && bandColAttribute == "1") ? 1 : 0
        }

        var thisTblStyle;
        var tbleStyleId = getTblPr["a:tableStyleId"];
        if (tbleStyleId !== undefined) {
            // 简化版本，返回基本表格结构
            return "<div class='block table' style='z-index: " + order + ";'>[表格内容]</div>";
        }
        return "<div class='block table' style='z-index: " + order + ";'>表格</div>";
    }

    // 生成图表 HTML
    function genChart(node, warpObj) {
        var order = node["attrs"]["order"];
        var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
        var result = "<div id='chart" + chartID + "' class='block content' style='" +
            getPosition(xfrmNode, node, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";'></div>";

        var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
        var refName = warpObj["slideResObj"][rid]["target"];
        var content = readXmlFile(warpObj["zip"], refName);
        var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

        var chartData = null;
        for (var key in plotArea) {
            switch (key) {
                case "c:lineChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "lineChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:barChart":
                    chartData = {
                        "type": "createChart",
                        "data": {
                            "chartID": "chart" + chartID,
                            "chartType": "barChart",
                            "chartData": extractChartData(plotArea[key]["c:ser"])
                        }
                    };
                    break;
                case "c:pieChart":
                case "c:pie3DChart":
                    if (chartData.length > 0) {
                        data = chartData[0].values;
                    }
                    break;
                default:
            }
        }
        chartID++;
        return result;
    }

    // 生成图表数据
    function extractChartData(serNode) {
        // 简化版本
        return [];
    }

    // Convert plain numeric lists to proper HTML numbered lists
    function setNumericBullets(jqSelector) {
        jqSelector.find('li').each(function () {
            var $li = $(this);
            var html = $li.html();
            // If it starts with a number and a dot, treat as numbered list item
            if (/^\d+\.\s/.test(html)) {
                // Ensure parent is ol if not already
                var $parent = $li.parent();
                if (!$parent.is('ol')) {
                    $parent.each(function () {
                        if (!$(this).is('ol')) {
                            $(this).filter('ul').replaceWith(function () {
                                return $('<ol></ol>').append($(this).contents());
                            });
                        }
                    });
                }
            }
        });
    }

    // Process message queue and update UI accordingly
    function processMsgQueue(msgQueue) {
        if (!msgQueue || msgQueue.length === 0) return;

        // Example: could be used to show warnings, logs, or final status
        // Here we simply log and clear
        for (var i = 0; i < msgQueue.length; i++) {
            console.log("PPTXjs Message:", msgQueue[i]);
        }
        // Clear after processing
        msgQueue.length = 0;
    }

    // 处理单个消息
    function processSingleMsg(d) {
        var chartID = d.chartID;
        var chartType = d.chartType;
        var chartData = d.chartData;

        var data = [];

        var chart = null;
        switch (chartType) {
            case "lineChart":
                data = chartData;
                chart = nv.models.lineChart()
                    .useInteractiveGuideline(true);
                chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                break;
            case "barChart":
                data = chartData;
                chart = nv.models.multiBarChart();
                chart.xAxis.tickFormat(function (d) { return chartData[0].xlabels[d] || d; });
                break;
            case "pieChart":
            case "pie3DChart":
                if (chartData.length > 0) {
                    data = chartData[0].values;
                }
                break;
            default:
        }
        // 简化处理
    }

    // 获取背景
    function getBackground(warpObj, slideSize, index) {
        var bgResult = "";
        if (warpObj.processFullTheme === true) {
            // 读取 slide 节点中的背景
            var bgNode = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    // 纯色填充
                    var solidFill = getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = PPTXUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgResult = "<div class='slide-background-" + index + "' style='position:absolute;width:" + slideSize.width + "px;height:" + slideSize.height + "px;background-color:" + color + ";'></div>";
                        }
                    }
                    // 图片填充等可在此扩展
                }
            }
        }
        return bgResult;
    }

    // 获取幻灯片背景填充
    function getSlideBackgroundFill(warpObj, index) {
        var bgColor = "";
        if (warpObj.processFullTheme == "colorsAndImageOnly") {
            var bgNode = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:bg"]);
            if (bgNode) {
                var bgPr = bgNode["p:bgPr"];
                if (bgPr) {
                    var solidFill = getTextByPathList(bgPr, ["a:solidFill"]);
                    if (solidFill) {
                        var color = PPTXUtils.getFillColor(solidFill, warpObj.themeContent, warpObj.themeResObj, warpObj.slideLayoutClrOvride);
                        if (color) {
                            bgColor = "background-color:" + color + ";";
                        }
                    }
                }
            }
        }
        return bgColor;
    }

    // 更新加载进度条
    function updateProgressBar(percent) {
        var progressBarElemtnt = $(".slides-loading-progress-bar");
        progressBarElemtnt.width(percent + "%");
        progressBarElemtnt.html("<span style='text-align: center;'>Loading...(" + percent + "%)</span>");
    }

    // 公开 API
    window.PPTXHtml = {
        genGlobalCSS: genGlobalCSS,
        genTable: genTable,
        genChart: genChart,
        setNumericBullets: setNumericBullets,
        processMsgQueue: processMsgQueue,
        processSingleMsg: processSingleMsg,
        getBackground: getBackground,
        getSlideBackgroundFill: getSlideBackgroundFill,
        extractChartData: extractChartData
    };

})();