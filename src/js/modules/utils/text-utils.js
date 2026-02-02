/**
 * 文本工具函数模块
 * 提供文本处理、编号转换等功能
 */

/**
 * alphaNumeric - 将数字转换为字母编号
 * @param {number} num - 数字
 * @param {string} upperLower - "upperCase" 或 "lowerCase"
 * @returns {string} 字母编号
 */

var PPTXTextUtils = (function() {
    function alphaNumeric(num, upperLower) {
    num = Number(num) - 1;
    let aNum = "";
    if (upperLower == "upperCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
    } else if (upperLower == "lowerCase") {
        aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
    }
    return aNum;
}

/**
 * romanize - 将数字转换为罗马数字
 * @param {number} num - 数字
 * @returns {string} 罗马数字字符串
 */
    function romanize(num) {
    if (!+num)
        return false;
    const digits = String(+num).split(""),
        key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
            "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
            "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
        roman = "",
        i = 3;
    while (i--)
        roman = (key[+digits.pop() + (i * 10)] || "") + roman;
    return Array(+digits.join("") + 1).join("M") + roman;
}

/**
 * archaicNumbers - 创建古老数字格式化器（如希伯来数字）
 * @param {Array} arr - 数字和对应字符的数组
 * @returns {Object} 格式化器对象
 */
    function archaicNumbers(arr) {
    const arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length });
    return {
        format: function (n) {
            let ret = '';
            jQuery.each(arr, function () {
                const num = this[0];
                if (parseInt(num) > 0) {
                    for (; n >= num; n -= num) ret += this[1];
                } else {
                    ret = ret.replace(num, this[1]);
                }
            });
            return ret;
        }
    };
}

/**
 * 希伯来数字格式化器
 */
    const hebrew2Minus = archaicNumbers([
    [1000, ''],
    [400, 'ת'],
    [300, 'ש'],
    [200, 'ר'],
    [100, 'ק'],
    [90, 'צ'],
    [80, 'פ'],
    [70, 'ע'],
    [60, 'ס'],
    [50, 'נ'],
    [40, 'מ'],
    [30, 'ל'],
    [20, 'כ'],
    [10, 'י'],
    [9, 'ט'],
    [8, 'ח'],
    [7, 'ז'],
    [6, 'ו'],
    [5, 'ה'],
    [4, 'ד'],
    [3, 'ג'],
    [2, 'ב'],
    [1, 'א'],
    [/יה/, 'ט״ו'],
    [/יו/, 'ט״ז'],
    [/([א-ת])([א-ת])$/, '$1״$2'],
    [/^([א-ת])$/, "$1׳"]
]);

/**
 * getNumTypeNum - 根据编号类型获取编号字符串
 * @param {string} numTyp - 编号类型
 * @param {number} num - 数字
 * @returns {string} 编号字符串
 */
    function getNumTypeNum(numTyp, num) {
    let rtrnNum = "";
    switch (numTyp) {
        case "arabicPeriod":
            rtrnNum = num + ". ";
            break;
        case "arabicParenR":
            rtrnNum = num + ") ";
            break;
        case "alphaLcParenR":
            rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
            break;
        case "alphaLcPeriod":
            rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
            break;
        case "alphaUcParenR":
            rtrnNum = alphaNumeric(num, "upperCase") + ") ";
            break;
        case "alphaUcPeriod":
            rtrnNum = alphaNumeric(num, "upperCase") + ". ";
            break;
        case "romanUcPeriod":
            rtrnNum = romanize(num) + ". ";
            break;
        case "romanLcParenR":
            rtrnNum = romanize(num) + ") ";
            break;
        case "hebrew2Minus":
            rtrnNum = hebrew2Minus.format(num) + "-";
            break;
        default:
            rtrnNum = num;
    }
    return rtrnNum;
}

/**
 * setNumericBullets - 设置数字项目符号
 * @param {Array} elem - 段落元素数组
 */
    function setNumericBullets(elem) {
    const prgrphs_arry = elem;
    for (let i = 0; i < prgrphs_arry.length; i++) {
        const buSpan = $(prgrphs_arry[i]).find('.numeric-bullet-style');
        if (buSpan.length > 0) {
            let prevBultTyp = "";
            let prevBultLvl = "";
            let buletIndex = 0;
            const tmpArry = [];
            let tmpArryIndx = 0;
            const buletTypSrry = [];

            for (let j = 0; j < buSpan.length; j++) {
                const bult_typ = $(buSpan[j]).data("bulltname");
                const bult_lvl = $(buSpan[j]).data("bulltlvl");

                if (buletIndex == 0) {
                    prevBultTyp = bult_typ;
                    prevBultLvl = bult_lvl;
                    tmpArry[tmpArryIndx] = buletIndex;
                    buletTypSrry[tmpArryIndx] = bult_typ;
                    buletIndex++;
                } else {
                    if (bult_typ == prevBultTyp && bult_lvl == prevBultLvl) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        buletIndex++;
                        tmpArry[tmpArryIndx] = buletIndex;
                        buletTypSrry[tmpArryIndx] = bult_typ;
                    } else if (bult_typ != prevBultTyp && bult_lvl == prevBultLvl) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        tmpArryIndx++;
                        tmpArry[tmpArryIndx] = buletIndex;
                        buletTypSrry[tmpArryIndx] = bult_typ;
                        buletIndex = 1;
                    } else if (bult_typ != prevBultTyp && Number(bult_lvl) > Number(prevBultLvl)) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        tmpArryIndx++;
                        tmpArry[tmpArryIndx] = buletIndex;
                        buletTypSrry[tmpArryIndx] = bult_typ;
                        buletIndex = 1;
                    } else if (bult_typ != prevBultTyp && Number(bult_lvl) < Number(prevBultLvl)) {
                        prevBultTyp = bult_typ;
                        prevBultLvl = bult_lvl;
                        tmpArryIndx--;
                        buletIndex = tmpArry[tmpArryIndx] + 1;
                    }
                }

                const numIdx = getNumTypeNum(buletTypSrry[tmpArryIndx], buletIndex);
                $(buSpan[j]).html(numIdx);
            }
        }
    }
}


    return {
        alphaNumeric: alphaNumeric,
        romanize: romanize,
        archaicNumbers: archaicNumbers,
        hebrew2Minus: hebrew2Minus,
        getNumTypeNum: getNumTypeNum,
        setNumericBullets: setNumericBullets
    };
})();