// Description: This script is used to get the number character based on the number format.

import { JSONPath } from '../vendor/jsonpath-plus/dist/index-browser-esm.js';

function IntToChineseCounting(nValue) {
    var sResult = '';
    var arrChinese = [
        String.fromCharCode(0x25CB),
        String.fromCharCode(0x4E00),
        String.fromCharCode(0x4E8C),
        String.fromCharCode(0x4E09),
        String.fromCharCode(0x56DB),
        String.fromCharCode(0x4E94),
        String.fromCharCode(0x516D),
        String.fromCharCode(0x4E03),
        String.fromCharCode(0x516B),
        String.fromCharCode(0x4E5D),
        String.fromCharCode(0x5341)
    ];

    var nQuotient = (nValue / 10) | 0;
    var nRemainder = nValue - nQuotient * 10;

    if (nQuotient < 10 && nQuotient > 0) {
        if (0 !== nRemainder)
            sResult = arrChinese[nRemainder] + sResult;

        sResult = arrChinese[10] + sResult;

        if (1 === nQuotient)
            nQuotient = 0;
    }
    else {
        sResult = arrChinese[nRemainder] + sResult;
    }


    var nRemValue = nQuotient;
    while (nQuotient > 0) {
        nQuotient = (nRemValue / 10) | 0;
        nRemainder = nRemValue - nQuotient * 10;

        sResult = arrChinese[nRemainder] + sResult;

        nRemValue = nQuotient;
    }
    return sResult;
}

function IntToChineseCountingThousand(nValue) {
    var sResult = '';
    var arrChinese = {
        0: String.fromCharCode(0x25CB),
        1: String.fromCharCode(0x4E00),
        2: String.fromCharCode(0x4E8C),
        3: String.fromCharCode(0x4E09),
        4: String.fromCharCode(0x56DB),
        5: String.fromCharCode(0x4E94),
        6: String.fromCharCode(0x516D),
        7: String.fromCharCode(0x4E03),
        8: String.fromCharCode(0x516B),
        9: String.fromCharCode(0x4E5D),
        10: String.fromCharCode(0x5341),
        100: String.fromCharCode(0x767E),
        1000: String.fromCharCode(0x5343),
        10000: String.fromCharCode(0x4E07)
    };
    if (nValue === 0) {
        sResult += arrChinese[0];
    }
    else if (nValue < 1000000) {
        var nRemValue = nValue;

        while (true) {
            var nTTQuotient = (nRemValue / 10000) | 0;
            var nTTRemainder = nRemValue - nTTQuotient * 10000;

            nRemValue = nTTQuotient;

            var sGroup = "", isPrevZero = false;

            if (nTTQuotient > 0)
                sGroup += arrChinese[10000];
            else
                isPrevZero = true;

            if (nTTRemainder <= 0) {
                sResult = sGroup + sResult;

                if (nRemValue <= 0)
                    break;

                continue;
            }

            var nQuotient = (nTTRemainder / 1000) | 0;
            var nRemainder = nTTRemainder - nQuotient * 1000;

            if (0 !== nQuotient) {
                sGroup += arrChinese[nQuotient] + arrChinese[1000];
                isPrevZero = false;
            }
            else if (nTTQuotient > 0) {
                sGroup += arrChinese[0];
                isPrevZero = true;
            }

            if (nRemainder <= 0) {
                sResult = sGroup + sResult;

                if (nRemValue <= 0)
                    break;

                continue;
            }

            nQuotient = (nRemainder / 100) | 0;
            nRemainder = nRemainder - nQuotient * 100;

            if (0 !== nQuotient) {
                sGroup += arrChinese[nQuotient] + arrChinese[100];
                isPrevZero = false;
            }
            else if (!isPrevZero) {
                sGroup += arrChinese[0];
                isPrevZero = true;
            }

            if (nRemainder <= 0) {
                sResult = sGroup + sResult;

                if (nRemValue <= 0)
                    break;

                continue;
            }

            nQuotient = (nRemainder / 10) | 0;
            nRemainder = nRemainder - nQuotient * 10;

            if (0 !== nQuotient) {
                if (nValue < 20)
                    sGroup += arrChinese[10];
                else
                    sGroup += arrChinese[nQuotient] + arrChinese[10];

                isPrevZero = false;
            }
            else if (!isPrevZero) {
                sGroup += arrChinese[0];
                isPrevZero = true;
            }

            if (0 !== nRemainder)
                sGroup += arrChinese[nRemainder];

            sResult = sGroup + sResult;

            if (nRemValue <= 0)
                break;
        }
    }
    return sResult;
}

let getNumChar = function (numFmt, no) {
    switch (numFmt) {
        case "decimal":
            return no;
        case "lowerLetter":
            return String.fromCharCode(97 + no);
        case "upperLetter":
            return String.fromCharCode(65 + no);
        case "lowerRoman":
            return toRoman(no).toLowerCase();
        case "upperRoman":
            return toRoman(no);
        case "chineseCounting":
            return IntToChineseCounting(no);
        case "chineseCountingThousand":
            return IntToChineseCountingThousand(no);
        default:
            return no;
    }
}


// convert 
// "$['content'][9]['content'][0]['content'][0]" => "$['content'][8]['content'][-1]['content'][-1]"
function prev_paragraph(path, step) {
    var ids = path.match(/\d+/g).map(function (item) {
        return parseInt(item);
    });

    var pIdx  = ids.length - 3;

    if (ids[pIdx] == 0 && step > 0) {
        return path;
    }

    ids[pIdx] -= step;
    for (var i = pIdx+1; i < ids.length; i++) {
        ids[i] = -1;
    }

    return path.replace(/\[\d+\]/g, function (match) {
        return "[" + ids.shift() + "]";
    });
}

// 返回同一个层级的最后一个元素
function last_paragraph(path) {
    var ids = path.match(/\d+/g).map(function (item) {
        return parseInt(item);
    });

    var pIdx  = ids.length - 3;

    for (var i = pIdx; i < ids.length; i++) {
        ids[i] = -1;
    }

    return path.replace(/\[\d+\]/g, function (match) {
        return "[" + ids.shift() + "]";
    });
}

function first_paragraph(path)  {
    return path.replace(/\[\d+\]$/, function (match) {
        return "[0]";
    });
}

// convert path to array 
// input "$['content'][18]['content'][0]['content'][0]"
// out: ["content", 18, "content", 0, "content", 0]
// just split the path by ']' or '[' or ']['
function toPathArray(path) {
    return path.match(/(?<=\[)['\w]+(?=\])/g).map(function (item) {
        if (item.startsWith("'"))
            return item.replace(/'/g, '');
        return parseInt(item);
    });
}

// convert array to path
// input: ["content", 18, "content", 0, "content", 0]
// out: "$['content'][18]['content'][0]['content'][0]"
// number dont need to quote
function toPathString(array) {
    return "$" + array.map(function (item) {
        return "[" + (typeof item === 'number' ? item : "'" + item + "'") + "]";
    }).join("");    
}


function is_less_path(left, right) {
    if (left == right) {
        return false;
    }

    var lefts = left.split(/\[\]/);
    var rights = right.split(/\[\]/);
    for (var i = 0; i < lefts.length; i++) {
        if (lefts[i] == rights[i]) {
            continue;
        }

        var left_idx = parseInt(lefts[i].match(/\d+/)[0]);
        var right_idx = parseInt(rights[i].match(/\d+/)[0]);
        return left_idx < right_idx;
    }

    return false;
}

function min_path(left, right) {
    if (is_less_path(left, right)) {
        return left;
    } 
    return right;
}

// return the major position of the path $.content[(\d+)]
function major_pos(path) {    
    var ret = path.match(/\d+/g);
    if (ret) {
        return parseInt(ret[0]);
    }
    return 0;    
}


// sibling时候，获取结束标记的前一个段落
// parent时，获取自己所在区间的最后一个段落
// child时，获取结束标记同级标签的前一个段落
function calc_relation(pathArrayA, pathArrayB) {
    if (pathArrayA.length < pathArrayB.length) {
        return "parent"
    }
    if (pathArrayA.length > pathArrayB.length) {
        return "child";
    }
    
    for (var i = 0; i < pathArrayA.length && i < pathArrayB.length;
         i++) {
        if (pathArrayA[i] != pathArrayB[i])
            break;
    }
    if (i == pathArrayA.length - 5) {
        return "sibling";
    }
    return "child";
}


let newSplit = function (text) {
    var k = JSON.parse(text);
    // 删除不需要的属性
    JSONPath({
        path: '$..[pPr,rPr,tblPr,tcPr]^', json: k, callback: function (node) {

            // check node whether has pPr property
            if (node.pPr) {
                if (node.pPr.numPr) {
                    node.pPr = { numPr: node.pPr.numPr, type: "paraPr" };
                } else {
                    delete node.pPr;
                }
            }
            delete node.changes;
            delete node.bFromDocument;
            delete node.rPr;
            delete node.endnotes;
            delete node.footnotes;
            delete node.reviewInfo;
            delete node.reviewType;
            delete node.tblPr;
            delete node.tcPr;
        }
    });

    // 删除图片
    JSONPath({
        path: '$..content[?(@.type=="paraDrawing")]', json: k, resultType: "value", callback: function (node) {
            delete node.graphic;
            delete node.effectExtent;
            delete node.positionV;
            delete node.positionH;
            delete node.cNvGraphicFramePr;
        }
    });

    // {
    //    numId: [0, 0, 0, 0,]
    // }
    var numberState = {}
    let getNum = function (numId, ilvl) {
        if (!numberState[numId]) {
            numberState[numId] = [];
        }
        if (!numberState[numId][ilvl]) {
            numberState[numId][ilvl] = 0;
        }
        var ret = numberState[numId][ilvl];
        numberState[numId][ilvl]++;
        return ret;
    }

    // 合并属性相同的run
    JSONPath({
        path: '$..content[?(@.type=="paragraph")]', json: k, resultType: "value", callback: function (node) {
            if (!node.content) {
                return;
            }

            // merge run
            var oldContent = node.content;
            node.content = [];
            oldContent.forEach(function (item) {
                if (node.content.length === 0 || item.type !== "run" || item.content.length === 0 || item.content[0].type) {
                    node.content.push(item);
                } else {
                    var last = node.content[node.content.length - 1];
                    if (last.type === item.type && last.content.length == 1 && item.content.length == 1 && !last.content[0].type) {
                        last.content[0] += item.content[0];
                    } else {
                        node.content.push(item);
                    }
                }
            });


            // numberings
            let genNum = function (node) {
                if (!node.pPr || !node.pPr.numPr) {
                    return;
                }


                var ilvl = node.pPr.numPr.ilvl;
                var numId = node.pPr.numPr.numId;

                if (!k.numbering.num[numId]) {
                    return;
                }

                var abstractNumId = k.numbering.num[numId].abstractNumId;
                if (!k.numbering.abstractNum[abstractNumId]) {
                    return;
                }

                var abstractNum = k.numbering.abstractNum[abstractNumId];
                if (!abstractNum.lvl || !abstractNum.lvl[ilvl]) {
                    return;
                }

                var lvl = abstractNum.lvl[ilvl];
                var numFmt = lvl.numFmt.val;
                var lvlText = lvl.lvlText;
                var no = (lvl.start ? lvl.start : 1) + getNum(numId, ilvl);
                var txt = lvlText.replace("%1", getNumChar(numFmt, no)) + "\t";


                var done = false;
                for (var i = 0; i < node.content.length; i++) {
                    var item = node.content[i];
                    if (item.type === "run" && item.content.length === 1 && !item.content[0].type) {
                        item.content[0] = txt + item.content[0];
                        done = true;
                        break;
                    }
                }
                if (done == false) {
                    node.content = [{ type: "run",
                        content: [txt]
                    }].concat(node.content);
                }
            };
            if (k.numbering !== undefined) {
                genNum(node);
            }

        }
    });

    // 替换所有全角字符
    JSONPath({
        path: '$..content[?(@.type == "run")].content[?(@.length)]', json: k, resultType: "all", callback: function (res) {
            var text = res.value.replace(/[\uFF10-\uFF19]|．|（|）/g, function (c) {
                if (c === '．') {
                    return '.';
                }
                if (c === '（') {
                    return '(';
                }
                if (c === '）') {
                    return ')';
                }

                return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
            });
            res.parent[res.parentProperty] = text;
        }
    });

    var ranges = [];
    // 获取结构区域    
    // $..content[?(typeof(@) == "string"  && @.match('^\\d+\\.'))] 
    const structPatt = "$..content[?(typeof(@) == 'string'  && @.match('[一二三四五六七八九十]+、'))]";
    JSONPath({
        path: structPatt, json: k, resultType: "path", callback: function (res) {
            ranges.push({
                beg: res,
                end: prev_paragraph(res, 0),
                info: { 'regionType': 'struct', 'mode': 1,  column: 1  },
                controlType: 1,
                major_pos: major_pos(res)
            })
        }
    });

    // 获取题目区域
    // $..content[?(typeof(@) == "string"  && @.match('^\\d+\\.'))] 
    var startIndex = ranges.length;
    const quesPatt = `$..content[?(typeof(@) == 'string' && @.match('^[0-9]+\.[^0-9]'))]`;
    JSONPath({
        path: quesPatt, json: k, resultType: "path", callback: function (res) {
            ranges.push({
                beg: res,
                info: { 'regionType': 'question', 'mode': 2, column: 1, padding: [0, 0, 0.5, 0] },
                controlType: 1,
                major_pos: major_pos(res)
            })
        }
    });

    ranges.sort(function (a, b) {
        return a.major_pos - b.major_pos;
    });

  
    for (var i = 0  ; i < ranges.length; i++) {
        var range = ranges[i];
        if (range.end) continue;

        if (i + 1 < ranges.length) {
            var next_range = ranges[i + 1];            
            // 如果下一个开始符号在同一个层级
            var boundary = next_range.beg;
            var thisA=toPathArray(range.beg);
            var nextA=toPathArray(boundary);

            var relation = calc_relation(thisA, nextA);
            if (relation == "sibling") {
                range.end = prev_paragraph(boundary, 1);
            } else if (relation == "parent") {
                // 如果下一个开始符号在低下一个层级 
                //array to string
                boundary=toPathString(nextA.slice(0,thisA.length));     
                range.end = prev_paragraph(boundary, 1);                
            } else {
                // 如果下一个开始符号在高一个层级
                range.end = last_paragraph(range.beg);
            }
        } else {
           range.end = `$['content'][-2]['content'][-1]['content'][-1]`;
        }
    }

    //console.log('k:', k)   
    //console.log('ranges', ranges)
    return ranges;
}

// define the function to convert the range to html
// window: the window object of the plugin
// range: the range object
// callback: the callback function to get the result
let rangeToHtml = function (window, range, callback) {
    Asc.scope.range = range;
    window.Asc.plugin.callCommand(function () {    
        var range = Asc.scope.range;
        // 如果range空，则获取当前选中的range
        var orange = range ? Api.asc_MakeRangeByPath(range.beg, range.end) : Api.GetDocument().GetRangeBySelect();

        orange.Select();

        let text_data = {
            data:     "",
            // 返回的数据中class属性里面有binary格式的dom信息，需要删除掉
            pushData: function (format, value) {
                this.data = value ? value.replace(/class="[a-zA-Z0-9-:;+"\/=]*/g, "") : "";
            }
        };

        Api.asc_CheckCopy(text_data, 2);
        return { methodType:"rangeToHtml", range: range, html: text_data.data};        
    },false, false, callback);
}

let insertHtml = function(window, pos, html, callback) {
    Asc.scope.insertPos = pos;
    Asc.scope.insertHtml = html;    
    if (html == null || html == undefined || html == "") {
        callback();
        return;
    }

    window.Asc.plugin.callCommand(function () {                
        console.log("insertHtml:", Asc.scope.insertHtml);
        
        var html = Asc.scope.insertHtml;
        Api.GetContentFromHtml(html, function(content) {
            console.log("content:", content);
            var arrContents = []
            for (var nElm = 0; nElm < content.Elements.length; nElm++) {
                var oElement = content.Elements[nElm];
                arrContents.push(oElement.Element);
            }
            Api.GetDocument().InsertContent(arrContents);
        });
    }, false, false, callback);
}

const _getNumChar = getNumChar;
export { _getNumChar as getNumChar, newSplit, rangeToHtml, insertHtml };

