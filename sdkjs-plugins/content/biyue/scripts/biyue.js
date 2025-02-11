import { getNumChar, newSplit, rangeToHtml, insertHtml, normalizeDoc } from "./dep.js";
import {
    getPaperInfo, initPaperInfo, updateCustomControls, clearStruct, getStruct, savePositons, showQuestionTree, updateQuestionScore, drawPosition, addQuesScore, addScoreField, delScoreField, changeScoreField,
    addImage, addMarkField, handleContentControlChange
} from './business.js'
import { toXml } from "./convert.js";
import { biyueCallCommand } from "./command.js";
import { getVersion } from "./ver.js"
import { ReplaceRubyField } from "./phonetic.js";

(function (window, undefined) {
    var styleEnable = false;
    var settingsWindow = null;

    function NewDefaultCustomData() {
        return {
            controlDesc: {}
        }
    }

    let splitQuestion = function (text_all, text_pos) {
        var index = 1;

        // replace uFF10-FF19 to 0-9
        text_all = text_all.replace(/[\uFF10-\uFF19]|．|（|）/g, function (c) {
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

        // 匹配 格式一、的题组结构
        var structPatt = /[一二三四五六七八九十]+、.*?(?=((\n|\r)[\d]+\.?)|((\n|\r)[一二三四五六七八九十]+、)|(?:$|[\r\n]))/gs
        // 匹配 格式1.的题目 同时避开结构
        var quesPatt = /(?<=^|\r|\n)[ ]?[\d]+[\.]?.*?(\n|\r).*?(?=((\n|\r)[ ]?[\d]+[\.]?)|((\n|\r)[一二三四五六七八九十]+、)|$)/gs
        // 匹配 批改作答区域
        var rangePatt = /(([\(]|[\（])(\s|\&nbsp\;)*([\）]|[\)]))|(___*)/gs
        //var inlineQuesPatt = /(?<=^|\r|\n|\t| )\d+[.].*?(?=(\d+\.)|\r|$)/gs;
        var isInlinePatt = /(?<=^|\r|\n)(\d+\..*)([ ]+\d+\..*)+/g;
        var inlineQuesPatt = /\(?\d+[\).].*?(?=(\d+\.)|(\(\d+\))|\r|$)/gs;


        var structTextArr = text_all.match(structPatt) || []
        var quesTextArr = text_all.match(quesPatt) || []

        //structTextArr = structTextArr.map(text => text.replace(/[\n]/g, ''));
        //quesTextArr = quesTextArr.map(text => text.replace(/[\n]/g, ''));

        //text_all = text_all.replace(/[\n]/g, '')
        // debugger
        var ranges = new Array();
        let info = {}
        structTextArr.map(item => {
            // 结构 例如：一、选择题
            var startIndex = text_all.indexOf(item);
            var endIndex = startIndex + item.length;
            info = { 'regionType': 'struct', 'mode': 1 }
            ranges.push({ beg: text_pos[startIndex], end: text_pos[endIndex], controlType: 1, info: info });
        });
        let no = 1
        quesTextArr.map(item => {
            // 题目 例如：1.xxx
            var startIndex = text_all.indexOf(item);
            var endIndex = startIndex + item.length;



            var column = 1;
            var isInlineArr = item.match(isInlinePatt) || []
            if (isInlineArr.length >= 1) {
                // debugger

                var r = isInlineArr[0].match(inlineQuesPatt);
                if (r !== undefined && r !== null) {
                    column = r.length;
                    console.log('column:', column, " of ques", no)
                }
            }


            info = { 'ques_no': no, 'regionType': 'question', 'mode': 2, padding: [0, 0, 0.5, 0], color: "" }
            ranges.push({ beg: text_pos[startIndex], end: text_pos[endIndex], controlType: 1, info: info, column: column });
            no++

            // 匹配题目里下划线和括号之类的批改/作答区域
            const processedIndexes = new Set();
            let match;

            info = { 'regionType': 'write', 'color': '#ffcccc' }
            while ((match = rangePatt.exec(item)) !== null) {
                const startPos = startIndex + match.index;
                const endPos = startIndex + match.index + match[0].length;

                if (!processedIndexes.has(startPos)) {
                    ranges.push({ beg: text_pos[startPos], end: text_pos[endPos], controlType: 2, info: info });
                    for (let i = startPos; i < endPos; i++) {
                        processedIndexes.add(i);
                    }
                }
            }
        });
        console.log('ranges:', ranges)
        return ranges;
    }

    // onlyoffice插件的api，执行command在callback里面继续执行command会失败
    // 用一个task stask来执行command
    let setupPostTask = function (window, task) {
        window.postTask = window.postTask || [];
        window.postTask.push(task);
    }

    let execPostTask = function (window, param) {
        while (window.postTask && window.postTask.length > 0) {
            var task = window.postTask.pop();
            var imm = task(param);
            if (imm !== true) {
                break;
            }
        }
    }

    let checkSubQuestion = function () {        
        return biyueCallCommand(window, function () {
            console.log("checkSubQuestion start.");
            // 通用匹配函数，在range对patt进行匹配，返回一个数组，数组的每个元素是一个数组都是对应的ApiRange
            let RangeMatch = function (range, patt) {
                let marker_log = function (str, ranges) {
                    let styledString = '';
                    let currentIndex = 0;
                    const styles = [];

                    ranges.forEach(([start, end], index) => {
                        // 添加高亮前的部分
                        if (start > currentIndex) {
                            styledString += '%c' + str.substring(currentIndex, start);
                            styles.push('');
                        }
                        // 添加高亮部分
                        styledString += '%c' + str.substring(start, end);
                        styles.push('border: 1px solid red; padding: 2px');
                        currentIndex = end;
                    });

                    // 添加剩余的部分
                    if (currentIndex < str.length) {
                        styledString += '%c' + str.substring(currentIndex);
                        styles.push('');
                    }

                    console.log(styledString, ...styles);
                };

                function CalcTextPos(text_all, text_plain) {
                    text_plain = text_plain.replace(/[\r]/g, '');
                    var text_pos = new Array(text_all.length);
                    var j = 0;
                    for (var i = 0, n = text_plain.length; i < n; i++) {
                        while (text_all[j] !== text_plain[i]) {
                            text_pos[j] = i;
                            j++;
                        }
                        text_pos[j] = i;
                        j++;
                    }
                    return text_pos;
                }

                // 用正则表达式实现
                // 自定义位置
                var text = range.GetText({ Math: false });
                var text_plain = range.GetText({ Math: false, Numbering: false });
                var text_pos = CalcTextPos(text, text_plain);

                var match;
                var matchRanges = [];
                var ranges = []
                while ((match = patt.exec(text)) !== null) {
                    var begPos = text_pos[match.index];
                    var endPos = text_pos[match.index + match[0].length];
                    ranges.push([match.index, match.index + match.length]);
                    matchRanges.push(range.GetRange(begPos, endPos));
                }
                if (ranges.length > 0) {
                    marker_log(text, ranges);
                }
                return matchRanges;
            }

            var oDocument = Api.GetDocument();
            var controls = oDocument.GetAllContentControls();
            controls.forEach(control => {
                var subQuesPatt = /(?<=^|\r|\n|\t)(\(|\（)\d+(\)|\）)/gs;
                var apiRanges = RangeMatch(control.GetRange(), subQuesPatt);
                if (apiRanges.length > 1) {
                    //debugger;
                    for (var i = 0; i < apiRanges.length; i++) {
                        var endRange = undefined;
                        if (i < apiRanges.length - 1) {

                            var nextRange = apiRanges[i + 1];
                            endRange = nextRange.GetParagraph(0).GetPrevious().GetRange();
                        } else {
                            // debugger;
                            var content = control.GetContent();
                            var endParaIndex = content.GetElementsCount() - 1;
                            var oPara = content.GetElement(endParaIndex);
                            endRange = oPara.GetRange();
                        }

                        var range = apiRanges[i].ExpandTo(endRange);
                        range.Select();
                        if (range !== undefined) {
                            var tag = JSON.stringify({ 'regionType': 'sub-question', 'mode': 3 });
                            Api.asc_AddContentControl(1, { "Tag": tag });
                            Api.asc_RemoveSelection();
                        }
                    }
                }

            });
            console.log("checkSubQuestion done.");
        }, false, false);
    }

    let checkAnswerRegion = function () {
        
        return biyueCallCommand(window, function () {
            // 在console中打印字符串，range指定的部分会被高亮显示
            let marker_log = function (str, ranges) {
                let styledString = '';
                let currentIndex = 0;
                const styles = [];

                ranges.forEach(([start, end], index) => {
                    // 添加高亮前的部分
                    if (start > currentIndex) {
                        styledString += '%c' + str.substring(currentIndex, start);
                        styles.push('');
                    }
                    // 添加高亮部分
                    styledString += '%c' + str.substring(start, end);
                    styles.push('border: 1px solid red; padding: 2px');
                    currentIndex = end;
                });

                // 添加剩余的部分
                if (currentIndex < str.length) {
                    styledString += '%c' + str.substring(currentIndex);
                    styles.push('');
                }

                console.log(styledString, ...styles);
            };

            // debugger;

            var oDocument = Api.GetDocument();
            var controls = oDocument.GetAllContentControls();

            for (var i = 0; i < controls.length; i++) {
                var control = controls[i];
                var obj = ''
                if (control && control.GetTag()) {
                    obj = control.GetTag() || ''
                    if (obj) {
                        try {
                            obj = JSON.parse(obj)
                        } catch (e) {
                            console.error('JSON解析失败', e)
                        }
                    }
                }
                if (obj.regionType === 'question' || obj.regionType === 'sub-question') {
                    var inlineSdts = control.GetAllContentControls().filter(e => e.GetTag() == JSON.stringify({ 'regionType': 'write', 'mode': 3 }));
                    if (inlineSdts.length > 0) {
                        console.log("已有inline sdt， 删除以后再执行", inlineSdts);
                        continue
                    }


                    // 标记inline的答题区域
                    var text = control.GetRange().GetText();
                    var rangePatt = /(([\(]|[\（])(\s|\&nbsp\;)*([\）]|[\)]))|(___*)/gs
                    var match;
                    var ranges = [];
                    var regionTexts = [];
                    while ((match = rangePatt.exec(text)) !== null) {
                        ranges.push([match.index, match.index + match[0].length]);
                        regionTexts.push(match[0]);
                    }
                    if (ranges.length > 0) {
                        marker_log(text, ranges);
                    }
                    var textSet = new Set();
                    regionTexts.forEach(e => textSet.add(e));

                    let includeRange = function(a, b)
                    {
                        return (a.Element === b.Element &&
                            a.Start >= b.Start &&
                            a.End <= b.End);
                    };
                    let mergeRange = function(arrA, arrB)
                    {
                        console.time("mergeRange");
                        let all = arrA.concat(arrB);
                        // Sort ranges by start position
                        all.sort(function(a, b) {
                            var len = Math.min(a.StartPos.length, b.StartPos.length);
                            for (var i = 0; i< len; i++) {
                                if (a.StartPos[i].Position != b.StartPos[i].Position)
                                    return a.StartPos[i].Position - b.StartPos[i].Position
                            }
                            return 0;
                        });

                        let ret = []
                        let last = null;
                        for (let range of all) {
                            if (last === null || last.End < range.Start || last.Element != range.Element) {
                                // No overlap, add to result
                                ret.push(range);
                                last = range;
                            } else if (last.End < range.End) {
                                // Overlapping ranges, merge them
                                last.End = range.End;
                            }
                        }
                        console.timeEnd("mergeRange");
                        return ret;
                    };
                                        
                    
                    var apiRanges = [];
                    console.time("查找所有答题区 no: "+i);
                    textSet.forEach(e => {
                        var ranges = control.Search(e, false);
                        //debugger;;
                        
                        apiRanges = mergeRange(apiRanges, ranges);
                    });
                    console.timeEnd("查找所有答题区 no: "+i);
                    

                    // search 有bug少返回一个字符            
                    console.time("插入答题区控件");
                    apiRanges.reverse().forEach(apiRange => {
                            apiRange.Select();
                            var tag = JSON.stringify({ 'regionType': 'write', 'mode': 3 });
                            Api.asc_AddContentControl(2, { "Tag": tag });
                            Api.asc_RemoveSelection();
                    });
                    console.timeEnd("插入答题区控件");

                    
                    // 标记空白行
                    {
                        //debugger;
                        var content = control.GetContent();
                        var elements = content.GetElementsCount();
                        for (var j = elements - 1; j >= 0; j--) {
                            var para = content.GetElement(j);
                            if (para.GetClassType() !== "paragraph") {
                                break;
                            }
                            var text = para.GetText();
                            if (text.trim() !== '') {
                                break;
                            }
                        }

                        if (j < elements - 1) {
                            var range = content.GetElement(j + 1).GetRange();
                            var endRange = content.GetElement(elements - 1).GetRange();
                            range = range.ExpandTo(endRange);
                            range.Select();
                            var tag = JSON.stringify({ 'regionType': 'write', 'mode': 5 });
                            Api.asc_AddContentControl(1, { "Tag": tag });
                            Api.asc_RemoveSelection();
                        }
                    }
                }
                
            }
            
        }, false, false);
    }

    // 插件初始化    
    window.Asc.plugin.init = function () {
        console.log("biyue plugin inited.");

        // create style        
        if (window.BiyueCustomData === undefined) {
            this.callCommand(
                function () {
                    var oDocument = Api.GetDocument();
                    var customData = Api.asc_GetBiyueCustomDataExt(undefined);
                    if (customData === undefined || customData.length === 0)
                        return undefined
                    return customData;
                },
                false,
                false,
                function (customData) {
                    console.log("customData", customData);
                    setTimeout(() => {
                        GetDocInfo()
                    }, 1000)
                    if (customData === undefined) {
                        console.log("customData inited.")
                        window.BiyueCustomData = NewDefaultCustomData();
                        return;
                    }
                    window.BiyueCustomId = customData[0].ItemId;
                    window.BiyueCustomData = customData[0].Content;
                }
            );
        }
    };

    function StoreCustomData(callback) {
        if (window.BiyueCustomData === undefined) {
            callback();
            return;
        }

        Asc.scope.BiyueCustomId = window.BiyueCustomId;
        Asc.scope.BiyueCustomData = window.BiyueCustomData;
        window.Asc.plugin.callCommand(
            function () {
                var id = Asc.scope.BiyueCustomId
                var data = Asc.scope.BiyueCustomData;
                console.log("store custom data");
                return Api.asc_SetBiyueCustomDataExt(id, data);                
            },
            false,
            false,
            callback
        );
    }

    function getContextMenuItems() {
        let settings = {
            guid: window.Asc.plugin.guid,
            items: [{
                id: "onMakeGroup",
                text: "设置分组",
            }]
        }

        let tagObj = ''
        if (activeQuesItem) {
            var controlTag = activeQuesItem ? activeQuesItem.Tag : "";
            if (controlTag != '') {
                try {
                    tagObj = JSON.parse(controlTag);
                    if (tagObj.group !== undefined && tagObj.group !== "") {
                        settings.items.push({
                            id: "onDismissGroup",
                            text: "解除分组",
                        })
                    }
                    if (tagObj.regionType == 'question') {
                        settings.items.push({
                            id: "onSettingDialog",
                            text: "设置题目信息",
                        })
                    }
                } catch (error) {
                    console.log(error)
                }
            }
        }
        return settings;
    }

    window.Asc.plugin.attachEvent('onContextMenuShow', function (options) {
        console.log(options);
        //     if (!options) return;

        if (options.type === 'Selection' || options.type === 'Target') {
            this.executeMethod('AddContextMenuItem', [getContextMenuItems()]);
        }
    });

    window.Asc.plugin.attachContextMenuClickEvent('onDismissGroup', function () {
        console.log("onDismissGroup");
        DismissGroup();
    });

    window.Asc.plugin.attachContextMenuClickEvent('onMakeGroup', function () {
        console.log("onMakeGroup");
        if (window.prevControl === undefined) {
            return;
        }
        MakeGroup(window.prevControl, window.currControl);
    });

    function onGetPos(rect) {
        if (rect === undefined) {
            return;
        }

        if (rect.length !== undefined && rect.length > 0) {
            console.log('onGetPos:', rect);    
            rect = rect[0];        
        }


        console.log('onGetPos:', rect);
        document.getElementById("p-value").innerHTML = rect.Page ? rect.Page + 1 : 1;
        document.getElementById("x-value").innerHTML = mmToPx(rect.X0);
        document.getElementById("y-value").innerHTML = mmToPx(rect.Y0);
        document.getElementById("w-value").innerHTML = mmToPx(rect.X1 - rect.X0);
        document.getElementById("h-value").innerHTML = mmToPx(rect.Y1 - rect.Y0);
    };

    let setCurrentContentControlLock = function (lock) {
        // 直接设置Lock以后，lock都不能操作了，锁定了不能通过插件操作？

        Asc.scope.lock = lock;
        window.Asc.plugin.callCommand(function () {
            // 获取当前控件id
            const sContentControlId = Api.asc_GetCurrentContentControl();
            if (sContentControlId) {
                const prop = { Lock: Asc.scope.lock };
                Api.asc_SetContentControlProperties(prop, sContentControlId, false);
            }
        }, false, false, undefined);
    };

    window.Asc.plugin.onCommandCallback = function (result) {
        //console.log("onCommandCallback", result);
    };

    let createContentControl = function (ranges) {
        Asc.scope.ranges = ranges;

        return biyueCallCommand(window, function () {
            var ranges = Asc.scope.ranges;

            let MakeRange = function (beg, end) {
                if (typeof beg === 'number')
                    return Api.GetDocument().GetRange().GetRange(beg, end);
                try {
                    return Api.asc_MakeRangeByPath(e.beg, e.end);
                }
                catch (error) {
                    console.error('MakeRange error:', error, e.beg, e.end);
                    return undefined;
                }
            }

            console.log('createContentControl count=', ranges.length);

            var results = [];
            // reverse order loop to keep the order
            for (var i = ranges.length - 1; i >= 0; i--) {
                // set selection
                var e = ranges[i];
                //console.log('createContentControl:', e);                    
                var range = MakeRange(e.beg, e.end);
                range.Select()
                
                // 特殊处理，如果range的最后一个段落是DocumentContent的最后一个段落，创建ContentControl过程中会额外增加一个段落
                // 需要在创建ContentControl之后，将DocumentContent的最后一个段落删除
                


                var oResult = Api.asc_AddContentControl(e.controlType || 1, { "Tag": e.info ? JSON.stringify(e.info) : '' });

                Api.asc_RemoveSelection();
                if (e.column !== undefined && e.column > 1) {
                    results.push({
                        "InternalId": oResult.InternalId,
                        "Tag": oResult.Tag,
                        "Column": e.column
                    })
                }
            }

            console.log('command createContentControl done');
            
            return results;
        }, false, false);
    }

    let insertDrawingObject = function () {
        console.log("insertDrawingObject")
        window.Asc.plugin.callCommand(function () {
            var oDocument = Api.GetDocument();
            // get current paragraph
            var pos = oDocument.Document.CurPos.ContentPos;
            var oElement = oDocument.GetElement(pos)
            while (oElement.GetClassType !== "paragraph") {
                if (oElement.GetClassType() === "blockLvlSdt") {
                    oElement = oElement.GetContent();
                } else if (oElement.GetClassType() === "documentContent") {
                    pos = oElement.Document.CurPos.ContentPos;
                    oElement = oElement.GetElement(pos);
                } else if (oElement.GetClassType() === "table") {
                    var colIndex = oElement.Table.CurCell.Index;
                    var rowIndex = oElement.Table.CurCell.Row.Index;
                    oElement = oElement.GetCell(rowIndex, colIndex).GetContent();
                } else {
                    break;
                }
            }
            var oParagraph = oElement;
            console.log("oParagraph", oParagraph);
            var oRGBColor = Api.CreateRGBColor(111, 111, 61);
            var oFill = Api.CreateSolidFill(oRGBColor);
            var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
            var oDrawing = Api.CreateShape("rect", 1908000, 1404000, oFill, oStroke);
            oDrawing.SetDistances(457200, 457200, 457200, 0);
            oDrawing.SetWrappingStyle("square");
            oDrawing.SetHorAlign("page", "center");
            oParagraph.AddDrawing(oDrawing);

        }, false, true, undefined);
    }

    let showMultiPagePos = function (window, onGetPos) {
        window.Asc.plugin.executeMethod("GetCurrentContentControl");
        window.Asc.plugin.onMethodReturn = function (returnValue) {
            if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {
                console.log("controlId", JSON.stringify(returnValue));

                if (returnValue) {
                    Asc.scope.controlId = returnValue;
                    window.Asc.plugin.callCommand(function () {
                        var isPageCoord = true;
                        var rects = Api.asc_GetContentControlBoundingRectExt(Asc.scope.controlId, isPageCoord);
                        return rects;
                    }, false, false, onGetPos);

                }
            }
        };
    }



    let SetContentProp = function (id, key, value) {
        window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function (obj) {
            window.Asc.plugin.currentContentControl = obj;
            var controlTag = obj ? obj.Tag : "";

            var tagObj = JSON.parse(controlTag);
            tagObj[key] = value;
            var newTag = JSON.stringify(tagObj);
            window.Asc.plugin.callCommand(function () {
                Api.asc_SetContentControlTag(newTag, id);
            }, false, false, undefined);

        });

    }


    let toggleControlStyle = function () {
        if (styleEnable) {
            styleEnable = false;
        } else {
            styleEnable = true;
        }
        console.log("styleEnable", styleEnable)

        // 方案1:通过调整ContentControl的属性来显示，现在看来只支持选中的control
        // 方案2：调整contentcontrol内的段落对应的样式，来做显示效果，多个段落如何做到效果一致，可以用A，B效果来加强

        // get all controls

        Asc.scope.styleEnable = styleEnable;
        window.Asc.plugin.callCommand(function () {
            const styleEnable = Asc.scope.styleEnable;
            Api.asc_SetGlobalContentControlShowHighlight(styleEnable, 204, 255, 255);
        }, false, true, undefined);

        // const applyBorder = function(para, border, order) {
        //     console.log("border", order, order % 2);
        //     var width = border ? 8 : 0;
        //     if (order % 2 === 1) {
        //         para.SetTopBorder("single", width, 0, 255, 111, 61);
        //         para.SetBottomBorder("single", width, 0, 255, 111, 61);
        //         para.SetLeftBorder("single", width, 0, 255, 111, 61);
        //         para.SetRightBorder("single", width, 0, 255, 111, 61);
        //     } else {
        //         para.SetTopBorder("single", width, 0, 0, 61, 111);
        //         para.SetBottomBorder("single", width, 0, 0, 61, 111);
        //         para.SetLeftBorder("single", width, 0, 0, 61, 111);
        //         para.SetRightBorder("single", width, 0, 0, 61, 111);
        //     }
        // };


        // var oDocument = Api.GetDocument();
        // var controls = oDocument.GetAllContentControls();

        // for (var i = 0; i < controls.length; i++) {
        //     var e = controls[i];
        //     var paras = e.GetAllParagraphs();
        //     for (var j=0; j < paras.length; j++)  {
        //         applyBorder(paras[j], styleEnable, i);
        //     }
        // }


    }


    function mmToPx(mm) {
        // 1 英寸 = 25.4 毫米
        // 1 英寸 = 96 像素（常见的屏幕分辨率）
        // 因此，1 毫米 = (96 / 25.4) 像素
        const pixelsPerMillimeter = 96 / 25.4;
        return Math.floor(mm * pixelsPerMillimeter);
    }

    function CalcTextPos(text_all, text_plain) {
        text_plain = text_plain.replace(/[\r]/g, '');
        var text_pos = new Array(text_all.length);
        var j = 0;
        for (var i = 0, n = text_plain.length; i < n; i++) {
            while (text_all[j] !== text_plain[i]) {
                text_pos[j] = i;
                j++;
            }
            text_pos[j] = i;
            j++;
        }

        return text_pos;
    }


    $(document).ready(function () {
        // 获取文档描述信息
        document.getElementById("getDocInfo").onclick = function () {
            GetDocInfo();
        }

        // 切题
        document.getElementById("splitQuestionBtn").onclick = function () {
            // get all text
            biyueCallCommand(window, function () {

                // Api.asc_EditSelectAll();
                // var text = Api.asc_GetSelectedText();
                // Api.asc_RemoveSelection();
                var oDocument = Api.GetDocument();
                var text_all = oDocument.GetRange().GetText({ Math: false, TableCellSeparator: "\u24D2", TableRowSeparator: "\u24E1" }) || "";
                var text_plain = oDocument.GetRange().GetText({ Math: false, Numbering: false, TableCellSeparator: "\u24D2", TableRowSeparator: "\u24E1" });

                return { text_all, text_plain };
            }, false, false)
            .then(function (result) {
                // split with token
                console.log('splitQuestion:', result)
                var text_pos = CalcTextPos(result.text_all, result.text_plain);
                var ranges = splitQuestion(result.text_all, text_pos);
                createContentControl(ranges);
            })
            .then(() => processTableColumn(undefined) );
        }

        var btnClear = document.getElementById("clearControl")
        if (btnClear) {
            btnClear.onclick = function () {
                window.Asc.plugin.executeMethod("GetAllContentControls");
                window.Asc.plugin.onMethodReturn = function (controls) {
                    if (window.Asc.plugin.info.methodName == "GetAllContentControls") {
                        Asc.scope.controls = controls;
                        biyueCallCommand(window, function () {
                            var controls = Asc.scope.controls;

                            for (var i = 0; i < controls.length; i++) {
                                // set selection
                                var e = controls[i];
                                Api.asc_RemoveContentControlWrapper(e.InternalId);
                            }
                        }, false, false)
                        .then(() => console.log("删除所有控件完成"));
                    }
                };

            };

            document.getElementById("checkAnswerRegionBtn").onclick = function () {
                checkAnswerRegion();
            }

            document.getElementById("toggleStyleBtn").onclick = function () {
                toggleControlStyle();
            }
            // Todo 考虑其他实现方法
            // 锁定控件操作
            document.getElementById("unlockBtn").onclick = function () {
                /* Asc.c_oAscSdtLockType.Unlocked */
                setCurrentContentControlLock(3);
            };

            var btnLock = document.getElementById("lockBtn")
            if (btnLock) {
                btnLock.onclick = function () {
                    // 1 Asc.c_oAscSdtLockType.SdtContentLocked
                    setCurrentContentControlLock(1);
                };
            }
            var btnShowPos = document.getElementById("showPosBtn")
            if (btnShowPos) {
                btnShowPos.onclick = function () {
                    console.log("showPosition on button clicked");
                    showPosition(window, onGetPos);
                }
            }
            var btnGetSelection = document.getElementById("getSelectionBtn")
            if (btnGetSelection) {
                btnGetSelection.onclick = function () {
                    console.log("getSelection on button clicked");
                    getSelection();
                }
            }
            var btnShowContent = document.getElementById("showContentBtn")
            if (btnShowContent) {
                btnShowContent.onclick = function () {
                    console.log("showContent on button clicked");
                    showContent();
                }
            }
            var btnShowScoreContent = document.getElementById("showScoreContentBtn")
            if (btnShowScoreContent) {
                btnShowScoreContent.onclick = function () {
                    console.log("showScoreContent on button clicked");
                    var fun = function () {
                        var controls = Api.GetDocument().GetAllContentControls();
                        for (var i = 0; i < controls.length; i++) {
                            var control = controls[i];
                            var obj = ''
                            if (control && control.GetTag()) {
                                obj = control.GetTag() || ''
                                if (obj) {
                                    try {
                                        obj = JSON.parse(obj)
                                    } catch (e) {
                                        console.error('JSON解析失败', e)
                                    }
                                }
                            }
                            if (obj && obj.ques_no <= 4) {
                                var oDocument = Api.GetDocument();
                                var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
                                var oTable = Api.CreateTable(5, 1);
                                // oTable.SetWidth("percent", 100);
                                oTable.SetStyle(oTableStyle);
                                oTable.SetWrappingStyle(false);
                                oTable.GetCell(0, 0).GetContent().GetElement(0).AddText("1")
                                oTable.GetCell(0, 0).SetWidth('twips', 283)
                                oTable.GetCell(0, 1).GetContent().GetElement(0).AddText("2")
                                oTable.GetCell(0, 1).SetWidth('twips', 283)
                                oTable.GetCell(0, 2).GetContent().GetElement(0).AddText("3")
                                oTable.GetCell(0, 2).SetWidth('twips', 283)
                                oTable.GetCell(0, 3).GetContent().GetElement(0).AddText("4")
                                oTable.GetCell(0, 3).SetWidth('twips', 283)
                                oTable.GetCell(0, 4).GetContent().GetElement(0).AddText("5")
                                oTable.GetCell(0, 4).SetWidth('twips', 283)
                                // oDocument.Push(oTable) // 添加到文档的底部

                                // 表格-高级设置 相关参数
                                var Props = {
                                    CellSelect: true,
                                    Locked: false,
                                    PositionV: {
                                        Align: 1,
                                        RelativeFrom: 2,
                                        UseAlign: true,
                                        Value: 0
                                    },
                                    PositionH: {
                                        Align: 4,
                                        RelativeFrom: 0,
                                        UseAlign: true,
                                        Value: 0
                                    },
                                    TableDefaultMargins: {
                                        Bottom: 0,
                                        Left: 0,
                                        Right: 0,
                                        Top: 0
                                    }
                                }
                                // Api.tblApply(Props)
                                oTable.Table.Set_Props(Props);
                                control.AddElement(oTable, 0) // 添加到控件的0的开始位置
                            }
                        }
                    }
                    window.Asc.plugin.callCommand(fun, false, true, undefined)
                }
            }
            var btnInsertDrawingObject = document.getElementById("insertDrawingObjectBtn")
            if (btnInsertDrawingObject) {
                btnInsertDrawingObject.onclick = function () {
                    insertDrawingObject();
                }
            }
            var btnShowMultiPagePos = document.getElementById("showMultiPagePosBtn")
            if (btnShowMultiPagePos) {
                btnShowMultiPagePos.onclick = function () {
                    showMultiPagePos(window, onGetPos);
                }
            }
            var btnToTableColumn = document.getElementById("toTableColumnBtn")
            if (btnToTableColumn) {
                btnToTableColumn.onclick = function () {
                    toTableColumn(window);
                }
            }
            var btnGetQuestionPositions = document.getElementById("getQuestionPositionsBtn")
            if (btnGetQuestionPositions) {
                btnGetQuestionPositions.onclick = function () {
                    getQuestionPositions(window);
                }
            }
            var btnSplitJsonPath = document.getElementById("jsonPathSplitQuestionBtn")
            if (btnSplitJsonPath) {
                btnSplitJsonPath.onclick = function () {
                    // get all text 执行顺序为先进后出，下面的先执行                    
                    biyueCallCommand(window, function () {

                        // Api.asc_EditSelectAll();
                        // var text = Api.asc_GetSelectedText();
                        // Api.asc_RemoveSelection();
                        var oDocument = Api.GetDocument();
                        var text_all = oDocument.GetRange().GetText({ Math: false, TableCellSeparator: "\u24D2", TableRowSeparator: "\u24E1" }) || "";
                        var text_json = oDocument.ToJSON(false, false, false, false, true, true);

                        return { text_all, text_json };
                    }, false, false)
                    .then((result) => {
                            var ranges = newSplit(result.text_json);
                            console.log('splitQuestion:', ranges)
                            return createContentControl(ranges);})
                    //.then(() => { console.log("1.处理第三级子题"); return checkSubQuestion(); })
                    .then(() => { console.log("2.处理需要分列的题目"); return processTableColumn(undefined); })
                    .then(() => { console.log("3.处理答题区域"); return checkAnswerRegion(); }) 
                    .then(() => { console.log("4.刷新控件");  return updateCustomControls(); });
                }
            }

            document.getElementById("normalizeDoc").onclick = function () {
                window.Asc.plugin.callCommand(function () {

                    // Api.asc_EditSelectAll();
                    // var text = Api.asc_GetSelectedText();
                    // Api.asc_RemoveSelection();
                    var oDocument = Api.GetDocument();
                    var text_all = oDocument.GetRange().GetText({ Math: false, TableCellSeparator: "\u24D2", TableRowSeparator: "\u24E1" }) || "";
                    var text_json = oDocument.GetRange().ToJSON(true);

                    return { text_all, text_json };
                }, false, false, function (result) {
                    var ranges = normalizeDoc(result.text_json);
                    console.log('normal:', ranges)
                    execModify(ranges);
                });
            }

            document.getElementById("phoneticGuide").onclick = function () {
                ReplaceRubyField(window, undefined, function () {
                    console.log("ReplaceRubyField done");
                });
                
            }


            document.getElementById("selectionToHtml").onclick = function () {
                rangeToHtml(window, undefined, function (html) {
                    console.log(html);
                });
            }

            document.getElementById("selectionToXml").onclick = function () {
                toXml(window, undefined, function (xml) {
                    console.log(xml);
                });
            }

            document.getElementById("insertAsHtml").onclick = function () {

                var html = `<p
            style="margin-top:0pt;margin-bottom:10pt;border:none;border-left:none;border-top:none;border-right:none;border-bottom:none;mso-border-between:none">
            <span style="font-family:'Arial';font-size:11pt;color:#000000;mso-style-textfill-fill-color:#000000">Hello word</span>
         </p>
         <table>
         <tr>
            <td>1</td>
            <td>2</td>
            </tr>
            <tr>
            <td>3</td>
            <td>4</td>
            </tr>
            </table>
        `;

                insertHtml(window, undefined, html, function (res) {
                    console.log(res);
                });


            }

            document.getElementById("versionTag").innerHTML = getVersion();
        }
    });

    // 在editor面板的插件按钮被点击
    window.Asc.plugin.button = function (id, windowID) {
        console.log("on plugin button id=${id} ${windowID}", id, windowID);
        if (windowID) {
            if (id === -1) {
                window.Asc.plugin.executeMethod('CloseWindow', [windowID])
            }
            return
        }
        if (id == -1) {
            console.log('StoreCustomData', window.BiyueCustomData)
            StoreCustomData(function (uuid) {
                window.BiyueCustomId = id;
                console.log("store custom data done");
                window.Asc.plugin.executeCommand("close", '');
            });
            return;
        }
    };

    function showPosition(window, onGetPos) {
        window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function (returnValue) {
            console.log("control", returnValue);

            if (returnValue) {
                Asc.scope.controlId = returnValue.InternalId;
                biyueCallCommand(window, function () {
                    // get logic document
                    // get control
                    // get bound rect
                    // transform
                    // return
                    const isPageCoord = true;
                    var rect = Api.asc_GetContentControlBoundingRect(Asc.scope.controlId, isPageCoord);
                    return rect;
                }, false, false)
                .then((rect) => onGetPos(rect));
            }
        });
    };

    function onGetTag(tag) {
        console.log("onGetTag", tag);
    }

    // 获取当前控件的tag
    function getCurrentTag(window) {
        window.Asc.plugin.callCommand(function () {
            var prop = Api.asc_GetContentControlProperties();
            return prop.Tag;
        }, false, false, function (tag) {
            console.log("tag=>", tag);
        });
    };

    // 设置当前控件的tag
    function setCurrentTag(window, tag) {
        Asc.scope.tag = tag;
        window.Asc.plugin.executeMethod("GetCurrentContentControl");
        window.Asc.plugin.onMethodReturn = function (returnValue) {
            if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {
                if (returnValue) {
                    Asc.scope.controlId = returnValue;
                    window.Asc.plugin.callCommand(function () {
                        var controls = Api.GetDocument().GetAllContentControls();
                        for (var i = 0; i < controls.length; i++) {
                            var control = controls[i];
                            console.log("control", control, Asc.scope.controlId)
                            if (control.Sdt.GetId() === Asc.scope.controlId) {
                                control.SetTag(Asc.scope.tag);
                            }
                        }
                    }, false, true, undefined);

                }
            }
        };
    }

    function showAllContent() {
        Asc.scope.styleEnable = !Asc.scope.styleEnable;
        window.Asc.plugin.callCommand(function () {
            const styleEnable = Asc.scope.styleEnable;
            // 设置控件的高亮颜色
            Api.asc_SetGlobalContentControlShowHighlight(styleEnable, 255, 204, 204);
        }, false, false, undefined);
    }

    function getSelection() {
        // [类型]: 1为block的 2的为inline
        // window.Asc.plugin.executeMethod ("AddContentControl", [1]);

        //   window.Asc.plugin.callCommand(function () {
        //   var oDocument = Api.GetDocument();
        //   var pos = oDocument.Document.CurPos.ContentPos;
        //   var oElement = oDocument.GetElement(pos)
        //   while (oElement.GetClassType !== "paragraph") {
        //       if (oElement.GetClassType() === "blockLvlSdt") {
        //           oElement = oElement.GetContent();
        //       } else if (oElement.GetClassType() === "documentContent") {
        //           pos =  oElement.Document.CurPos.ContentPos;
        //           oElement = oElement.GetElement(pos);
        //       } else if (oElement.GetClassType() === "table") {
        //           var colIndex = oElement.Table.CurCell.Index;
        //           var rowIndex = oElement.Table.CurCell.Row.Index;
        //           oElement = oElement.GetCell(rowIndex, colIndex).GetContent();
        //       } else {
        //           break;
        //       }
        //   }
        //   var oParagraph = oElement;
        //   console.log('GetParentContentControl:', oParagraph.GetParentContentControl() )
        // }, false, false, undefined);

        window.Asc.plugin.callCommand(function () {
            var oDocument = Api.GetDocument();
            var aSections = oDocument.GetSections();
            var sClassType = aSections[0].GetClassType();
            var oParagraph = oDocument.GetElement(0)
            // var oRange = oDocument.GetRange(8, 11);
            var oRange = oDocument.GetRangeBySelect()
            // Api.asc_AddContentControl(1);
            // Api.asc_RemoveSelection();
            // oRange.SetBold(true);
            var hasContentControl = oRange.Paragraphs[0].GetParentContentControl()
            var type = 1
            if (hasContentControl) {
                // sdt.Pr.Tag 存储题目相关信息
                type = 2
            }
            console.log('oRange::', oRange.Paragraphs[0].GetParentContentControl())
            console.log("aSections::", oDocument.GetRangeBySelect(), oDocument.GetRange())
            let allText = oDocument.GetRange().Text
            let selectText = oRange.Text
            console.log('-------:', allText.indexOf(selectText))
            return { type }
        }, false, false, function (obj) {
            if (obj && obj.type === 2) {
                window.Asc.plugin.executeMethod("AddContentControl", [2]);
            } else {
                window.Asc.plugin.executeMethod("AddContentControl", [1]);
                let obj = { 'regionType': 'question', 'mode': 2, padding: [0, 0, 0.5, 0] }
                let Tag = JSON.stringify(obj)
                setTag(window, Tag)
            }
        });
    }

    // 获取当前控件的tag
    function getTag(window) {
        window.Asc.plugin.callCommand(function () {
            var prop = Api.asc_GetContentControlProperties();
            return prop.Tag;
        }, false, false, function (tag) {
            console.log("tag=>", tag);
        });
    };

    // 设置当前控件的tag
    function setTag(window, tag) {
        Asc.scope.tag = tag;
        window.Asc.plugin.executeMethod("GetCurrentContentControl");
        window.Asc.plugin.onMethodReturn = function (returnValue) {
            if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {
                if (returnValue) {
                    Asc.scope.controlId = returnValue;
                    window.Asc.plugin.callCommand(function () {
                        var controls = Api.GetDocument().GetAllContentControls();
                        for (var i = 0; i < controls.length; i++) {
                            var control = controls[i];
                            console.log("control", control, Asc.scope.controlId)
                            if (control.Sdt.GetId() === Asc.scope.controlId) {
                                control.SetTag(Asc.scope.tag);
                            }
                        }
                    }, false, false, undefined);

                }
            }
        };
    }


    window.test = {};
    window.test["setCurrentTag"] = setCurrentTag;
    window.test["getCurrentTag"] = getCurrentTag;

    window.prevControl = undefined;

    window.Asc.plugin.event_onBlurContentControl = function (control) {
        //console.log("onBlurControl", control);
        window.prevControl = control;
    }

    let DismissGroup = function () {
        window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function (obj) {
            if (obj === undefined || obj === null || obj.Tag === undefined || !obj.Tag.includes("group")) {
                return;
            }

            try {
                var tagObj = JSON.parse(obj.Tag);
                if (tagObj.group !== undefined && tagObj.group !== "") {
                    window.Asc.plugin.executeMethod("GetAllContentControls", [], function (controls) {
                        var ots = [];
                        for (var i = 0; i < controls.length; i++) {
                            var e = controls[i];
                            if (e.Tag === undefined || e.Tag === "") {
                                continue;
                            }

                            try {
                                var tag = JSON.parse(e.Tag);
                                if (tag.group === tagObj.group) {
                                    tag.group = undefined;
                                    ots.push({ Id: e.InternalId, tag: JSON.stringify(tag) });
                                }
                            }
                            catch (e) {
                                // console.log(e);
                            }

                        }
                        setBatchTag(window, ots);
                    });
                }
            }
            catch (e) {
                console.log(e);
            }
        });
    };

    function setBatchTag(window, objTagPairList) {
        Asc.scope.objTagPairList = objTagPairList;

        window.Asc.plugin.callCommand(function () {
            var objTagPairList = Asc.scope.objTagPairList;
            let findTag = function (sdt) {
                for (var i = 0; i < objTagPairList.length; i++) {
                    var objTagPair = objTagPairList[i];
                    if (objTagPair.InternalId !== undefined && objTagPair.InternalId === sdt.Id) {
                        return objTagPair.tag;
                    }
                    if (objTagPair.Id !== undefined && objTagPair.Id === sdt.Id) {
                        return objTagPair.tag;
                    }
                }
                return undefined;
            }

            var controls = Api.GetDocument().GetAllContentControls();
            for (var i = 0; i < controls.length; i++) {
                var control = controls[i];
                console.log("control", control, Asc.scope.controlId);
                var tag = findTag(control.Sdt);
                if (tag !== undefined) {
                    control.SetTag(tag);
                }
            }
        }, false, false, undefined);
    };


    let MakeGroup = function (prevControl, curControl) {
        console.log("MakeGroup", prevControl, curControl);
        if (prevControl === undefined || curControl === undefined) {
            return;
        }

        var prevTagObj = JSON.parse(prevControl.Tag || "{}") || {};
        var curTagObj = JSON.parse(curControl.Tag || "{}") || {};

        if (prevTagObj.group === undefined) {
            prevTagObj.group = prevControl.InternalId;
        }

        if (prevTagObj.group === curControl.group) {
            return;
        }

        curTagObj.group = prevTagObj.group;

        var prevTag = JSON.stringify(prevTagObj);
        var curTag = JSON.stringify(curTagObj);

        // set tag
        setBatchTag(window, [{ "InternalId": prevControl.InternalId, "tag": prevTag }, { "InternalId": curControl.InternalId, "tag": curTag }]);
    };
    let SettingDialog = function () {
        // 题目设置信息窗口
        let location = window.location;
        let start = location.pathname.lastIndexOf('/') + 1;
        let file = location.pathname.substring(start);

        let variation = {
            url: location.href.replace(file, 'settings.html'),
            description: window.Asc.plugin.tr('题目设置'),
            isVisual: true,
            isModal: true,
            isViewer: true,
            buttons: [
                {
                    'text': window.Asc.plugin.tr('确定'),
                    'primary': true
                },
                {
                    'text': window.Asc.plugin.tr('取消'),
                    'primary': false
                }],
            EditorsSupport: ["word"],
            size: [592, 200]
        };

        if (!settingsWindow) {
            settingsWindow = new window.Asc.PluginWindow();
            settingsWindow.attachEvent("onBiyueMessage", function (message) {
                settingsWindow.command('onParams', activeQuesItem)
            });
            settingsWindow.attachEvent("getSettingsMessage", function (params) {
                console.log('getSettingsMessage:', params)
                let Tag = ''
                if (params && params.activeQuesItem && params.form) {
                    let form = params.form || {}
                    let itemObj = JSON.parse(params.activeQuesItem.Tag) || {}
                    itemObj.score = form.score || 0
                    itemObj.mode = form.mode || 0
                    Tag = JSON.stringify(itemObj)
                }

                settingsWindow.close();
                settingsWindow = null
                setTag(window, Tag)
            })
        }
        settingsWindow.show(variation);
    }


    window.Asc.plugin.event_onFocusContentControl = function (control) {
        biyueCallCommand(window,
            function () {
                return AscCommon.global_keyboardEvent.CtrlKey;                
            },
            false,
            true)
        .then((ctrlKey) => {
                if (true === ctrlKey && prevControl !== undefined && control !== undefined && control.InternalId != prevControl.InternalId) {
                    window.currControl = control;
                } else {
                    window.prevControl = undefined;
                    window.currControl = undefined;
                }
            });
    }


    window.Asc.plugin.event_onClick = function (isSelectionUse) {
        console.log("event click");
        showPosition(window,
            function (data) {
                onGetPos(data);
                window.Asc.plugin.callCommand(function () {
                    Api.GetDocument().Document.Recalculate(true);
                },
                    false,
                    true);

            }
        );

    };

    // 将一行多题目的控件转为表格
    // oPr InternalId ID
    // oPr Tag 标签    
    function processTableColumn(oPr) {        
        // if oPr is not array 
        if (oPr && oPr.length === undefined) {
            if (oPr === undefined || oPr === null || oPr.Tag === undefined) {
                return;
            }
            oPr = [oPr];
        }

        Asc.scope.controlPrs = oPr;
        return biyueCallCommand(window, function () {
            function CalcTextPos(text_all, text_plain) {
                text_plain = text_plain.replace(/[\r]/g, '');
                var text_pos = new Array(text_all.length);
                var j = 0;
                for (var i = 0, n = text_plain.length; i < n; i++) {
                    while (text_all[j] !== text_plain[i]) {
                        text_pos[j] = i;
                        j++;
                    }
                    text_pos[j] = i;
                    j++;
                }

                return text_pos;
            }

            var oPrs = Asc.scope.controlPrs;
            var oDocument = Api.GetDocument();
            var oControls = oDocument.GetAllContentControls();
            if (oPrs !== undefined) {
                oControls = oControls.filter(function (control) {
                    for (var i = 0; i < oPrs.length; i++) {
                        if (control.Sdt.Id === oPrs[i].InternalId) {
                            return true;
                        }
                    }
                    return false;
                });
            }

            if (oControls.length === 0) {
                return;
            }

            oControls.forEach(function (oControl, index) {



                // 不能有子节点
                if (oControl.GetAllContentControls().length > 0) {
                    return;
                }

                var oPrInternalId = oControl.Sdt.GetId();
                var oPrTag = oControl.GetTag();
                var text = oControl.GetRange().GetText({ Math: false })
                var text_plain = oControl.GetRange().GetText({ Math: false, Numbering: false });
                var text_pos = CalcTextPos(text, text_plain);

                text = text.replace(/[\uFF10-\uFF19]|．|（|）/g, function (c) {
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
                var isInlinePatt = [
                    /(?<=^)(\d+\..*)([ ]+\d+\..*)+/g,
                    /(?<=^)(\(\d+\).*)([ ]+\(\d+\).*)+/g
                ];

                var inlineQuesPatt = /\(?\d+[\).].*?(?=(\d+\.)|(\(\d+\))|\r|$)/gs;

                var inlineArr;
                for (var i = 0; i < isInlinePatt.length; i++) {
                    var inlineArr = text.match(isInlinePatt[i]);
                    if (inlineArr !== null) {
                        break;
                    }
                }
                if (inlineArr === null || inlineArr === undefined) {
                    return;
                }


                var lines = text.match(/\n/g);
                if (lines == null || lines.length <= 1) {
                    return;
                }

                var quesTextArr = inlineArr[0].match(inlineQuesPatt);
                if (quesTextArr === null || quesTextArr.length <= 1) {
                    return;
                }

                // 在第几行
                var quesLineNoArr = quesTextArr.map(function (item) {
                    return text.substr(0, text.indexOf(item)).split('\n').length - 1;
                });
                console.log("oControl", oControl);

                // 共几行
                var quesLineCountArr = []
                var totalLines = text.split('\n').length - 1;
                for (var i = 0; i < quesLineNoArr.length; i++) {
                    var found = false
                    for (var j = i + 1; j < quesLineNoArr.length; j++) {
                        if (quesLineNoArr[j] > quesLineNoArr[i]) {
                            quesLineCountArr.push(quesLineNoArr[j] - quesLineNoArr[i]);
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        quesLineCountArr.push(totalLines - quesLineNoArr[i]);
                    }
                }

                var rows = 0;
                var cols = 0;
                var max_cols = 1;
                var quesTablePos = []
                quesLineNoArr.forEach(function (item, index) {
                    if (index === 0) {
                        rows = 1;
                        cols = 1;
                        quesTablePos.push({ row: 0, col: 0, spaceLines: quesLineCountArr[index] });
                    } else {
                        if (item === quesLineNoArr[index - 1]) {
                            cols++;
                            if (cols > max_cols) {
                                max_cols = cols;
                            }
                        } else {
                            rows++;
                            cols = 1;
                        }
                        quesTablePos.push({ row: rows - 1, col: cols - 1, spaceLines: quesLineCountArr[index] });
                    }
                })


                // create table 1xN 
                var oTable = Api.CreateTable(max_cols, rows);
                oTable.SetWidth("percent", 100);
                var oTableStyle = oDocument.CreateStyle("CustomTableStyle", "table");
                var oTableCellPr = oTableStyle.GetTableCellPr();
                oTableCellPr.SetWidth("percent", 100 / max_cols);
                oTable.SetStyle(oTableStyle);

                // todo set table no border                    

                // split text                
                var lastItem = quesTextArr[quesTextArr.length - 1];
                var maxPos = text_pos[text.length - 1];

                quesTextArr = quesTextArr.map(function (item) {
                    return item.replace(/[ ]+/g, ' ');
                });

                quesTextArr.forEach(function (item, index) {
                    var tablePos = quesTablePos[index];
                    var oCell = oTable.GetCell(tablePos.row, tablePos.col);
                    var oContent = oControl.GetContent().GetContent(true)
                    // add text to table cell                    

                    oContent.forEach(function (item, paraIndex) {
                        oCell.GetContent().Push(item);
                    });
                    oCell.GetContent().RemoveElement(0);

                    // var beg = text.indexOf(item);
                    // var end = beg + item.length;
                    // var begPos = text_pos[beg];
                    // var endPos = text_pos[end];

                    // if (begPos > 1) {
                    //     oCell.GetContent().GetRange(0, begPos - 1).Delete();
                    // }

                    // if (endPos < maxPos) {
                    //     oCell.GetContent().GetRange(endPos, maxPos).Delete();
                    // }
                });

                oControl.GetContent().RemoveAllElements();
                oControl.GetContent().Push(oTable);
                oControl.GetContent().RemoveElement(0);

                // loop table                
                for (var i = 0; i < quesTablePos.length; i++) {
                    var row = quesTablePos[i].row;
                    var col = quesTablePos[i].col;
                    var oCell = oTable.GetCell(row, col);

                    var item = quesTextArr[i];
                    var beg = text.indexOf(item);
                    var end = beg + item.length;
                    var begPos = text_pos[beg];
                    var endPos = text_pos[end];

                    if (endPos < maxPos) {
                        oCell.GetContent().GetRange(endPos, maxPos).Delete();
                    }

                    if (begPos > 1) {
                        oCell.GetContent().GetRange(0, begPos - 1).Delete();
                    }

                    var range = oCell.GetContent().GetRange();
                    range.Select()
                    var oResult = Api.asc_AddContentControl(1, { "Tag": oPrTag });
                    Api.asc_RemoveSelection();
                }

                // remove root content control
                Api.asc_RemoveContentControlWrapper(oPrInternalId);
            });
        }, false, true);
    }

    function toTableColumn(window) {
        window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], processTableColumn, false, false, undefined);
    }

    function execModify(ranges) {

    }

    function GetDocInfo() {
        window.Asc.plugin.callCommand(function () {
            return Api.DocInfo;
        }
            , false, false, function (docInfo) {
                console.log("docInfo", docInfo);
            });
    }

    window.insertHtml = insertHtml;

})(window, undefined);

