(function (window, undefined) {
    var styleEnable = false;

    let splitQuestion = function (text) {
        var index = 1;
        const dlimiter = function(text) {
            var dlims = ['．', '[\.]'];
            for (var i = 0; i < dlims.length; i++) {
                if (-1 !== text.search(dlims[i])) {
                    return dlims[i];
                }
            }
            return dlims[0];
        }(text);

        let token = function (index) { return `${index}${dlimiter}`; }
        var nextstart = 0;
        var next = text.substring(nextstart).search(token(index));
        index = index + 1;
        var ranges = new Array();
        while (next !== -1) {
            ranges.push({ beg: nextstart, end: nextstart + next })
            nextstart = nextstart + next;
            next = text.substring(nextstart).search(token(index));
            index = index + 1;
        }

        if (nextstart < text.length) {
            ranges.push({ beg: nextstart, end: text.length });
        }

        ranges.shift();
        return ranges;
    }
    // 插件初始化    
    window.Asc.plugin.init = function () {
        console.log("biyue plugin inited.");

        // create style
        
        this.callCommand(
            function() {
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
                if (customData === undefined)
                    return;
                window.BiyueCustomId = customData[0].ItemId;
                window.BiyueCustomData = customData;
            }
        );
    };

    function StoreCustomData(callback) {
        if (window.BiyueCustomData === undefined) {
            callback();
            return;
        }

        Asc.scope.BiyueCustomId = window.BiyueCustomId;
        Asc.scope.BiyueCustomData = window.BiyueCustomData;
        window.Asc.plugin.callCommand(            
            function() {
                var id = Asc.scope.BiyueCustomId
                var data = Asc.scope.BiyueCustomData;
                Api.asc_SetBiyueCustomDataExt(id, data);                
            },
            false, 
            false, 
            callback
        );
    }

    function getContextMenuItems() {
        let settings = {
            guid: window.Asc.plugin.guid,
            items : [
                {
                    id: "onDismissGroup",
                    text: "解除分组",
                }
            ]
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

    window.Asc.plugin.attachContextMenuClickEvent('onDismissGroup', function() {
        console.log("onDismissGroup");
        DismissGroup();
    });



    function onGetPos(rect) {
        if (rect === undefined) {
            return;
        }
        console.log(rect);
        document.getElementById("p-value").innerHTML = rect.Page;
        document.getElementById("x-value").innerHTML = rect.X0;
        document.getElementById("y-value").innerHTML = rect.Y0;
        document.getElementById("w-value").innerHTML = rect.X1 - rect.X0;
        document.getElementById("h-value").innerHTML = rect.Y1 - rect.Y0;
    };

    setCurrentContentControlLock = function (lock) {
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

    createContentControl = function (ranges) {
        Asc.scope.ranges = ranges;
        window.Asc.plugin.callCommand(function () {
            var ranges = Asc.scope.ranges;

            for (var i = 0; i < ranges.length; i++) {
                // set selection
                var e = ranges[i];
                console.log(e);
                Api.GetDocument().GetRange(e.beg, e.end - 1).Select()
                Api.asc_AddContentControl(1);
                Api.asc_RemoveSelection();
            }
        }, false, false, undefined);
    }

    insertDrawingObject = function () {
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
                    pos =  oElement.Document.CurPos.ContentPos;
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

    showMultiPagePos = function (window, onGetPos) {
        window.Asc.plugin.executeMethod("GetCurrentContentControl");
        window.Asc.plugin.onMethodReturn = function (returnValue) {
            if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {
                console.log("controlId", JSON.stringify(returnValue));

                if (returnValue) {
                    Asc.scope.controlId = returnValue;
                    window.Asc.plugin.callCommand(function () {
                        var rects = Api.asc_GetContentControlBoundingRectExt(Asc.scope.controlId, isPageCoord);
                        return rect;
                    }, false, false, onGetPos);

                }
            }
        };
    }



    SetContentProp = function(id, key, value) {
        window.Asc.plugin.executeMethod ("GetCurrentContentControlPr", [], function (obj) {
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


    toggleControlStyle = function() {
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

    $(document).ready(function () {
        // 切题
        document.getElementById("splitQuestionBtn").onclick = function () {
            // get all text
            window.Asc.plugin.callCommand(function () {
                Api.asc_EditSelectAll();
                var text = Api.asc_GetSelectedText();

                Api.asc_RemoveSelection();
                return text;
            }, false, false, function (text) {
                // split with token
                var ranges = splitQuestion(text);

                createContentControl(ranges);
            });


        }

        document.getElementById("clearControl").onclick = function() {
            window.Asc.plugin.executeMethod("GetAllContentControls");
            window.Asc.plugin.onMethodReturn = function (controls) {
                if (window.Asc.plugin.info.methodName == "GetAllContentControls") {                    
                    Asc.scope.controls = controls;
                    window.Asc.plugin.callCommand(function () {
                        var controls = Asc.scope.controls;
            
                        for (var i = 0; i < controls.length; i++) {
                            // set selection
                            var e = controls[i];
                            Api.asc_RemoveContentControlWrapper(e.InternalId);                            
                        }
                    }, false, false, undefined);
                }
            };

        };

        document.getElementById("toggleStyleBtn").onclick = function() {
            toggleControlStyle();
        }

        // Todo 考虑其他实现方法
        // 锁定控件操作
        document.getElementById("unlockBtn").onclick = function () {
            /* Asc.c_oAscSdtLockType.Unlocked */
            setCurrentContentControlLock(3);
        };

        document.getElementById("lockBtn").onclick = function () {
            // 1 Asc.c_oAscSdtLockType.SdtContentLocked 
            setCurrentContentControlLock(1);
        };

        document.getElementById("showPosition").onclick = function () {
            console.log("showPosition on button clicked");
            showPosition(window, onGetPos);
        }

        document.getElementById("insertDrawingObject").onclick = function () {
            insertDrawingObject();
        }

        document.getElementById("showMultiPagePos").onclick = function () {
            showMultiPagePos(window, onGetPos);
        }
    });

    // 在editor面板的插件按钮被点击
    window.Asc.plugin.button = function (id, windowID) {
        console.log("on plugin button id=${id} ${windowID}");
        if (id == -1) {
            this.executeCommand("close", '');
            StoreCustomData(function() {
                console.log("store custom data done");                
            });            
            return;
        }
    };

    window.Asc.plugin.onFocusContentControl = function (control) {
        console.log("onFocusControl");
        console.log(control);
    };

    window.Asc.plugin.attachEvent("onSelectionChanged", function (data) {
        console.log("on SelectionChange");
    });


    function showPosition(window, onGetPos) {
        window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function (returnValue) {            
            console.log("control", returnValue);

            if (returnValue) {
                Asc.scope.controlId = returnValue.InternalId;
                window.Asc.plugin.callCommand(function () {
                    // get logic document
                    // get control 
                    // get bound rect
                    // transform
                    // return
                    const isPageCoord = true;
                    var rect = Api.asc_GetContentControlBoundingRect(Asc.scope.controlId, isPageCoord);
                    return rect;
                }, false, false, onGetPos);
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
        }, false, false, function(tag) {
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

    DismissGroup = function() {
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
                                    ots.push({Id: e.InternalId, tag: JSON.stringify(tag)});
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
            let findTag = function(sdt) {                 
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


    MakeGroup = function(prevControl, curControl) {
        console.log("MakeGroup", prevControl, curControl);
        if (prevControl === undefined || curControl === undefined) {
            return;
        }

        var prevTagObj =JSON.parse(prevControl.Tag || "{}") || {};
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
        setBatchTag(window, [{"InternalId": prevControl.InternalId, "tag": prevTag}, {"InternalId": curControl.InternalId, "tag": curTag}]);            
    };


    window.Asc.plugin.event_onFocusContentControl = function (control) {
        window.Asc.plugin.callCommand(
            function() {
                return AscCommon.global_keyboardEvent.CtrlKey;
            },
            false, 
            true, 
            function(ctrlKey) {
                if (true ===  ctrlKey && prevControl !== undefined && control !== undefined && control.InternalId != prevControl.InternalId) {
                    MakeGroup(prevControl, control);                    
                } else {
                    
                }
                window.prevControl = undefined;
            }
        );
    }



    window.Asc.plugin.event_onClick = function (isSelectionUse) {
        console.log("event click");
        showPosition(window, 
            function(data) {
                onGetPos(data);
                window.Asc.plugin.callCommand(function() {            
                    Api.GetDocument().Document.Recalculate(true);
                },
                false, 
                true);
                
            }
        );        
        

        
    };
})(window, undefined);

