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

        // this.callCommand(function() {
        //     var oDocument = Api.GetDocument();
        //     var oParagraph = Api.CreateParagraph();
        //     oParagraph.AddText("Biyue Hello world!");
        //     oDocument.InsertContent([oParagraph]);
        // }, true);        
        // if (!window.isInit) {
        //     window.isInit = true;
        //     window.Asc.plugin.currentText = "";
        //     window.Asc.plugin.createInputHelper();
        //     window.Asc.plugin.getInputHelper().createWindow();
        // }

    };

    function onGetPos(rect) {
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

            const applyBorder = function(para, border, order) {      
                console.log("border", order, order % 2);
                var width = border ? 8 : 0;
                if (order % 2 === 1) {
                    para.SetTopBorder("single", width, 0, 255, 111, 61);
                    para.SetBottomBorder("single", width, 0, 255, 111, 61);
                    para.SetLeftBorder("single", width, 0, 255, 111, 61);
                    para.SetRightBorder("single", width, 0, 255, 111, 61);
                } else {
                    para.SetTopBorder("single", width, 0, 0, 61, 111);
                    para.SetBottomBorder("single", width, 0, 0, 61, 111);
                    para.SetLeftBorder("single", width, 0, 0, 61, 111);
                    para.SetRightBorder("single", width, 0, 0, 61, 111);
                }
            };


            var oDocument = Api.GetDocument();
            var controls = oDocument.GetAllContentControls();                    
            
            for (var i = 0; i < controls.length; i++) {
                var e = controls[i];                        
                var paras = e.GetAllParagraphs();
                for (var j=0; j < paras.length; j++)  {                            
                    applyBorder(paras[j], styleEnable, i);
                }
            }
            
        }, false, true, undefined);
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
    });

    // 在editor面板的插件按钮被点击
    window.Asc.plugin.button = function (id, windowID) {
        console.log("on plugin button id=${id} ${windowID}");
        if (id == -1) {
            this.executeCommand("close", '');
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
        window.Asc.plugin.executeMethod("GetCurrentContentControl");
        window.Asc.plugin.onMethodReturn = function (returnValue) {
            if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {
                console.log("controlId", JSON.stringify(returnValue));

                if (returnValue) {
                    Asc.scope.controlId = returnValue;
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
            }
        };
    }


    window.Asc.plugin.event_onClick = function (isSelectionUse) {
        console.log("event click");
        showPosition(window, onGetPos);
    };
})(window, undefined);

