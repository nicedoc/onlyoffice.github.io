(function (window, undefined) {  
    let splitQuestion = function(text) {
        //const token = /(\d+．)|(（\d+）)/g;    
        const token = /(\d+．)/g;    
        var nextstart = 0;
        var next = text.substring(nextstart).search(token);
        var offset = 0;
        var ranges = new Array();
        while (next !== -1) {
            let elem = text.substring(nextstart, nextstart + offset + next);
            ranges.push({beg: nextstart, end: nextstart + offset + next})
            offset = 2;
            nextstart = nextstart + offset + next;
            next = text.substring(nextstart+offset).search(token);        
        }
    
        if (nextstart < text.length) {
            let elem = text.substring(nextstart);
            ranges.push({beg: nextstart, end: text.length});
        }
        
        ranges.shift();
        return ranges;
    }

    // 插件初始化    
    window.Asc.plugin.init = function () {
        console.log("biyue plugin inited.");
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

        document.getElementById("x-value").innerHTML = rect[0];
        document.getElementById("y-value").innerHTML = rect[1];
        document.getElementById("w-value").innerHTML = rect[2] - rect[0];
        document.getElementById("h-value").innerHTML = rect[3] - rect[1];
    };

    setCurrentContentControlLock = function(lock) {
        // 直接设置Lock以后，lock都不能操作了，锁定了不能通过插件操作？

        Asc.scope.lock = lock;
        window.Asc.plugin.callCommand(function()  {
            // 获取当前控件id
            const sContentControlId = Api.asc_GetCurrentContentControl();
            if (sContentControlId) {
                const prop = { Lock: Asc.scope.lock };                
                Api.asc_SetContentControlProperties(prop, sContentControlId, false);
            }                
        }, false, false, undefined);
    };

    createContentControl = function(ranges) {
        Asc.scope.ranges = ranges;
        window.Asc.plugin.callCommand(function() {
            var ranges = Asc.scope.ranges;

            for (var i = 0; i < ranges.length; i++) {                                
                // set selection
                var e = ranges[i];
                console.log(e);
                Api.GetDocument().GetRange(e.beg, e.end-1).Select()                
                Api.asc_AddContentControl(1);
                Api.asc_RemoveSelection();                                
            }
        }, false, false, undefined);
    }

    $(document).ready(function() {
        // 切题
        document.getElementById("splitQuestionBtn").onclick = function() {
            // get all text
            window.Asc.plugin.callCommand(function() {
                Api.asc_EditSelectAll();
                var text = Api.asc_GetSelectedText();
                
                Api.asc_RemoveSelection();
                return text;
            }, false, false, function(text) {                
                // split with token
                var ranges = splitQuestion(text);                
                
                createContentControl(ranges);
            });
            
            
        }

        // Todo 考虑其他实现方法
        // 锁定控件操作
        document.getElementById("unlockBtn").onclick = function() {
            /* Asc.c_oAscSdtLockType.Unlocked */
            setCurrentContentControlLock(3 );
        };

        document.getElementById("lockBtn").onclick = function() {
            // 1 Asc.c_oAscSdtLockType.SdtContentLocked 
            setCurrentContentControlLock(1);
        };

        document.getElementById("showPosition").onclick = function() {
            console.log("showPosition on button clicked");
            window.Asc.plugin.executeMethod("GetCurrentContentControl");
            window.Asc.plugin.onMethodReturn = function(returnValue) {                
                if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {                    
                    console.log("controlId", JSON.stringify(returnValue));                
                
                    if (returnValue) {
                        Asc.scope.controlId = returnValue;
                        window.Asc.plugin.callCommand(function()  {
                            var rect = Api.asc_GetContentControlBoundingRect(Asc.scope.controlId);                            
                            return rect;
                        }, false, false, onGetPos);
                        
                    }
                }
            };
            
        }
    });
        
    // 在editor面板的插件按钮被点击
    window.Asc.plugin.button = function(id, windowID) {
        console.log("on plugin button id=${id} ${windowID}");
        if (id == -1) {
            this.executeCommand("close", '');
            return;
        }      
   
    };

    window.Asc.plugin.onFocusContentControl = function(control)
    {
        console.log("onFocusControl");
        console.log(control);
    };

    window.Asc.plugin.attachEvent("onSelectionChanged", function(data) {
        console.log("on SelectionChange");
    });
})(window, undefined);

