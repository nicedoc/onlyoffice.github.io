(function (window, undefined) {
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

    $(document).ready(function() {
        document.getElementById("showPosition").onclick = function() {
            console.log("showPosition on button clicked");
            window.Asc.plugin.executeMethod("GetCurrentContentControl");
            window.Asc.plugin.onMethodReturn = function(returnValue) {                
                if (window.Asc.plugin.info.methodName == "GetCurrentContentControl") {                    
                    console.log(JSON.stringify(returnValue));
                
                
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

